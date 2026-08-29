from __future__ import annotations

from datetime import datetime
from typing import Callable

import pandas as pd

import fba_reconciliation as fba

SYSTEM_RULES = ("FBA-001","FBA-003","FBA-004","FBA-009","FBA-012","FBA-014","FBA-017","FBA-019")


def _txt(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def _num(v):
    try:
        if v is None or pd.isna(v):
            return 0.0
        return float(v)
    except Exception:
        return 0.0


def rebuild_validated_cases(db: Callable) -> dict:
    """Rebuild only machine-generated cases using rules validated against the supplied production workbook.

    Key workbook finding: Shipment.DIFF is Units located minus transferred Qty. Therefore shipment
    discrepancy must use transfer_qty when present; Units expected must not be substituted when transfer_qty
    is zero because historical/unavailable shipments otherwise become false shortages.
    """
    with db() as c:
        placeholders = ",".join("?" for _ in SYSTEM_RULES)
        case_rows = c.execute(
            f"SELECT case_pk FROM fba_discrepancy_case WHERE rule_id IN ({placeholders})",
            SYSTEM_RULES,
        ).fetchall()
        case_pks = [_txt(r[0]) for r in case_rows]
        for case_pk in case_pks:
            c.execute("DELETE FROM fba_case_match WHERE case_pk=?", (case_pk,))
            c.execute("DELETE FROM fba_case_event WHERE case_pk=?", (case_pk,))
        c.execute(
            f"DELETE FROM fba_discrepancy_case WHERE rule_id IN ({placeholders})",
            SYSTEM_RULES,
        )
        c.commit()

    counts = {
        "shipment_short": 0,
        "shipment_excess": 0,
        "window_expired": 0,
        "status_conflict": 0,
        "cn_pending": 0,
        "mapping_missing": 0,
        "reimbursement_qty_issue": 0,
    }

    # Shipment rules validated against workbook DIFF = Units located - Qty.
    with db() as c:
        shipments = c.execute("""
            SELECT shipment_id,ship_to_fc,expected_qty,located_qty,transfer_qty,
                   amazon_status,remarks,last_updated
            FROM fba_shipment
        """).fetchall()

    for shipment_id, fc, expected, located, transfer, amazon_status, remarks, updated in shipments:
        transfer = _num(transfer)
        located = _num(located)
        expected = _num(expected)
        remark_status = fba.normalize_status(remarks)

        # No ERP transfer quantity = no reliable shipment discrepancy basis.
        if transfer <= 0:
            continue

        diff = located - transfer
        if diff < -0.000001:
            shortage = abs(diff)
            expired = remark_status == "WINDOW_EXPIRED"
            fba.upsert_case(
                db,
                rule_id="FBA-017" if expired else "FBA-003",
                discrepancy_type="Claim window expired unresolved" if expired else "Inbound shipment short",
                fulfillment_center=_txt(fc),
                source_ref=_txt(shipment_id),
                discrepancy_qty=shortage,
                open_qty=shortage,
                status="WINDOW_EXPIRED" if expired else ("CLAIM_SUBMITTED" if remark_status == "CLAIM_SUBMITTED" else "CLAIM_REQUIRED"),
                priority="Critical" if expired else "High",
                owner_team="E-commerce / Amazon",
                details=f"Transfer qty {transfer:g}; Amazon located {located:g}; shortage {shortage:g}; expected {expected:g}; workbook remark: {_txt(remarks) or '-'}",
            )
            counts["window_expired" if expired else "shipment_short"] += 1
        elif diff > 0.000001:
            excess = diff
            fba.upsert_case(
                db,
                rule_id="FBA-004",
                discrepancy_type="Inbound shipment excess",
                fulfillment_center=_txt(fc),
                source_ref=_txt(shipment_id),
                discrepancy_qty=excess,
                open_qty=excess,
                status="ACTION_REQUIRED",
                priority="Medium",
                owner_team="Warehouse / E-commerce",
                details=f"Transfer qty {transfer:g}; Amazon located {located:g}; excess {excess:g}; expected {expected:g}; workbook remark: {_txt(remarks) or '-'}",
            )
            counts["shipment_excess"] += 1

        # Contradictory manual remark: marked all received but quantity is not reconciled.
        if remark_status == "RESOLVED" and abs(diff) > 0.000001:
            fba.upsert_case(
                db,
                rule_id="FBA-019",
                discrepancy_type="Status conflicts with shipment quantity",
                fulfillment_center=_txt(fc),
                source_ref=_txt(shipment_id),
                discrepancy_qty=abs(diff),
                open_qty=abs(diff),
                status="UNDER_REVIEW",
                priority="High",
                owner_team="Reconciliation",
                details=f"Workbook says All received, but transfer qty {transfer:g} and located qty {located:g} differ by {diff:g}.",
            )
            counts["status_conflict"] += 1

    # Return/CN pending rule.
    with db() as c:
        returns = c.execute("""
            SELECT return_key,return_date,amazon_order_id,erp_sku,qty,fulfillment_center,
                   return_doc,cn_pending,plant
            FROM fba_return
        """).fetchall()
    for return_key, return_date, order_id, sku, qty, fc, return_doc, cn_pending, plant in returns:
        pending = _num(cn_pending)
        if pending > 0.000001 or (not _txt(return_doc) and _num(qty) > 0):
            open_qty = pending if pending > 0 else _num(qty)
            fba.upsert_case(
                db, rule_id="FBA-009", discrepancy_type="Customer return CN pending",
                plant=_txt(plant), fulfillment_center=_txt(fc), erp_sku=_txt(sku),
                source_ref=_txt(order_id) or _txt(return_key), source_date=_txt(return_date),
                discrepancy_qty=open_qty, open_qty=open_qty, status="ACTION_REQUIRED",
                priority="High", owner_team="Accounts / ERP",
                details=f"Amazon return qty {_num(qty):g}; CN pending {pending:g}; ERP return document: {_txt(return_doc) or 'missing'}.",
            )
            counts["cn_pending"] += 1

    # Missing SKU mapping on normalized Amazon events. Production workbook currently maps all events,
    # but keep this guard active for future uploads.
    with db() as c:
        missing = c.execute("""
            SELECT fnsku,msku,asin,COUNT(*),SUM(ABS(COALESCE(quantity,0)))
            FROM fba_amazon_inventory_event
            WHERE COALESCE(erp_sku,'')=''
            GROUP BY fnsku,msku,asin
        """).fetchall()
    for fnsku, msku, asin, row_count, qty in missing:
        ref = _txt(fnsku) or _txt(msku) or _txt(asin) or "UNKNOWN"
        fba.upsert_case(
            db, rule_id="FBA-001", discrepancy_type="SKU/FNSKU mapping missing",
            source_ref=ref, discrepancy_qty=_num(qty), open_qty=_num(qty),
            status="ACTION_REQUIRED", priority="High", owner_team="Master Data",
            details=f"No ERP SKU mapping for FNSKU {_txt(fnsku) or '-'}, MSKU {_txt(msku) or '-'}, ASIN {_txt(asin) or '-'} across {int(row_count or 0)} inventory event(s).",
        )
        counts["mapping_missing"] += 1

    # Reimbursement data-quality rule: a quantity-bearing reimbursement should have a mapped SKU.
    with db() as c:
        reimbursements = c.execute("""
            SELECT reimbursement_id,case_id,reason,erp_sku,reimb_total_qty,amount_total,remarks
            FROM fba_reimbursement
        """).fetchall()
    for rid, case_id, reason, sku, reimb_qty, amount, remarks in reimbursements:
        qty = _num(reimb_qty)
        if qty != 0 and not _txt(sku):
            fba.upsert_case(
                db, rule_id="FBA-012", discrepancy_type="Reimbursement SKU/quantity needs review",
                source_ref=_txt(rid), discrepancy_qty=abs(qty), open_qty=abs(qty),
                expected_value=abs(_num(amount)), status="UNDER_REVIEW", priority="High",
                owner_team="Accounts / E-commerce",
                details=f"Reimbursement {rid}; case {_txt(case_id) or '-'}; reason {_txt(reason) or '-'}; qty {qty:g}; amount {_num(amount):g}; ERP SKU mapping missing.",
            )
            counts["reimbursement_qty_issue"] += 1

    counts["total_cases"] = sum(counts.values())
    return counts
