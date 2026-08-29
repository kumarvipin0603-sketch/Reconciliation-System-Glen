from __future__ import annotations

from datetime import date, datetime
import hashlib
import re
from typing import Callable

import pandas as pd

OPEN_STATUSES = {
    "NEW", "UNDER_REVIEW", "ACTION_REQUIRED", "AWAITING_INTERNAL",
    "CLAIM_REQUIRED", "CLAIM_SUBMITTED", "AWAITING_AMAZON",
    "PARTIALLY_RESOLVED", "REIMBURSED", "ACCOUNTING_PENDING",
    "STOCK_FOUND", "WINDOW_EXPIRED", "WRITE_OFF_REVIEW",
}
CLOSED_STATUSES = {"AUTO_RESOLVED", "RESOLVED"}

STATUS_ALIASES = {
    "receving": "RECEIVING",
    "receiving": "RECEIVING",
    "claim window expired": "WINDOW_EXPIRED",
    "window closed": "WINDOW_EXPIRED",
    "reimbersed": "REIMBURSED",
    "reimbursed": "REIMBURSED",
    "all received": "RESOLVED",
}

RULES = [
    ("FBA-001", "SKU/FNSKU mapping missing", "Master Data", "High"),
    ("FBA-002", "FC/Plant mapping missing", "Master Data", "High"),
    ("FBA-003", "Inbound shipment short", "E-commerce / Amazon", "High"),
    ("FBA-004", "Inbound shipment excess", "Warehouse / E-commerce", "Medium"),
    ("FBA-005", "ERP transfer not linked to shipment", "Warehouse", "High"),
    ("FBA-006", "Amazon receipt without ERP transfer", "ERP / Warehouse", "High"),
    ("FBA-007", "FBA stock mismatch", "Reconciliation", "High"),
    ("FBA-008", "Amazon inventory lost/misplaced", "E-commerce / Amazon", "High"),
    ("FBA-009", "Customer return CN pending", "Accounts / ERP", "High"),
    ("FBA-010", "Return plant mismatch", "Reconciliation / ERP", "Medium"),
    ("FBA-011", "Reimbursement pending", "E-commerce / Amazon", "High"),
    ("FBA-012", "Reimbursement quantity mismatch", "Accounts / E-commerce", "High"),
    ("FBA-013", "Reimbursement value mismatch", "Accounts", "Medium"),
    ("FBA-014", "Reimbursement not billed/accounted", "Accounts", "High"),
    ("FBA-015", "Shipment claim action required", "E-commerce / Amazon", "Critical"),
    ("FBA-016", "Claim deadline approaching", "E-commerce / Amazon", "Critical"),
    ("FBA-017", "Claim window expired unresolved", "E-commerce Manager", "Critical"),
    ("FBA-018", "Duplicate source transaction", "System Admin", "Medium"),
    ("FBA-019", "Unknown/manual status", "Data Quality", "Medium"),
    ("FBA-020", "Unassigned discrepancy", "Reconciliation Manager", "High"),
]

DDL = [
    """CREATE TABLE IF NOT EXISTS fba_upload_batch(
        batch_id TEXT PRIMARY KEY, source_type TEXT NOT NULL, filename TEXT,
        file_hash TEXT, uploaded_at TEXT, uploaded_by TEXT, row_count INTEGER DEFAULT 0,
        status TEXT, error_message TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_sku_master(
        fnsku TEXT, asin TEXT, msku TEXT, erp_sku TEXT NOT NULL,
        active_from TEXT, active_to TEXT, updated_at TEXT,
        PRIMARY KEY(fnsku, msku)
    )""",
    """CREATE TABLE IF NOT EXISTS fba_fc_plant_master(
        fulfillment_center TEXT PRIMARY KEY, erp_plant TEXT NOT NULL,
        effective_from TEXT, effective_to TEXT, updated_at TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_shipment(
        shipment_id TEXT PRIMARY KEY, shipment_date TEXT, ship_to_fc TEXT,
        expected_qty REAL DEFAULT 0, located_qty REAL DEFAULT 0,
        amazon_status TEXT, transfer_no TEXT, transfer_qty REAL DEFAULT 0,
        remarks TEXT, last_updated TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_shipment_line(
        shipment_id TEXT NOT NULL, erp_sku TEXT NOT NULL,
        expected_qty REAL DEFAULT 0, located_qty REAL DEFAULT 0,
        transfer_qty REAL DEFAULT 0, difference_qty REAL DEFAULT 0,
        last_updated TEXT, PRIMARY KEY(shipment_id, erp_sku)
    )""",
    """CREATE TABLE IF NOT EXISTS fba_reimbursement(
        reimbursement_id TEXT PRIMARY KEY, approval_date TEXT, case_id TEXT,
        amazon_order_id TEXT, reason TEXT, sku TEXT, fnsku TEXT, erp_sku TEXT,
        amount_per_unit REAL DEFAULT 0, amount_total REAL DEFAULT 0,
        reimb_cash_qty REAL DEFAULT 0, reimb_inventory_qty REAL DEFAULT 0,
        reimb_total_qty REAL DEFAULT 0, accounting_status TEXT, remarks TEXT,
        last_updated TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_return(
        return_key TEXT PRIMARY KEY, return_date TEXT, amazon_order_id TEXT,
        sku TEXT, fnsku TEXT, erp_sku TEXT, qty REAL DEFAULT 0,
        fulfillment_center TEXT, disposition TEXT, reason TEXT,
        sale_doc TEXT, return_doc TEXT, cn_pending REAL DEFAULT 0,
        site_code TEXT, plant TEXT, last_updated TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_discrepancy_case(
        case_pk TEXT PRIMARY KEY, case_no TEXT UNIQUE NOT NULL, rule_id TEXT NOT NULL,
        discrepancy_type TEXT NOT NULL, plant TEXT, fulfillment_center TEXT,
        erp_sku TEXT, source_ref TEXT, source_date TEXT,
        discrepancy_qty REAL DEFAULT 0, open_qty REAL DEFAULT 0,
        expected_value REAL DEFAULT 0, recovered_value REAL DEFAULT 0,
        status TEXT NOT NULL, priority TEXT, owner_team TEXT, owner_user TEXT,
        due_date TEXT, claim_deadline TEXT, first_detected_at TEXT,
        last_seen_at TEXT, resolved_at TEXT, details TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_case_event(
        id INTEGER PRIMARY KEY AUTOINCREMENT, case_pk TEXT NOT NULL,
        event_time TEXT NOT NULL, old_status TEXT, new_status TEXT,
        action_type TEXT, note TEXT, user_id TEXT, source_batch_id TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_case_match(
        id INTEGER PRIMARY KEY AUTOINCREMENT, case_pk TEXT NOT NULL,
        match_type TEXT, source_table TEXT, source_record_id TEXT,
        matched_qty REAL DEFAULT 0, matched_value REAL DEFAULT 0,
        matched_at TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_rule_config(
        rule_id TEXT PRIMARY KEY, rule_name TEXT NOT NULL, enabled INTEGER DEFAULT 1,
        tolerance_qty REAL DEFAULT 0, tolerance_value REAL DEFAULT 0,
        grace_days INTEGER DEFAULT 0, warning_days INTEGER DEFAULT 7,
        default_owner TEXT, priority TEXT, updated_at TEXT
    )""",
]


def _txt(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _num(value) -> float:
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def normalize_status(value) -> str:
    raw = re.sub(r"\s+", " ", _txt(value)).strip()
    if not raw:
        return ""
    lower = raw.lower()
    if lower in STATUS_ALIASES:
        return STATUS_ALIASES[lower]
    if lower.startswith("case submitted"):
        return "CLAIM_SUBMITTED"
    return raw.upper().replace("-", "_").replace(" ", "_")


def init_fba_db(db: Callable):
    now = datetime.now().isoformat(timespec="seconds")
    with db() as c:
        for ddl in DDL:
            c.execute(ddl)
        c.executemany(
            """INSERT INTO fba_rule_config(
                rule_id,rule_name,enabled,tolerance_qty,tolerance_value,
                grace_days,warning_days,default_owner,priority,updated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(rule_id) DO UPDATE SET
                rule_name=excluded.rule_name,
                default_owner=excluded.default_owner,
                priority=excluded.priority
            """,
            [(rid, name, 1, 0, 0, 0, 7, owner, priority, now)
             for rid, name, owner, priority in RULES],
        )
        c.commit()


def make_case_key(rule_id: str, plant: str = "", sku: str = "", source_ref: str = "") -> str:
    raw = "|".join([_txt(rule_id), _txt(plant), _txt(sku), _txt(source_ref)])
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:24]


def make_case_no(case_pk: str) -> str:
    return f"FBA-{case_pk[:10].upper()}"


def upsert_case(db: Callable, *, rule_id: str, discrepancy_type: str,
                plant: str = "", fulfillment_center: str = "", erp_sku: str = "",
                source_ref: str = "", source_date: str = "", discrepancy_qty: float = 0,
                open_qty: float = 0, expected_value: float = 0, recovered_value: float = 0,
                status: str = "NEW", priority: str = "High", owner_team: str = "",
                due_date: str = "", claim_deadline: str = "", details: str = "",
                source_batch_id: str = "") -> str:
    case_pk = make_case_key(rule_id, plant, erp_sku, source_ref)
    case_no = make_case_no(case_pk)
    now = datetime.now().isoformat(timespec="seconds")
    status = normalize_status(status) or "NEW"
    resolved_at = now if status in CLOSED_STATUSES else None

    with db() as c:
        previous = c.execute(
            "SELECT status,open_qty FROM fba_discrepancy_case WHERE case_pk=?", (case_pk,)
        ).fetchone()
        old_status = _txt(previous[0]) if previous else ""

        c.execute("""
            INSERT INTO fba_discrepancy_case(
                case_pk,case_no,rule_id,discrepancy_type,plant,fulfillment_center,
                erp_sku,source_ref,source_date,discrepancy_qty,open_qty,
                expected_value,recovered_value,status,priority,owner_team,
                due_date,claim_deadline,first_detected_at,last_seen_at,resolved_at,details
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(case_pk) DO UPDATE SET
                discrepancy_qty=excluded.discrepancy_qty,
                open_qty=excluded.open_qty,
                expected_value=excluded.expected_value,
                recovered_value=excluded.recovered_value,
                status=excluded.status,
                priority=excluded.priority,
                owner_team=excluded.owner_team,
                due_date=excluded.due_date,
                claim_deadline=excluded.claim_deadline,
                last_seen_at=excluded.last_seen_at,
                resolved_at=excluded.resolved_at,
                details=excluded.details
        """, (
            case_pk, case_no, rule_id, discrepancy_type, _txt(plant),
            _txt(fulfillment_center), _txt(erp_sku), _txt(source_ref), _txt(source_date),
            _num(discrepancy_qty), _num(open_qty), _num(expected_value),
            _num(recovered_value), status, priority, owner_team, _txt(due_date),
            _txt(claim_deadline), now, now, resolved_at, _txt(details),
        ))

        if not previous or old_status != status:
            c.execute("""
                INSERT INTO fba_case_event(
                    case_pk,event_time,old_status,new_status,action_type,note,source_batch_id
                ) VALUES(?,?,?,?,?,?,?)
            """, (
                case_pk, now, old_status, status,
                "CASE_CREATED" if not previous else "STATUS_CHANGED",
                details, source_batch_id,
            ))
        c.commit()
    return case_pk


def detect_shipment_cases(db: Callable) -> int:
    """First production rule: expected/transfer vs Amazon located shipment quantity."""
    with db() as c:
        rows = c.execute("""
            SELECT shipment_id,ship_to_fc,expected_qty,located_qty,transfer_qty,
                   amazon_status,remarks,last_updated
            FROM fba_shipment
        """).fetchall()

    count = 0
    for shipment_id, fc, expected, located, transfer, amazon_status, remarks, updated in rows:
        expected = _num(expected)
        located = _num(located)
        transfer = _num(transfer)
        base = transfer if transfer > 0 else expected
        diff = base - located
        if diff > 0.000001:
            status = "WINDOW_EXPIRED" if normalize_status(remarks) == "WINDOW_EXPIRED" else "CLAIM_REQUIRED"
            rule_id = "FBA-017" if status == "WINDOW_EXPIRED" else "FBA-003"
            dtype = "Claim window expired unresolved" if status == "WINDOW_EXPIRED" else "Inbound shipment short"
            priority = "Critical" if status == "WINDOW_EXPIRED" else "High"
            upsert_case(
                db, rule_id=rule_id, discrepancy_type=dtype, fulfillment_center=fc,
                source_ref=shipment_id, discrepancy_qty=diff, open_qty=diff,
                status=status, priority=priority, owner_team="E-commerce / Amazon",
                details=f"Expected/transfer {base:g}; Amazon located {located:g}; shortage {diff:g}",
            )
            count += 1
        elif diff < -0.000001:
            upsert_case(
                db, rule_id="FBA-004", discrepancy_type="Inbound shipment excess",
                fulfillment_center=fc, source_ref=shipment_id,
                discrepancy_qty=abs(diff), open_qty=abs(diff), status="ACTION_REQUIRED",
                priority="Medium", owner_team="Warehouse / E-commerce",
                details=f"Expected/transfer {base:g}; Amazon located {located:g}; excess {abs(diff):g}",
            )
            count += 1
    return count


def auto_close_zero_cases(db: Callable) -> int:
    now = datetime.now().isoformat(timespec="seconds")
    with db() as c:
        rows = c.execute("""
            SELECT case_pk,status FROM fba_discrepancy_case
            WHERE ABS(COALESCE(open_qty,0)) <= 0.000001
              AND status NOT IN ('RESOLVED','AUTO_RESOLVED')
        """).fetchall()
        for case_pk, old_status in rows:
            c.execute("""
                UPDATE fba_discrepancy_case
                SET status='AUTO_RESOLVED', resolved_at=?, last_seen_at=?
                WHERE case_pk=?
            """, (now, now, case_pk))
            c.execute("""
                INSERT INTO fba_case_event(
                    case_pk,event_time,old_status,new_status,action_type,note
                ) VALUES(?,?,?,?,?,?)
            """, (case_pk, now, old_status, "AUTO_RESOLVED", "AUTO_CLOSE", "Open quantity reached zero"))
        c.commit()
    return len(rows)


def load_cases(db: Callable, include_closed: bool = False) -> pd.DataFrame:
    sql = "SELECT * FROM fba_discrepancy_case"
    if not include_closed:
        sql += " WHERE status NOT IN ('RESOLVED','AUTO_RESOLVED')"
    sql += " ORDER BY CASE priority WHEN 'Critical' THEN 1 WHEN 'High' THEN 2 WHEN 'Medium' THEN 3 ELSE 4 END, first_detected_at"
    with db() as c:
        return pd.read_sql_query(sql, c)


def case_summary(cases: pd.DataFrame) -> dict:
    if cases is None or cases.empty:
        return {"open_cases": 0, "open_qty": 0.0, "open_value": 0.0, "critical": 0, "claim_required": 0, "window_expired": 0}
    status = cases["status"].fillna("").astype(str)
    priority = cases["priority"].fillna("").astype(str)
    return {
        "open_cases": int(len(cases)),
        "open_qty": float(pd.to_numeric(cases["open_qty"], errors="coerce").fillna(0).sum()),
        "open_value": float(pd.to_numeric(cases["expected_value"], errors="coerce").fillna(0).sum()),
        "critical": int(priority.eq("Critical").sum()),
        "claim_required": int(status.eq("CLAIM_REQUIRED").sum()),
        "window_expired": int(status.eq("WINDOW_EXPIRED").sum()),
    }
