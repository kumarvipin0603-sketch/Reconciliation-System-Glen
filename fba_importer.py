from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
import hashlib
import math
import re
from typing import Callable

import pandas as pd

import fba_reconciliation as fba

SOURCE_SHEETS = [
    "ASIN vs SKU", "MSD-Stock", "Detailed-Stock", "Stock In", "Stock Out",
    "FBA Transfer", "FBA Return", "Shipment", "Reimbursement",
]

SOURCE_DDL = [
    """CREATE TABLE IF NOT EXISTS fba_amazon_inventory_event(
        row_hash TEXT PRIMARY KEY, batch_id TEXT, fnsku TEXT, asin TEXT, msku TEXT,
        event_type TEXT, reference_id TEXT, quantity REAL DEFAULT 0,
        fulfillment_center TEXT, plant TEXT, disposition TEXT,
        reconciled_quantity REAL DEFAULT 0, unreconciled_quantity REAL DEFAULT 0,
        event_date TEXT, erp_sku TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_erp_stock_ledger(
        row_hash TEXT PRIMARY KEY, batch_id TEXT, posting_date TEXT, entry_type TEXT,
        document_type TEXT, document_no TEXT, erp_sku TEXT, description TEXT,
        branch_code TEXT, department_code TEXT, location_code TEXT,
        remaining_quantity REAL DEFAULT 0, quantity REAL DEFAULT 0,
        invoiced_quantity REAL DEFAULT 0, sales_amount REAL DEFAULT 0,
        cost_amount REAL DEFAULT 0, is_open TEXT, entry_no TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_stock_in(
        row_hash TEXT PRIMARY KEY, batch_id TEXT, order_no TEXT, order_date TEXT,
        invoice_no TEXT, invoice_date TEXT, po_number TEXT, erp_sku TEXT,
        quantity REAL DEFAULT 0, unit_price REAL DEFAULT 0, line_amount REAL DEFAULT 0,
        gross_amount REAL DEFAULT 0, branch_code TEXT, location_code TEXT, plant TEXT,
        remarks TEXT, shipment_id TEXT, located_qty REAL DEFAULT 0,
        difference_qty REAL DEFAULT 0, mapped_sku TEXT, mapped_qty REAL DEFAULT 0,
        source_status TEXT, last_updated TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_stock_out(
        row_hash TEXT PRIMARY KEY, batch_id TEXT, order_no TEXT, order_date TEXT,
        invoice_no TEXT, invoice_date TEXT, po_number TEXT, erp_sku TEXT,
        quantity REAL DEFAULT 0, unit_price REAL DEFAULT 0, line_amount REAL DEFAULT 0,
        gross_amount REAL DEFAULT 0, branch_code TEXT, location_code TEXT, plant TEXT,
        remarks TEXT, category TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS fba_transfer(
        row_hash TEXT PRIMARY KEY, batch_id TEXT, transaction_type TEXT,
        transaction_id TEXT, order_id TEXT, ship_from_fc TEXT, ship_to_fc TEXT,
        invoice_no TEXT, invoice_date TEXT, invoice_value REAL DEFAULT 0,
        asin TEXT, sku TEXT, quantity REAL DEFAULT 0, taxable_value REAL DEFAULT 0,
        ship_from_plant TEXT, transfer_no TEXT
    )""",
]


def _txt(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    if s.endswith(".0") and re.fullmatch(r"-?\d+\.0", s):
        s = s[:-2]
    return s


def _num(v) -> float:
    try:
        if v is None or pd.isna(v):
            return 0.0
        return float(v)
    except Exception:
        return 0.0


def _date(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        # Excel serial-date convention (1900 date system).
        if 20000 <= float(v) <= 80000:
            dt = datetime(1899, 12, 30) + timedelta(days=float(v))
            return dt.date().isoformat()
    dt = pd.to_datetime(v, errors="coerce", utc=True)
    if pd.isna(dt):
        return _txt(v)
    try:
        return dt.date().isoformat()
    except Exception:
        return str(dt)[:10]


def _datetime_text(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 <= float(v) <= 80000:
            dt = datetime(1899, 12, 30) + timedelta(days=float(v))
            return dt.isoformat(timespec="seconds")
    dt = pd.to_datetime(v, errors="coerce", utc=True)
    if pd.isna(dt):
        return _txt(v)
    try:
        return dt.isoformat()
    except Exception:
        return _txt(v)


def _hash(*parts) -> str:
    raw = "|".join(_txt(x) for x in parts)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def _file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [re.sub(r"\s+", " ", _txt(c)).strip() for c in out.columns]
    return out


def _excel_engine(path: Path) -> str | None:
    suffix = path.suffix.lower()
    if suffix == ".xlsb":
        return "pyxlsb"
    if suffix in {".xlsx", ".xlsm"}:
        return "openpyxl"
    return None


def read_fba_workbook(path) -> dict[str, pd.DataFrame]:
    path = Path(path)
    engine = _excel_engine(path)
    if not engine:
        raise ValueError("FBA workbook must be .xlsb, .xlsx or .xlsm")
    xls = pd.ExcelFile(path, engine=engine)
    present = set(xls.sheet_names)
    missing = [s for s in SOURCE_SHEETS if s not in present]
    if missing:
        raise ValueError("FBA workbook missing required sheet(s): " + ", ".join(missing))
    return {
        sheet: _clean_columns(pd.read_excel(path, sheet_name=sheet, engine=engine))
        for sheet in SOURCE_SHEETS
    }


def init_source_db(db: Callable):
    with db() as c:
        for ddl in SOURCE_DDL:
            c.execute(ddl)
        c.commit()


def _chunked(rows, size=5000):
    for i in range(0, len(rows), size):
        yield rows[i:i + size]


def _replace_table(c, table: str):
    c.execute(f"DELETE FROM {table}")


def _mapping_dict(master: pd.DataFrame) -> dict[str, str]:
    mapping = {}
    for _, r in master.iterrows():
        erp = _txt(r.get("Correct SKU")).upper()
        if not erp:
            continue
        for key in (_txt(r.get("FNSKU")), _txt(r.get("MSKU")), _txt(r.get("ASIN"))):
            if key:
                mapping[key.upper()] = erp
    return mapping


def _erp_sku(row, mapping: dict[str, str], direct="Correct SKU") -> str:
    direct_value = _txt(row.get(direct)).upper()
    if direct_value:
        return direct_value
    for col in ("FNSKU", "fnsku", "MSKU", "sku", "Sku", "ASIN", "asin", "Asin"):
        key = _txt(row.get(col)).upper()
        if key and key in mapping:
            return mapping[key]
    for col in ("Item", "SKU", "Product/Item No", "Item No."):
        key = _txt(row.get(col)).upper()
        if key:
            return key
    return ""


def import_workbook(path, db: Callable, uploaded_by: str = "") -> dict:
    """Replace current FBA source snapshot, preserve case history, then run rules."""
    path = Path(path)
    fba.init_fba_db(db)
    init_source_db(db)

    file_hash = _file_hash(path)
    batch_id = "FBA-" + file_hash[:20].upper()
    now = datetime.now().isoformat(timespec="seconds")

    with db() as c:
        existing = c.execute(
            "SELECT status,row_count FROM fba_upload_batch WHERE batch_id=?", (batch_id,)
        ).fetchone()
    if existing and _txt(existing[0]).upper() == "COMPLETED":
        return {
            "batch_id": batch_id,
            "duplicate_file": True,
            "rows": int(existing[1] or 0),
            "message": "This exact FBA workbook has already been imported.",
        }

    sheets = read_fba_workbook(path)
    mapping = _mapping_dict(sheets["ASIN vs SKU"])
    total_rows = sum(len(df) for df in sheets.values())

    with db() as c:
        c.execute("""
            INSERT INTO fba_upload_batch(
                batch_id,source_type,filename,file_hash,uploaded_at,uploaded_by,row_count,status,error_message
            ) VALUES(?,?,?,?,?,?,?,?,?)
            ON CONFLICT(batch_id) DO UPDATE SET
                uploaded_at=excluded.uploaded_at, uploaded_by=excluded.uploaded_by,
                row_count=excluded.row_count, status=excluded.status, error_message=excluded.error_message
        """, (batch_id, "FBA_WORKBOOK", path.name, file_hash, now, uploaded_by, total_rows, "PROCESSING", ""))
        c.commit()

    try:
        with db() as c:
            # Current workbook represents the latest complete FBA snapshot.
            for table in [
                "fba_sku_master", "fba_amazon_inventory_event", "fba_erp_stock_ledger",
                "fba_stock_in", "fba_stock_out", "fba_transfer", "fba_return",
                "fba_shipment", "fba_reimbursement",
            ]:
                _replace_table(c, table)

            # SKU master
            sku_rows = []
            for _, r in sheets["ASIN vs SKU"].iterrows():
                erp = _txt(r.get("Correct SKU")).upper()
                if not erp:
                    continue
                sku_rows.append((
                    _txt(r.get("FNSKU")).upper(), _txt(r.get("ASIN")).upper(),
                    _txt(r.get("MSKU")).upper(), erp, "", "", now,
                ))
            if sku_rows:
                c.executemany("""
                    INSERT INTO fba_sku_master(fnsku,asin,msku,erp_sku,active_from,active_to,updated_at)
                    VALUES(?,?,?,?,?,?,?)
                    ON CONFLICT(fnsku,msku) DO UPDATE SET asin=excluded.asin,erp_sku=excluded.erp_sku,updated_at=excluded.updated_at
                """, sku_rows)

            # Detailed Amazon inventory events
            rows = []
            for i, r in sheets["Detailed-Stock"].iterrows():
                fnsku, asin, msku = _txt(r.get("FNSKU")), _txt(r.get("ASIN")), _txt(r.get("MSKU"))
                event_type, ref = _txt(r.get("Event Type")), _txt(r.get("Reference ID"))
                fc, plant = _txt(r.get("Fulfillment Center")), _txt(r.get("Plant"))
                disposition, event_date = _txt(r.get("Disposition")), _date(r.get("Date"))
                sku = _erp_sku(r, mapping)
                key = _hash(fnsku, asin, msku, event_type, ref, _num(r.get("Quantity")), fc, disposition, event_date, i)
                rows.append((
                    key,batch_id,fnsku,asin,msku,event_type,ref,_num(r.get("Quantity")),fc,plant,
                    disposition,_num(r.get("Reconciled Quantity")),_num(r.get("Unreconciled Quantity")),event_date,sku,
                ))
            for chunk in _chunked(rows):
                c.executemany("""INSERT INTO fba_amazon_inventory_event(
                    row_hash,batch_id,fnsku,asin,msku,event_type,reference_id,quantity,fulfillment_center,plant,
                    disposition,reconciled_quantity,unreconciled_quantity,event_date,erp_sku
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", chunk)

            # ERP MSD stock ledger
            rows = []
            for i, r in sheets["MSD-Stock"].iterrows():
                doc = _txt(r.get("Document No.")); sku = _txt(r.get("Item No.")).upper(); entry = _txt(r.get("Entry No."))
                key = _hash(entry, doc, sku, _date(r.get("Posting Date")), i)
                rows.append((
                    key,batch_id,_date(r.get("Posting Date")),_txt(r.get("Entry Type")),_txt(r.get("Document Type")),
                    doc,sku,_txt(r.get("Description")),_txt(r.get("Branch Code")),_txt(r.get("Department Code")),
                    _txt(r.get("Location Code")),_num(r.get("Remaining Quantity")),_num(r.get("Quantity")),
                    _num(r.get("Invoiced Quantity")),_num(r.get("Sales Amount (Actual)")),_num(r.get("Cost Amount (Actual)")),
                    _txt(r.get("Open")),entry,
                ))
            for chunk in _chunked(rows):
                c.executemany("""INSERT INTO fba_erp_stock_ledger(
                    row_hash,batch_id,posting_date,entry_type,document_type,document_no,erp_sku,description,
                    branch_code,department_code,location_code,remaining_quantity,quantity,invoiced_quantity,
                    sales_amount,cost_amount,is_open,entry_no
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", chunk)

            # ERP stock in
            rows = []
            for i, r in sheets["Stock In"].iterrows():
                inv=_txt(r.get("Invoice No")); sku=_txt(r.get("Product/Item No")).upper(); shipment=_txt(r.get("Shipment ID"))
                key=_hash(inv,sku,_num(r.get("Quantity")),shipment,i)
                rows.append((
                    key,batch_id,_txt(r.get("Sales Order No.")),_date(r.get("Order Date")),inv,_date(r.get("Invoice Date")),
                    _txt(r.get("Po Number")),sku,_num(r.get("Quantity")),_num(r.get("Unit Price")),_num(r.get("Line Amount")),
                    _num(r.get("Gross Amount")),_txt(r.get("Branch Code")),_txt(r.get("Location Code")),_txt(r.get("Plant")),
                    _txt(r.get("Remarks")),shipment,_num(r.get("Located")),_num(r.get("Difference")),
                    _txt(r.get("SKU")).upper(),_num(r.get("QTY")),fba.normalize_status(r.get("Status")),_datetime_text(r.get("Last Updated")),
                ))
            for chunk in _chunked(rows):
                c.executemany("""INSERT INTO fba_stock_in(
                    row_hash,batch_id,order_no,order_date,invoice_no,invoice_date,po_number,erp_sku,quantity,unit_price,
                    line_amount,gross_amount,branch_code,location_code,plant,remarks,shipment_id,located_qty,difference_qty,
                    mapped_sku,mapped_qty,source_status,last_updated
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", chunk)

            # ERP stock out
            rows=[]
            for i,r in sheets["Stock Out"].iterrows():
                inv=_txt(r.get("Invoice No")); sku=_txt(r.get("Product/Item No")).upper(); po=_txt(r.get("Po Number"))
                key=_hash(inv,sku,_num(r.get("Quantity")),po,i)
                rows.append((
                    key,batch_id,_txt(r.get("Sales Order No.")),_date(r.get("Order Date")),inv,_date(r.get("Invoice Date")),po,
                    sku,_num(r.get("Quantity")),_num(r.get("Unit Price")),_num(r.get("Line Amount")),_num(r.get("Gross Amount")),
                    _txt(r.get("Branch Code")),_txt(r.get("Location Code")),_txt(r.get("Plant")),_txt(r.get("Remarks")),_txt(r.get("Catogary")),
                ))
            for chunk in _chunked(rows):
                c.executemany("""INSERT INTO fba_stock_out(
                    row_hash,batch_id,order_no,order_date,invoice_no,invoice_date,po_number,erp_sku,quantity,unit_price,
                    line_amount,gross_amount,branch_code,location_code,plant,remarks,category
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", chunk)

            # FBA transfer/removal
            rows=[]
            for i,r in sheets["FBA Transfer"].iterrows():
                tx=_txt(r.get("Transaction Id")); sku=_txt(r.get("Sku")).upper(); qty=_num(r.get("Quantity"))
                key=_hash(tx,sku,qty,i)
                rows.append((
                    key,batch_id,_txt(r.get("Transaction Type")),tx,_txt(r.get("Order Id")),_txt(r.get("Ship From Fc")),
                    _txt(r.get("Ship To Fc")),_txt(r.get("Invoice Number")),_date(r.get("Invoice Date")),_num(r.get("Invoice Value")),
                    _txt(r.get("Asin")),sku,qty,_num(r.get("Taxable Value")),_txt(r.get("Ship From")),_txt(r.get("Transfer No")),
                ))
            for chunk in _chunked(rows):
                c.executemany("""INSERT INTO fba_transfer(
                    row_hash,batch_id,transaction_type,transaction_id,order_id,ship_from_fc,ship_to_fc,invoice_no,invoice_date,
                    invoice_value,asin,sku,quantity,taxable_value,ship_from_plant,transfer_no
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", chunk)

            # Customer returns
            rows=[]
            for i,r in sheets["FBA Return"].iterrows():
                order=_txt(r.get("order-id")); fnsku=_txt(r.get("fnsku")); ret_date=_date(r.get("return-date")); qty=_num(r.get("quantity"))
                key=_hash(order,fnsku,ret_date,_txt(r.get("license-plate-number")),i)
                rows.append((
                    key,ret_date,order,_txt(r.get("sku")),fnsku,_erp_sku(r,mapping,direct="Item"),qty,
                    _txt(r.get("fulfillment-center-id")),_txt(r.get("detailed-disposition")),_txt(r.get("reason")),
                    _txt(r.get("Sale")),_txt(r.get("Return")),_num(r.get("CN Pending")),_txt(r.get("Site Code")),_txt(r.get("Plant")),now,
                ))
            for chunk in _chunked(rows):
                c.executemany("""INSERT INTO fba_return(
                    return_key,return_date,amazon_order_id,sku,fnsku,erp_sku,qty,fulfillment_center,disposition,reason,
                    sale_doc,return_doc,cn_pending,site_code,plant,last_updated
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", chunk)

            # Shipment headers
            rows=[]
            for _,r in sheets["Shipment"].iterrows():
                shipment=_txt(r.get("Shipment ID"))
                if not shipment: continue
                rows.append((
                    shipment,_date(r.get("Shipment Date")),_txt(r.get("Ship to")),_num(r.get("Units expected")),
                    _num(r.get("Units located")),fba.normalize_status(r.get("Status")),_txt(r.get("Transfer")),_num(r.get("Qty")),
                    _txt(r.get("Remarks")),_datetime_text(r.get("Last updated")),
                ))
            if rows:
                c.executemany("""INSERT INTO fba_shipment(
                    shipment_id,shipment_date,ship_to_fc,expected_qty,located_qty,amazon_status,transfer_no,transfer_qty,remarks,last_updated
                ) VALUES(?,?,?,?,?,?,?,?,?,?)""", rows)

            # Reimbursements
            rows=[]
            for i,r in sheets["Reimbursement"].iterrows():
                rid=_txt(r.get("reimbursement-id"))
                if not rid: rid=_hash(_txt(r.get("case-id")),_txt(r.get("amazon-order-id")),_num(r.get("amount-total")),i)[:20]
                rows.append((
                    rid,_date(r.get("approval-date")),_txt(r.get("case-id")),_txt(r.get("amazon-order-id")),_txt(r.get("reason")),
                    _txt(r.get("sku")),_txt(r.get("fnsku")),_erp_sku(r,mapping,direct="SKU"),_num(r.get("amount-per-unit")),
                    _num(r.get("amount-total")),_num(r.get("quantity-reimbursed-cash")),_num(r.get("quantity-reimbursed-inventory")),
                    _num(r.get("quantity-reimbursed-total")),"",_txt(r.get("Remarks")),now,
                ))
            if rows:
                c.executemany("""INSERT INTO fba_reimbursement(
                    reimbursement_id,approval_date,case_id,amazon_order_id,reason,sku,fnsku,erp_sku,amount_per_unit,amount_total,
                    reimb_cash_qty,reimb_inventory_qty,reimb_total_qty,accounting_status,remarks,last_updated
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", rows)

            c.execute("""UPDATE fba_upload_batch SET status='COMPLETED',error_message='' WHERE batch_id=?""", (batch_id,))
            c.commit()

        generated = fba.detect_shipment_cases(db)
        closed = fba.auto_close_zero_cases(db)
        return {
            "batch_id": batch_id,
            "duplicate_file": False,
            "rows": total_rows,
            "sheet_rows": {name: int(len(df)) for name, df in sheets.items()},
            "shipment_cases_generated": int(generated),
            "cases_auto_closed": int(closed),
        }
    except Exception as exc:
        with db() as c:
            c.execute("UPDATE fba_upload_batch SET status='FAILED',error_message=? WHERE batch_id=?", (str(exc)[:1000], batch_id))
            c.commit()
        raise
