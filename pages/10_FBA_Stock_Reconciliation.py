from __future__ import annotations

import hashlib
import io
import os
import sqlite3
import tempfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import psycopg2
import streamlit as st

import fba_importer
import fba_reconciliation as fba

st.set_page_config(page_title="FBA Stock Reconciliation", page_icon="📦", layout="wide")

DATA_DIR = Path.home() / "ReconciliationTowerData"
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "reconciliation_tower.db"


def get_database_url():
    try:
        value = st.secrets.get("DATABASE_URL", "")
        if value:
            return str(value).strip()
    except Exception:
        pass
    return os.getenv("DATABASE_URL", "").strip()


class PgResult:
    def __init__(self, cursor):
        self.cursor = cursor

    def fetchall(self):
        return self.cursor.fetchall()

    def fetchone(self):
        return self.cursor.fetchone()


class PgCompat:
    def __init__(self, url):
        self.conn = psycopg2.connect(
            url,
            connect_timeout=15,
            application_name="glen_fba_reconciliation",
        )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type is None:
                self.conn.commit()
            else:
                self.conn.rollback()
        finally:
            self.conn.close()
        return False

    @staticmethod
    def _ddl(sql):
        return sql.replace(
            "id INTEGER PRIMARY KEY AUTOINCREMENT",
            "id BIGSERIAL PRIMARY KEY",
        )

    def execute(self, sql, params=None):
        cur = self.conn.cursor()
        cur.execute(self._ddl(str(sql)).replace("?", "%s"), params or ())
        return PgResult(cur)

    def executemany(self, sql, seq):
        cur = self.conn.cursor()
        cur.executemany(
            self._ddl(str(sql)).replace("?", "%s"),
            list(seq),
        )
        return PgResult(cur)

    def cursor(self):
        return self.conn.cursor()

    def commit(self):
        self.conn.commit()


def db():
    url = get_database_url()
    if url:
        return PgCompat(url)
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def pin_hash(pin):
    return hashlib.sha256(str(pin).encode("utf-8")).hexdigest()


def admin_hash():
    try:
        with db() as c:
            row = c.execute(
                "SELECT setting_value FROM app_settings WHERE setting_key=?",
                ("admin_pin_hash",),
            ).fetchone()
        return str(row[0]).strip() if row and row[0] else ""
    except Exception:
        return ""


def admin_unlocked():
    return bool(st.session_state.get("fba_admin_unlocked", False))


def safe_df(sql, params=()):
    with db() as c:
        target = c.conn if isinstance(c, PgCompat) else c
        query = sql.replace("?", "%s") if get_database_url() else sql
        return pd.read_sql_query(query, target, params=params)


def excel_bytes(df: pd.DataFrame, sheet_name="FBA Cases"):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    out.seek(0)
    return out.getvalue()


def multi_sheet_excel_bytes(sheets: dict[str, pd.DataFrame]):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    out.seek(0)
    return out.getvalue()


def init_all():
    fba.init_fba_db(db)
    fba_importer.init_source_db(db)


def mail_action_text(row):
    status = str(row.get("Status", ""))
    discrepancy = str(row.get("Discrepancy", ""))
    if status == "WINDOW_EXPIRED":
        return "Review expired claim immediately and confirm recovery/write-off action."
    if status == "CLAIM_REQUIRED":
        return "Raise Amazon claim and share Case ID once submitted."
    if "short" in discrepancy.lower():
        return "Verify receipt shortage and raise/track Amazon claim."
    if "excess" in discrepancy.lower():
        return "Verify excess receipt against transfer and correct source/ERP quantity."
    if "reimbursement" in discrepancy.lower():
        return "Verify reimbursement quantity/value and confirm accounting entry."
    if "return" in discrepancy.lower():
        return "Verify return/CN posting and confirm closure reference."
    if "mapping" in discrepancy.lower():
        return "Correct mapping and confirm revised SKU/plant master."
    return "Review discrepancy and share closure evidence/reference."


# ------------------------------------------------------------------
# PRIVATE ACCESS GATE
# No FBA reconciliation data is initialized, queried or displayed
# until the main Reconciliation Tower Admin PIN is successfully entered.
# ------------------------------------------------------------------
st.title("FBA Stock Reconciliation Control Tower")
st.caption("Private management view — FBA discrepancy identification and mail action reporting.")

with st.sidebar:
    st.markdown("### Private Access")
    if admin_unlocked():
        st.success("Tower unlocked")
        if st.button("Lock Tower", use_container_width=True):
            st.session_state["fba_admin_unlocked"] = False
            st.rerun()
    else:
        stored_hash = admin_hash()
        pin = st.text_input("Admin PIN", type="password", key="fba_admin_pin")
        if st.button("Open FBA Tower", type="primary", use_container_width=True):
            if stored_hash and pin_hash(pin) == stored_hash:
                st.session_state["fba_admin_unlocked"] = True
                st.rerun()
            elif not stored_hash:
                st.error("Admin PIN is not configured in the main Reconciliation Tower.")
            else:
                st.error("Invalid Admin PIN")

if not admin_unlocked():
    st.info(
        "This FBA Control Tower is restricted to Admin. "
        "Team users will not see FBA reconciliation data or case details."
    )
    st.stop()

init_all()

with st.sidebar:
    st.divider()
    include_closed = st.checkbox("Include closed cases", value=False)

with st.expander("Upload & Run FBA Reconciliation", expanded=False):
    uploaded = st.file_uploader(
        "Upload complete FBA workbook",
        type=["xlsb", "xlsx", "xlsm"],
        accept_multiple_files=False,
        key="fba_source_upload",
    )
    st.caption(
        "Required sheets: ASIN vs SKU, MSD-Stock, Detailed-Stock, Stock In, Stock Out, "
        "FBA Transfer, FBA Return, Shipment and Reimbursement."
    )
    if uploaded and st.button(
        "Run FBA Reconciliation",
        type="primary",
        use_container_width=True,
    ):
        suffix = Path(uploaded.name).suffix.lower()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ("fba_source" + suffix)
            path.write_bytes(uploaded.getbuffer())
            with st.status("Processing FBA workbook…", expanded=True) as status:
                status.write("Validating sheets and source structure…")
                result = fba_importer.import_workbook(
                    path,
                    db,
                    uploaded_by="Private FBA Tower Admin",
                )
                if result.get("duplicate_file"):
                    status.update(
                        label="Workbook already imported",
                        state="complete",
                        expanded=False,
                    )
                    st.info(result.get("message", "This workbook already exists."))
                else:
                    status.write(f"Loaded {result['rows']:,} source rows.")
                    status.write(
                        f"Generated/updated {result['shipment_cases_generated']:,} "
                        "shipment discrepancy cases."
                    )
                    status.write(
                        f"Auto-closed {result['cases_auto_closed']:,} zero-balance cases."
                    )
                    status.update(
                        label="FBA reconciliation completed",
                        state="complete",
                        expanded=False,
                    )
                    st.success(
                        "FBA source snapshot saved and discrepancy engine completed successfully."
                    )
        st.cache_data.clear()
        st.rerun()

try:
    cases = fba.load_cases(db, include_closed=include_closed)
except Exception as exc:
    st.error(
        "FBA reconciliation data could not be loaded. "
        "Check the database connection and deployment logs."
    )
    st.exception(exc)
    st.stop()

open_cases = (
    cases[cases["status"].isin(list(fba.OPEN_STATUSES))].copy()
    if not cases.empty
    else cases
)
summary = fba.case_summary(open_cases)

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Open Cases", f"{summary['open_cases']:,}")
k2.metric("Open Qty", f"{summary['open_qty']:,.0f}")
k3.metric("Open Value", f"₹{summary['open_value']:,.2f}")
k4.metric("Critical", f"{summary['critical']:,}")
k5.metric("Claim Required", f"{summary['claim_required']:,}")
k6.metric("Window Expired", f"{summary['window_expired']:,}")

if cases.empty:
    st.info("No FBA discrepancy cases are available yet. Upload the complete FBA workbook above.")
    st.stop()

view = cases.copy()
for col in [
    "first_detected_at",
    "last_seen_at",
    "resolved_at",
    "source_date",
    "due_date",
    "claim_deadline",
]:
    if col in view.columns:
        view[col] = pd.to_datetime(view[col], errors="coerce")

view["aging_days"] = (
    pd.Timestamp(date.today())
    - pd.to_datetime(view["first_detected_at"], errors="coerce").dt.normalize()
).dt.days.fillna(0).clip(lower=0).astype(int)

st.markdown("### Management Filters")
f1, f2, f3, f4, f5 = st.columns(5)
status_options = sorted(
    [x for x in view["status"].dropna().astype(str).unique() if x]
)
priority_options = sorted(
    [x for x in view["priority"].dropna().astype(str).unique() if x]
)
team_options = sorted(
    [x for x in view["owner_team"].dropna().astype(str).unique() if x]
)
plant_options = sorted(
    [x for x in view["plant"].dropna().astype(str).unique() if x]
)
rule_options = sorted(
    [x for x in view["rule_id"].dropna().astype(str).unique() if x]
)

sel_status = f1.multiselect("Status", status_options)
sel_priority = f2.multiselect("Priority", priority_options)
sel_team = f3.multiselect("Send To Team", team_options)
sel_plant = f4.multiselect("Plant", plant_options)
sel_rule = f5.multiselect("Rule", rule_options)

if sel_status:
    view = view[view["status"].isin(sel_status)]
if sel_priority:
    view = view[view["priority"].isin(sel_priority)]
if sel_team:
    view = view[view["owner_team"].isin(sel_team)]
if sel_plant:
    view = view[view["plant"].isin(sel_plant)]
if sel_rule:
    view = view[view["rule_id"].isin(sel_rule)]

q1, q2, q3 = st.columns([2, 1, 1])
search = q1.text_input("Search Case / SKU / Shipment / FC")
min_age = q2.number_input("Minimum Aging Days", min_value=0, value=0, step=1)
critical_only = q3.checkbox("Critical only", value=False)

if search.strip():
    search_value = search.strip().lower()
    mask = pd.Series(False, index=view.index)
    for col in [
        "case_no",
        "erp_sku",
        "source_ref",
        "fulfillment_center",
        "details",
    ]:
        if col in view.columns:
            mask |= (
                view[col]
                .fillna("")
                .astype(str)
                .str.lower()
                .str.contains(search_value, regex=False)
            )
    view = view[mask]

if min_age:
    view = view[view["aging_days"] >= int(min_age)]
if critical_only:
    view = view[view["priority"].eq("Critical")]

priority_rank = {"Critical": 1, "High": 2, "Medium": 3, "Low": 4}
view["__priority_rank"] = view["priority"].map(priority_rank).fillna(9)
view = view.sort_values(
    ["__priority_rank", "aging_days", "first_detected_at"],
    ascending=[True, False, True],
)

show_cols = [
    "case_no",
    "priority",
    "status",
    "aging_days",
    "rule_id",
    "discrepancy_type",
    "plant",
    "fulfillment_center",
    "erp_sku",
    "source_ref",
    "discrepancy_qty",
    "open_qty",
    "expected_value",
    "recovered_value",
    "owner_team",
    "due_date",
    "claim_deadline",
    "details",
    "last_seen_at",
]
show_cols = [c for c in show_cols if c in view.columns]

display = view[show_cols].rename(
    columns={
        "case_no": "Case No",
        "priority": "Priority",
        "status": "Status",
        "aging_days": "Aging Days",
        "rule_id": "Rule",
        "discrepancy_type": "Discrepancy",
        "plant": "Plant",
        "fulfillment_center": "FC",
        "erp_sku": "ERP SKU",
        "source_ref": "Source Reference",
        "discrepancy_qty": "Original Qty",
        "open_qty": "Open Qty",
        "expected_value": "Expected Value",
        "recovered_value": "Recovered Value",
        "owner_team": "Send To Team",
        "due_date": "Due Date",
        "claim_deadline": "Claim Deadline",
        "details": "Details",
        "last_seen_at": "Last Seen",
    }
)

if not display.empty:
    display["Action Required"] = display.apply(mail_action_text, axis=1)

st.markdown("### Discrepancy Detail")
st.caption(
    f"Showing {len(display):,} case(s). This is your management working view; "
    "team members do not access the Tower."
)
st.dataframe(display, use_container_width=True, hide_index=True, height=520)

st.markdown("### Mail Action Report")
st.caption(
    "Use this report to send the required discrepancy details to the respective team by email."
)

mail_columns = [
    "Send To Team",
    "Priority",
    "Case No",
    "Aging Days",
    "Plant",
    "FC",
    "ERP SKU",
    "Source Reference",
    "Discrepancy",
    "Original Qty",
    "Open Qty",
    "Expected Value",
    "Status",
    "Action Required",
    "Claim Deadline",
    "Details",
]
mail_columns = [c for c in mail_columns if c in display.columns]
mail_report = display[mail_columns].copy()

if not mail_report.empty:
    team_summary = (
        mail_report.groupby("Send To Team", dropna=False)
        .agg(
            Cases=("Case No", "size"),
            Open_Qty=("Open Qty", "sum"),
            Open_Value=("Expected Value", "sum"),
            Critical=("Priority", lambda s: s.eq("Critical").sum()),
        )
        .reset_index()
        .rename(
            columns={
                "Open_Qty": "Open Qty",
                "Open_Value": "Open Value",
            }
        )
        .sort_values(["Critical", "Cases"], ascending=False)
    )
else:
    team_summary = pd.DataFrame(
        columns=["Send To Team", "Cases", "Open Qty", "Open Value", "Critical"]
    )

m1, m2 = st.columns([1, 2])
with m1:
    st.markdown("#### Team Summary")
    st.dataframe(team_summary.round(2), use_container_width=True, hide_index=True)
with m2:
    st.markdown("#### Mail Detail")
    st.dataframe(mail_report, use_container_width=True, hide_index=True, height=300)

export_sheets = {
    "Team Summary": team_summary,
    "Mail Action Detail": mail_report,
    "Full Case Detail": display,
}

d1, d2 = st.columns(2)
d1.download_button(
    "Download Mail Action Report",
    data=multi_sheet_excel_bytes(export_sheets),
    file_name=f"FBA_Mail_Action_Report_{date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
d2.download_button(
    "Download Mail Detail CSV",
    data=mail_report.to_csv(index=False).encode("utf-8-sig"),
    file_name=f"FBA_Mail_Action_Detail_{date.today().isoformat()}.csv",
    mime="text/csv",
    use_container_width=True,
)

st.divider()
tab_mix, tab_age, tab_team, tab_history = st.tabs(
    ["Exception Mix", "Aging", "Team-wise Action", "Case History"]
)

with tab_mix:
    mix = (
        view.groupby(["rule_id", "discrepancy_type"], dropna=False)
        .agg(
            Cases=("case_pk", "size"),
            Open_Qty=("open_qty", "sum"),
            Open_Value=("expected_value", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "rule_id": "Rule",
                "discrepancy_type": "Discrepancy",
                "Open_Qty": "Open Qty",
                "Open_Value": "Open Value",
            }
        )
        .sort_values(["Cases", "Open Qty"], ascending=False)
    )
    st.dataframe(mix.round(2), use_container_width=True, hide_index=True)
    if not mix.empty:
        st.bar_chart(mix.set_index("Discrepancy")["Cases"])

with tab_age:
    labels = ["0–7", "8–15", "16–30", "31–60", "61+"]
    age = view.copy()
    age["Aging Bucket"] = pd.cut(
        age["aging_days"],
        bins=[-1, 7, 15, 30, 60, 10**9],
        labels=labels,
    )
    aging = (
        age["Aging Bucket"]
        .value_counts(sort=False)
        .reindex(labels, fill_value=0)
        .to_frame("Cases")
    )
    st.bar_chart(aging)
    st.dataframe(aging.reset_index(), use_container_width=True, hide_index=True)

with tab_team:
    st.caption("Use this only to decide which team should receive your email action list.")
    st.dataframe(team_summary.round(2), use_container_width=True, hide_index=True)

with tab_history:
    selected_case = st.selectbox(
        "Select Case",
        options=display["Case No"].tolist() if not display.empty else [],
    )
    if selected_case:
        row = cases[cases["case_no"].eq(selected_case)].iloc[0]
        history = safe_df(
            """SELECT event_time,old_status,new_status,action_type,note,user_id,source_batch_id
               FROM fba_case_event WHERE case_pk=? ORDER BY event_time DESC""",
            (row["case_pk"],),
        )
        st.dataframe(history, use_container_width=True, hide_index=True)

st.divider()
st.markdown("### Management Case Update")
update_case_no = st.selectbox(
    "Case to update",
    options=display["Case No"].tolist() if not display.empty else [],
    key="case_update_no",
)

if update_case_no:
    current = cases[cases["case_no"].eq(update_case_no)].iloc[0]
    status_values = [
        "NEW",
        "UNDER_REVIEW",
        "ACTION_REQUIRED",
        "AWAITING_INTERNAL",
        "CLAIM_REQUIRED",
        "CLAIM_SUBMITTED",
        "AWAITING_AMAZON",
        "PARTIALLY_RESOLVED",
        "REIMBURSED",
        "ACCOUNTING_PENDING",
        "STOCK_FOUND",
        "WINDOW_EXPIRED",
        "WRITE_OFF_REVIEW",
        "RESOLVED",
    ]
    current_status = str(current.get("status") or "NEW")
    current_index = (
        status_values.index(current_status)
        if current_status in status_values
        else 0
    )

    u1, u2 = st.columns(2)
    new_status = u1.selectbox(
        "Status",
        status_values,
        index=current_index,
    )
    open_qty = u2.number_input(
        "Open Qty",
        min_value=0.0,
        value=float(current.get("open_qty") or 0),
        step=1.0,
    )
    note = st.text_area(
        "Management Note / Team Feedback Received",
        placeholder="Paste the team response, Amazon Case ID, ERP correction reference, reimbursement confirmation, etc.",
    )

    if st.button("Save Management Update", type="primary"):
        now = datetime.now().isoformat(timespec="seconds")
        resolved_at = now if new_status == "RESOLVED" else None
        with db() as c:
            c.execute(
                """
                UPDATE fba_discrepancy_case
                SET status=?,open_qty=?,last_seen_at=?,resolved_at=?
                WHERE case_pk=?
                """,
                (
                    new_status,
                    float(open_qty),
                    now,
                    resolved_at,
                    current["case_pk"],
                ),
            )
            c.execute(
                """
                INSERT INTO fba_case_event(
                    case_pk,event_time,old_status,new_status,action_type,note,user_id
                ) VALUES(?,?,?,?,?,?,?)
                """,
                (
                    current["case_pk"],
                    now,
                    str(current["status"]),
                    new_status,
                    "MANAGEMENT_UPDATE",
                    note.strip(),
                    "Private Tower Admin",
                ),
            )
            c.commit()

        fba.auto_close_zero_cases(db)
        st.success("Management case update saved successfully.")
        st.rerun()
