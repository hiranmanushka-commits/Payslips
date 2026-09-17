from __future__ import annotations

import base64
import hashlib
import html
import json
import os
import re
import secrets
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

APP_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = os.path.join(os.path.dirname(__file__), "Salary_Slip 2026-2027.xlsx")
AUTH_STORE = APP_DIR / ".employee_auth.json"
PIN_MIN_LENGTH = 4

st.set_page_config(page_title="Employee Payslip Portal", page_icon="💼", layout="wide")


def image_data(path: Path) -> str:
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode("ascii")


def clean_month(value: object) -> str:
    if pd.isna(value):
        return ""
    parsed = pd.to_datetime(value, errors="coerce")
    return parsed.strftime("%B-%Y") if pd.notna(parsed) else str(value).strip()


def column_value(row: pd.Series, name: str, default: object = 0.0) -> object:
    for column in row.index:
        if str(column).strip().casefold() == name.casefold():
            value = row[column]
            return default if pd.isna(value) else value
    return default


def money(value: object) -> str:
    try:
        number = float(value)
        return "" if pd.isna(number) or number == 0 else f"{number:,.2f}"
    except (TypeError, ValueError):
        return str(value) if value else ""


def safe(value: object) -> str:
    return html.escape(str(value)) if value is not None else ""


def row_html(label: str, value: object) -> str:
    formatted = money(value)
    if not formatted:
        return ""
    return f"<tr><td>{safe(label)}</td><td class='amount'>{formatted}</td></tr>"


def installment_suffix(row: pd.Series) -> str:
    installments = []
    for column in row.index:
        if "loan instalement" in str(column).casefold():
            value = row[column]
            text = str(value).strip() if pd.notna(value) else ""
            if re.fullmatch(r"\d+(?:\s*/\s*\d+)?", text):
                installments.append(text)
    return f" ({', '.join(installments)})" if installments else ""


def load_workbook(uploaded_file):
    try:
        if uploaded_file is None:
            if not os.path.isfile(DEFAULT_WORKBOOK):
                st.error("Workbook not found: Salary_Slip 2026-2027.xlsx")
                return None
            source = DEFAULT_WORKBOOK
        else:
            source = uploaded_file
        return (
            pd.read_excel(source, sheet_name="Salary Data"),
            pd.read_excel(source, sheet_name="Employee Master"),
        )
    except Exception as error:
        st.error(f"Could not read the workbook: {error}")
        return None


def normalized_identity(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().casefold()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    normalized = re.sub(r"[^a-z0-9]", "", text)
    return normalized.lstrip("0") or "0" if normalized.isdigit() else normalized


def load_auth_store() -> dict:
    try:
        return json.loads(AUTH_STORE.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_auth_store(auth_store: dict) -> None:
    temporary_path = AUTH_STORE.with_suffix(".tmp")
    temporary_path.write_text(json.dumps(auth_store, indent=2), encoding="utf-8")
    temporary_path.replace(AUTH_STORE)


def pin_hash(pin: str, salt: bytes | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", pin.encode("utf-8"), salt, 200_000)
    return salt.hex(), digest.hex()


def pin_matches(pin: str, stored: dict) -> bool:
    try:
        _, digest = pin_hash(pin, bytes.fromhex(stored["salt"]))
        return secrets.compare_digest(digest, stored["hash"])
    except (KeyError, ValueError, TypeError):
        return False


def employee_record(name: str, employee_master: pd.DataFrame) -> pd.Series | None:
    if "Employee Name" not in employee_master.columns or "EPF No" not in employee_master.columns:
        return None
    submitted_name = normalized_identity(name)
    if not submitted_name:
        return None
    for _, employee in employee_master.iterrows():
        stored_name = str(employee.get("Employee Name", "")).strip()
        if secrets.compare_digest(normalized_identity(stored_name), submitted_name):
            return employee
    return None


def authenticate_employee(name: str, pin: str, auth_store: dict) -> bool:
    record = auth_store.get(normalized_identity(name))
    return bool(record and pin_matches(pin, record))


def valid_pin(pin: str) -> bool:
    return bool(re.fullmatch(r"\d{4,8}", pin))


def configured_admin() -> tuple[str, str]:
    username = os.environ.get("SUPER_ADMIN_USERNAME", "")
    password = os.environ.get("SUPER_ADMIN_PASSWORD", "")
    if username and password:
        return username, password
    try:
        return str(st.secrets.get("SUPER_ADMIN_USERNAME", "")), str(
            st.secrets.get("SUPER_ADMIN_PASSWORD", "")
        )
    except Exception:
        return "", ""


def make_payslip(row: pd.Series, employee: str, month: str, employees: pd.DataFrame) -> str:
    if "Employee Name" in employees.columns:
        employee_rows = employees[employees["Employee Name"].astype(str).str.strip() == employee]
    else:
        employee_rows = pd.DataFrame()
    employee_row = employee_rows.iloc[0] if not employee_rows.empty else pd.Series(dtype=object)

    def employee_value(*names: str, default: str = "") -> str:
        for name in names:
            value = column_value(employee_row, name, None)
            if value is not None and str(value).strip().casefold() not in {"", "nan"}:
                return str(value).strip()
        return default

    address_parts = [employee_value(name) for name in ("Address 01", "Address 02", "Address 03")]
    address = "<br>".join(safe(value) for value in address_parts if value) or "-"
    try:
        pay_date = f"25 {datetime.strptime(month, '%B-%Y').strftime('%B %Y')}"
    except ValueError:
        pay_date = f"25 {month}"

    earnings = [
        ("Basic Salary", "Basic Salary"),
        ("Budgeted Allowances", "Budgeted Allowance"),
        ("Performance Allowances", "Performance Allowance"),
        ("Field Allowances", "Field Allowance"),
        ("Travelling Allowances", "Travelling Allowance"),
        ("Bike Maintenance", "Bike Maintenance"),
        ("Phone Allowance", "Phone Allowance"),
        ("Incentives", "Incentives"),
    ]
    deductions = [
        ("EPF 8%", "EPF 8%"),
        ("Salary Advance", "Salary Advance"),
        ("Travelling Advance", "Travelling Advance"),
        (f"Staff Loan{installment_suffix(row)}", "Staff Loan"),
        ("Staff Loan Interest", "Staff Loan Interest"),
        ("Other Deduction", "Other Deduction"),
        ("Incentives Paid", "Incentives Paid"),
        ("PAYE Tax", "Paye Tax"),
    ]
    earnings_html = "".join(row_html(label, column_value(row, column)) for label, column in earnings)
    deductions_html = "".join(row_html(label, column_value(row, column)) for label, column in deductions)
    employer_html = "".join(
        row_html(label, column_value(row, column))
        for label, column in (("EPF 12%", "EPF 12%"), ("ETF 3%", "ETF 3%"))
    )
    logo = image_data(APP_DIR / "LOGO.PNG")
    signature = image_data(APP_DIR / "E-Signature.png")
    logo_html = f"<img class='logo' src='data:image/png;base64,{logo}'>" if logo else ""
    signature_html = f"<img class='signature' src='data:image/png;base64,{signature}'>" if signature else ""
    return f"""
    <style>
      * {{ box-sizing: border-box; }}
    @page {{ size: A4 portrait; margin: 0; }}
    body {{ margin: 0; background: #e8edf2; font-family: 'Times New Roman', Times, serif; color: #14202b; font-size: 12px; }}
      .toolbar {{ text-align: center; padding: 16px; }}
      button {{ background: #0b3b60; color: white; border: 0; border-radius: 6px; padding: 11px 20px; font-weight: 700; cursor: pointer; }}
    .page {{ width: 210mm; height: 297mm; min-height: 297mm; margin: 0 auto 20px; padding: 15mm; background: white; box-shadow: 0 3px 18px #0002; overflow: hidden; }}
      .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #0b3b60; padding-bottom: 10px; }}
    .brand {{ display: flex; align-items: center; gap: 10px; color: #0b3b60; font-weight: 800; font-size: 18px; white-space: nowrap; }}
      .logo {{ max-height: 48px; max-width: 90px; }}
    .title {{ font-size: 14px; font-weight: 800; }}
    .contact, .footer {{ color: #586673; font-size: 12px; text-align: center; padding: 7px 0; }}
      .info {{ background: #f3f6f8; border-radius: 6px; padding: 12px; margin: 12px 0; display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
    .label {{ color: #64727d; font-size: 12px; text-transform: uppercase; font-weight: 700; }}
    .value {{ font-size: 12px; font-weight: 700; margin-top: 3px; }}
    .meta {{ display: grid; grid-template-columns: 1fr 1fr; font-size: 12px; gap: 5px; }}
      .meta strong {{ text-align: right; }}
    .section {{ background: #0b3b60; color: white; font-size: 14px; font-weight: 700; padding: 7px 10px; display: flex; justify-content: space-between; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
      td {{ border-bottom: 1px solid #e4e9ed; padding: 6px; }}
      .amount {{ text-align: right; width: 150px; }}
      .total td {{ border-top: 2px solid #0b3b60; border-bottom: 0; color: #0b3b60; font-weight: 800; }}
    .total td {{ font-size: 14px; }}
    .net {{ margin: 12px 0; padding: 12px; background: #eaf2f7; border: 1px solid #c8d8e3; display: flex; justify-content: space-between; color: #0b3b60; font-size: 14px; font-weight: 800; }}
    .signature-block {{ width: 42%; margin: 24px 0 0 auto; text-align: center; }}
    .signature {{ display: block; width: 150px; height: 48px; object-fit: contain; margin: 0 auto 6px; }}
    .sign-line {{ border-top: 1px solid #27333d; padding-top: 6px; font-size: 14px; line-height: 1.35; color: #27333d; font-weight: 700; text-transform: uppercase; letter-spacing: 0.2px; }}
    .sign-role {{ display: none; }}
    @media print {{ .toolbar {{ display: none; }} body {{ background: white; print-color-adjust: exact; -webkit-print-color-adjust: exact; }} .page {{ width: 210mm; height: 297mm; min-height: 297mm; margin: 0; box-shadow: none; page-break-after: avoid; }} }}
      @media (max-width: 800px) {{ .page {{ width: 100%; min-height: auto; padding: 6vw; }} }}
    </style>
    <div class="toolbar"><button onclick="window.print()">Print / Save as PDF</button></div>
    <main class="page">
    <header class="header"><div class="brand">{logo_html}<span>CAPITAL BRIDGE PRIVATE LIMITED</span></div><div class="title">PAYSLIP</div></header>
      <div class="contact">No: 46/1, Thannekubura, Kandy | info@capitalbridge.lk | 081 763 2009</div>
      <section class="info"><div><div class="label">Employee Information</div><div class="value">{safe(employee)}</div><div>{address}<br>Tel: {safe(employee_value('Phone Number'))}</div></div><div class="meta"><span>Pay Date:</span><strong>{safe(pay_date)}</strong><span>Salary Month:</span><strong>{safe(month)}</strong><span>EPF No:</span><strong>{safe(employee_value('EPF No', default='-'))}</strong><span>Position:</span><strong>{safe(employee_value('Designation'))}</strong><span>Department:</span><strong>{safe(employee_value('Department'))}</strong></div></section>
      <div class="section"><span>EARNINGS</span><span>CURRENT (Rs.)</span></div><table>{earnings_html}<tr class="total"><td>Gross Pay</td><td class="amount">{money(column_value(row, 'Gross Salary'))}</td></tr></table>
      <div class="section"><span>DEDUCTIONS</span><span>CURRENT (Rs.)</span></div><table>{deductions_html}<tr class="total"><td>Total Deductions</td><td class="amount">{money(column_value(row, 'Total Deduction'))}</td></tr></table>
      <div class="net"><span>NET PAY</span><span>Rs. {money(column_value(row, 'Net Salary'))}</span></div>
      {f'<div class="section"><span>EMPLOYER CONTRIBUTIONS</span><span>CURRENT (Rs.)</span></div><table>{employer_html}</table>' if employer_html else ''}
    <div class="signature-block">{signature_html}<div class="sign-line">Authorized Signature</div></div>
    <div class="footer">If you have any questions about this payslip, please contact:<br><b>Capital bridge (Pvt) Ltd, Email: info@capitalbridge.lk Tel: 081 763 2009</b></div>
    </main>
    """


st.markdown("# 💼 Capital Bridge Payslip Portal")
st.caption("Direct employee login only. GitHub accounts are not used. Sign in with your Employee Name and personal PIN.")

loaded = load_workbook(None)
if loaded is None:
    st.stop()
salary_data, employee_master = loaded
if (
    "Employee Name" not in salary_data.columns
    or "Month" not in salary_data.columns
    or "Employee Name" not in employee_master.columns
    or "EPF No" not in employee_master.columns
):
    st.error("The workbook must contain Employee Name, EPF No, and Month columns in the expected sheets.")
    st.stop()

salary_data = salary_data.dropna(subset=["Employee Name", "Month"]).copy()
salary_data["Clean Month"] = salary_data["Month"].map(clean_month)

auth_store = load_auth_store()
login_names = sorted(
    employee_master["Employee Name"]
    .dropna()
    .astype(str)
    .str.strip()
    .loc[lambda names: names.ne("")]
    .unique()
)

if "authenticated_employee" not in st.session_state:
    st.session_state.authenticated_employee = None
if "authenticated_admin" not in st.session_state:
    st.session_state.authenticated_admin = False
if "admin_login_requested" not in st.session_state:
    st.session_state.admin_login_requested = False

if not st.session_state.authenticated_admin and st.session_state.authenticated_employee is None:
    with st.sidebar:
        if st.button("Super Admin login"):
            st.session_state.admin_login_requested = True
            st.rerun()

if st.session_state.admin_login_requested and not st.session_state.authenticated_admin:
    st.subheader("Super Admin login")
    admin_username, admin_password = configured_admin()
    if not admin_username or not admin_password:
        st.error("Super Admin is not configured. Set SUPER_ADMIN_USERNAME and SUPER_ADMIN_PASSWORD before starting the app.")
        st.stop()
    with st.form("admin_login"):
        submitted_username = st.text_input("Admin username")
        submitted_password = st.text_input("Admin password", type="password")
        admin_submitted = st.form_submit_button("Log in", type="primary")
    if admin_submitted:
        if secrets.compare_digest(submitted_username, admin_username) and secrets.compare_digest(submitted_password, admin_password):
            st.session_state.authenticated_admin = True
            st.session_state.admin_login_requested = False
            st.rerun()
        st.error("The admin username or password is incorrect.")
    st.stop()

if st.session_state.authenticated_admin:
    st.sidebar.success("Signed in as Super Admin")
    if st.sidebar.button("Log out"):
        st.session_state.clear()
        st.rerun()
    st.subheader("Super Admin")
    st.info("PINs are stored as secure hashes and cannot be viewed. You can see setup status and reset a PIN.")
    status_rows = []
    for name in login_names:
        status_rows.append({"Employee Name": name, "PIN Status": "Configured" if normalized_identity(name) in auth_store else "Not configured"})
    st.dataframe(pd.DataFrame(status_rows), hide_index=True, use_container_width=True)
    reset_name = st.selectbox("Employee whose PIN should be reset", ["Select an employee"] + login_names)
    if st.button("Reset employee PIN", disabled=reset_name == "Select an employee"):
        auth_store.pop(normalized_identity(reset_name), None)
        save_auth_store(auth_store)
        st.success(f"PIN reset for {reset_name}. The employee must create a new PIN using their EPF number.")
        st.rerun()
    st.stop()

if st.session_state.authenticated_employee is None:
    employee_name = st.selectbox("Employee Name", login_names, index=None, placeholder="Select your name")
    has_pin = bool(employee_name and normalized_identity(employee_name) in auth_store)
    login_tab, setup_tab = st.tabs(["Log in", "Create PIN"])
    with login_tab:
        with st.form("employee_login"):
            pin = st.text_input("PIN", type="password", max_chars=8, help="Enter your 4 to 8 digit personal PIN.")
            submitted = st.form_submit_button("Log in", type="primary")
        if submitted:
            if not employee_name or not has_pin or not authenticate_employee(employee_name, pin, auth_store):
                st.error("The Employee Name or PIN is incorrect, or a PIN has not been created yet.")
            else:
                st.session_state.authenticated_employee = employee_name
                st.rerun()
    with setup_tab:
        with st.form("employee_pin_setup"):
            epf_number = st.text_input("EPF Number", type="password")
            new_pin = st.text_input("Create PIN", type="password", max_chars=8)
            confirm_pin = st.text_input("Confirm PIN", type="password", max_chars=8)
            setup_submitted = st.form_submit_button("Create PIN")
        if setup_submitted:
            record = employee_record(employee_name, employee_master) if employee_name else None
            expected_epf = normalized_identity(record.get("EPF No", "")) if record is not None else ""
            if not employee_name or record is None or not secrets.compare_digest(expected_epf, normalized_identity(epf_number)):
                st.error("The Employee Name or EPF Number is incorrect.")
            elif not valid_pin(new_pin):
                st.error("PIN must contain 4 to 8 digits.")
            elif new_pin != confirm_pin:
                st.error("PIN confirmation does not match.")
            elif has_pin:
                st.error("A PIN already exists. Ask the Super Admin to reset it if necessary.")
            else:
                salt, digest = pin_hash(new_pin)
                auth_store[normalized_identity(employee_name)] = {"salt": salt, "hash": digest}
                save_auth_store(auth_store)
                st.success("PIN created successfully. You can now use the Log in tab.")
                st.rerun()
    st.stop()

selected_employee = st.session_state.authenticated_employee
selected_epf = normalized_identity(employee_record(selected_employee, employee_master).get("EPF No", ""))
with st.sidebar:
    st.success(f"Signed in as {selected_employee}")
    if st.button("Super Admin login"):
        st.session_state.admin_login_requested = True
    if st.button("Log out"):
        st.session_state.clear()
        st.rerun()

employee_scope = salary_data["Employee Name"].map(normalized_identity) == normalized_identity(selected_employee)
if "EPF No" in salary_data.columns:
    employee_scope &= salary_data["EPF No"].map(normalized_identity) == selected_epf
employee_months = sorted(salary_data.loc[employee_scope, "Clean Month"].unique())
if not employee_months:
    st.warning("No payslips are available for your account.")
    st.stop()
selected_month = st.selectbox("Salary month", employee_months)

matches = salary_data[
    employee_scope & (salary_data["Clean Month"] == selected_month)
]
if matches.empty:
    st.error("No payslip was found for that employee and month.")
    st.stop()

payslip_html = make_payslip(matches.iloc[0], selected_employee, selected_month, employee_master)
st.download_button(
    "Download payslip",
    data=f"<!doctype html><html><head><meta name='viewport' content='width=device-width, initial-scale=1'></head><body>{payslip_html}</body></html>",
    file_name=f"Payslip-{selected_employee}-{selected_month}.html",
    mime="text/html",
)
components.html(payslip_html, height=1200, scrolling=True)
