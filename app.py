import streamlit as st
import pandas as pd
import os
import glob
import datetime
import base64
import streamlit.components.v1 as components

st.set_page_config(page_title="Employee Payslip Portal", page_icon="💼", layout="wide")

folder_path = r'D:\Hiran\Formats\PayslipApp'
excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

# Helper function to convert local image to base64
def get_image_base64(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode("utf-8")
    return ""

# Load Logo and Signature as Base64
logo_path = os.path.join(folder_path, "LOGO.PNG")
sig_path = os.path.join(folder_path, "E-Signature.png")

logo_base64 = get_image_base64(logo_path)
sig_base64 = get_image_base64(sig_path)

if not excel_files:
    st.error(f"⚠️ '{folder_path}' ෆෝල්ඩර් එකේ කිසිදු Excel (.xlsx) ගොනුවක් හමු වූයේ නැත!")
    st.stop()
else:
    excel_path = excel_files[0]
    
    try:
        xls = pd.ExcelFile(excel_path)
        df_data = pd.read_excel(excel_path, sheet_name="Salary Data")
        df_emp = pd.read_excel(excel_path, sheet_name="Employee Master")
    except Exception as e:
        st.error(f"⚠️ Excel ෆයිල් එක කියවීමේ දෝෂයක්: {e}")
        st.stop()

    df_data = df_data.dropna(subset=['Employee Name', 'Month'])

    def clean_month(val):
        if pd.isna(val):
            return ""
        val_str = str(val).strip()
        parsed_date = pd.to_datetime(val_str, errors='coerce')
        if pd.notna(parsed_date):
            return parsed_date.strftime('%B-%Y')
        return val_str

    df_data['Clean_Month'] = df_data['Month'].apply(clean_month)
    
    names_from_data = df_data['Employee Name'].dropna().astype(str).str.strip().tolist()
    names_from_master = df_emp['Employee Name'].dropna().astype(str).str.strip().tolist() if 'Employee Name' in df_emp.columns else []
    
    employee_list = sorted(list(set(names_from_data + names_from_master)))
    employee_list = [x for x in employee_list if x != '' and x.lower() != 'nan']

    # --- EMPLOYEE PORTAL UI ---
    st.markdown("### 💼 Capital Bridge - Employee Payslip Portal", unsafe_allow_html=True)
    st.markdown("ကျေးজතා කරුණාකර ඔබේ නම සහ අදාළ මාසය තෝරා ඔබේ වැටුප් පත්‍රිකාව (Payslip) නරඹන්න / මුද්‍රණය කරගන්න.")

    if not employee_list:
        st.error("⚠️ Excel ගොනුවේ කිසිදු සේවක නාමයක් හමු වූයේ නැත.")
        st.stop()

    col1, col2 = st.columns(2)
    with col1:
        # සේවකයාට පහසු වීමට මුලින් හිස් තේරීමක් (Select your name) ලබා දීම
        selected_emp = st.selectbox("Select Your Name:", ["-- Select Employee --"] + employee_list)
    
    with col2:
        if selected_emp != "-- Select Employee --":
            emp_months = df_data[df_data['Employee Name'].astype(str).str.strip() == selected_emp]['Clean_Month'].tolist()
            if not emp_months:
                emp_months = sorted(df_data['Clean_Month'].dropna().unique().tolist())
            selected_month_str = st.selectbox("Select Salary Month:", emp_months)
        else:
            selected_month_str = st.selectbox("Select Salary Month:", ["-- Select Month --"])

    if selected_emp == "-- Select Employee --" or selected_month_str == "-- Select Month --":
        st.info("💡 කරුණාකර ඉහත සඳහන් කොටු වලින් ඔබේ නම සහ මාසය තෝරන්න.")
        st.stop()

    # Dynamic Page Title update for Streamlit browser tab & PDF Save name
    st.markdown(f"""
        <script>
            document.title = "Payslip {selected_month_str} - {selected_emp}";
        </script>
    """, unsafe_allow_html=True)

    # Pay Date එක තෝරාගත් මාසයේ 25 වෙනි දිනය ලෙස ස්වයංක්‍රීයව සැකසීම
    try:
        parsed_sel_month = pd.to_datetime(selected_month_str, format='%B-%Y', errors='coerce')
        if pd.notna(parsed_sel_month):
            pay_date_str = f"25 {parsed_sel_month.strftime('%B %Y')}"
        else:
            pay_date_str = f"25 {selected_month_str}"
    except:
        pay_date_str = f"25 {selected_month_str}"

    filtered_df = df_data[(df_data['Employee Name'].astype(str).str.strip() == selected_emp) & (df_data['Clean_Month'] == selected_month_str)]
    
    if filtered_df.empty:
        row = pd.Series(index=df_data.columns, dtype=object)
    else:
        row = filtered_df.iloc[0]

    emp_info = df_emp[df_emp['Employee Name'].astype(str).str.strip() == selected_emp]
    
    address_1, address_2, address_3, phone_no, epf_no, designation, department = "-", "", "", "", "-", "", ""
    
    if not emp_info.empty:
        emp_row = emp_info.iloc[0]
        address_1 = str(emp_row.get('Address 01', '')) if pd.notna(emp_row.get('Address 01')) else str(emp_row.get('Address', ''))
        address_2 = str(emp_row.get('Address 02', '')) if pd.notna(emp_row.get('Address 02')) else ''
        address_3 = str(emp_row.get('Address 03', '')) if pd.notna(emp_row.get('Address 03')) else ''
        phone_no = str(emp_row.get('Phone Number', '')) if pd.notna(emp_row.get('Phone Number')) else ''
        epf_no = str(emp_row.get('EPF No', '')) if pd.notna(emp_row.get('EPF No')) else '-'
        designation = str(emp_row.get('Designation', '')) if pd.notna(emp_row.get('Designation')) else ''
        department = str(emp_row.get('Department', '')) if pd.notna(emp_row.get('Department')) else ''

    full_address = f"{address_1}"
    if address_2 and address_2 != 'nan' and address_2.strip() != '': full_address += f"<br>{address_2}"
    if address_3 and address_3 != 'nan' and address_3.strip() != '': full_address += f"<br>{address_3}"

    def val(col_name):
        for c in row.index:
            if c.strip().lower() == col_name.strip().lower():
                val_data = row[c]
                return val_data if pd.notna(val_data) else 0.0
        return 0.0

    basic = val('Basic Salary')
    budget_allow = val('Budgeted Allowance')
    perf_allow = val('Performance Allowance')
    field_allow = val('Field Allowance')
    trav_allow = val('Travelling Allowance')
    bike_maint = val('Bike Maintenance')
    phone_allow = val('Phone Allowance')
    incentives = val('Incentives')
    gross_pay = val('Gross Salary')

    epf_8 = val('EPF 8%')
    salary_adv = val('Salary Advance')
    trav_adv = val('Travelling Advance')
    staff_loan = val('Staff Loan')
    loan_interest = val('Staff Loan Interest')
    other_ded = val('Other Deduction')
    inc_paid = val('Incentives Paid')
    paye_tax = val('Paye Tax')
    total_ded = val('Total Deduction')
    net_pay = val('Net Salary')

    epf_12 = val('EPF 12%')
    etf_3 = val('ETF 3%')
    
    loan_inst = ''
    for c in row.index:
        if 'loan' in c.lower() and ('inst' in c.lower() or 'installment' in c.lower()):
            if pd.notna(row[c]):
                loan_inst = str(row[c])
                break

    def fmt(v):
        try:
            fv = float(v)
            if fv == 0 or pd.isna(fv):
                return ""
            return f"{fv:,.2f}"
        except:
            return str(v) if v else ""

    def render_row(label, val_num):
        formatted_val = fmt(val_num)
        if formatted_val == "":
            return ""  
        return f"<tr><td style='padding: 5px 6px; border-bottom: 1px solid #e9ecef; font-size: 11px;'>{label}</td><td style='text-align:right; width: 150px; padding: 5px 6px; border-bottom: 1px solid #e9ecef; font-size: 11px; font-weight: 500;'>{formatted_val}</td></tr>"

    earnings_html = ""
    earnings_html += render_row("Basic Salary", basic)
    earnings_html += render_row("Budgeted Allowances", budget_allow)
    earnings_html += render_row("Performance Allowances", perf_allow)
    earnings_html += render_row("Field Allowances", field_allow)
    earnings_html += render_row("Travelling Allowances", trav_allow)
    earnings_html += render_row("Bike Maintained Allowance", bike_maint)
    earnings_html += render_row("Phone Allowance", phone_allow)
    earnings_html += render_row("Incentives", incentives)

    deductions_html = ""
    deductions_html += render_row("EPF 8%", epf_8)
    deductions_html += render_row("Salary Advance", salary_adv)
    deductions_html += render_row("Travelling Advance", trav_adv)
    deductions_html += render_row(f"Staff Loan {f'({loan_inst})' if loan_inst else ''}", staff_loan)
    deductions_html += render_row("Staff Loan Interest", loan_interest)
    deductions_html += render_row("Other Deduction", other_ded)
    deductions_html += render_row("Incentives Paid", inc_paid)
    deductions_html += render_row("Paye Tax", paye_tax)

    # Employer Contributions rows
    employer_contrib_html = ""
    if fmt(epf_12) != "":
        employer_contrib_html += f"<tr><td style='padding: 4px 6px; font-size: 10.5px; color: #444;'>EPF 12%</td><td style='text-align:right; width: 140px; padding: 4px 6px; font-size: 10.5px; font-weight: 500;'>{fmt(epf_12)}</td></tr>"
    if fmt(etf_3) != "":
        employer_contrib_html += f"<tr><td style='padding: 4px 6px; font-size: 10.5px; color: #444;'>ETF 3%</td><td style='text-align:right; width: 140px; padding: 4px 6px; font-size: 10.5px; font-weight: 500;'>{fmt(etf_3)}</td></tr>"

    # HTML string for Logo
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="height: 48px; vertical-align: middle; margin-right: 10px;" />' if logo_base64 else ''
    
    # HTML string for Signature
    sig_html = f'<img src="data:image/png;base64,{sig_base64}" style="max-height: 55px; display: block; margin: 0 auto 2px auto;" />' if sig_base64 else '<div style="height: 35px;"></div>'

    payslip_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <title>Payslip {selected_month_str} - {selected_emp}</title>
    <style>
        body {{
            background-color: #e0e0e0;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
            color: #111111;
            margin: 0;
            padding: 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
        }}
        .a4-page {{
            width: 210mm;
            min-height: 297mm;
            background: #ffffff;
            padding: 12mm 16mm;
            margin: auto;
            box-shadow: 0 4px 15px rgba(0,0,0,0.15);
            box-sizing: border-box;
            position: relative;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}
        .content-container {{
            width: 100%;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th, td {{
            font-size: 11px;
            border: none;
        }}
        .print-btn {{
            margin-bottom: 20px;
            padding: 12px 25px;
            background-color: #002D62;
            color: white;
            text-align: center;
            font-weight: bold;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        }}
        .print-btn:hover {{
            background-color: #001f40;
        }}
        @page {{
            size: A4;
            margin: 0;
        }}
        @media print {{
            body {{
                background-color: #ffffff;
                padding: 0;
            }}
            .print-btn {{
                display: none;
            }}
            .a4-page {{
                box-shadow: none;
                margin: 0;
                width: 100%;
                min-height: 100vh;
                padding: 12mm 16mm;
            }}
        }}
    </style>
    </head>
    <body>

    <button class="print-btn" onclick="window.print()">🖨️ Print / Save as A4 PDF</button>

    <div class="a4-page">
        <div class="content-container">
            <!-- Header with Logo and Title -->
            <table>
                <tr>
                    <td style="width: 70%; vertical-align: middle; padding: 0;">
                        <div style="display: flex; align-items: center;">
                            {logo_html}
                            <div>
                                <div style="font-size: 19px; font-weight: bold; color: #002D62; letter-spacing: 0.5px;">CAPITAL BRIDGE</div>
                                <div style="font-size: 10.5px; font-weight: 600; color: #555; text-transform: uppercase; margin-top: 1px;">PRIVATE LIMITED</div>
                            </div>
                        </div>
                    </td>
                    <td style="width: 30%; font-size: 17px; font-weight: bold; text-align: right; color: #333; vertical-align: middle; padding: 0;">
                        PAYSLIP
                    </td>
                </tr>
            </table>
            <div style="font-size:10px; color:#555; margin-top: 6px; margin-bottom: 12px; border-bottom: 2px solid #002D62; padding-bottom: 6px;">
                No: 46/1, Thannekubura, Kandy &nbsp;|&nbsp; Email: info@capitalbridge.lk &nbsp;|&nbsp; Tel: 081 763 2009
            </div>
            
            <!-- Employee Info Section -->
            <table style="background: #f8f9fa; padding: 10px; border-radius: 6px; margin-bottom: 10px;">
                <tr>
                    <td style="width: 48%; vertical-align: top; padding: 3px;">
                        <div style="font-size: 9.5px; color: #555; font-weight: bold; text-transform: uppercase; margin-bottom: 2px;">Employee Information</div>
                        <div style="font-size: 13px; font-weight: bold; color: #000;">{selected_emp}</div>
                        <div style="font-size: 10.5px; color: #333; margin-top: 2px; line-height: 1.4;">
                            {full_address}
                            {f"<br>Tel: {phone_no}" if phone_no and phone_no != 'nan' else ""}
                        </div>
                    </td>
                    <td style="width: 52%; vertical-align: top; padding: 3px;">
                        <table>
                            <tr>
                                <td style="padding: 2px 0; color: #555;">Pay Date:</td>
                                <td style="padding: 2px 0; font-weight: bold; text-align: right; color: #111;">{pay_date_str}</td>
                            </tr>
                            <tr>
                                <td style="padding: 2px 0; color: #555;">Salary Month:</td>
                                <td style="padding: 2px 0; font-weight: bold; text-align: right; color: #111;">{selected_month_str}</td>
                            </tr>
                            <tr>
                                <td style="padding: 2px 0; color: #555;">Pay Type:</td>
                                <td style="padding: 2px 0; font-weight: bold; text-align: right; color: #111;">Monthly (Online)</td>
                            </tr>
                            <tr>
                                <td style="padding: 2px 0; color: #555;">EPF No:</td>
                                <td style="padding: 2px 0; font-weight: bold; text-align: right; color: #111;">{epf_no}</td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="padding-top: 6px; border-top: 1px solid #dee2e6; margin-top: 4px; font-size: 10.5px; color: #222;">
                        <b>Position:</b> {designation} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Department:</b> {department}
                    </td>
                </tr>
            </table>

            <!-- Earnings Header -->
            <div style="background: #002D62; color: white; font-weight: bold; padding: 5px 10px; font-size: 10.5px; border-radius: 4px 4px 0 0; display: flex; justify-content: space-between;">
                <span>EARNINGS</span>
                <span style="padding-right: 5px;">CURRENT (Rs.)</span>
            </div>
            
            <!-- Earnings Table -->
            <table style="padding: 0 5px; margin-bottom: 6px;">
                {earnings_html}
                <tr>
                    <td style="padding: 6px 6px; font-weight: bold; color: #002D62; border-top: 2px solid #002D62; font-size: 11px;">Gross Pay</td>
                    <td style="text-align:right; width: 150px; padding: 6px 6px; font-weight: bold; color: #002D62; border-top: 2px solid #002D62; font-size: 11px;">{gross_pay:,.2f}</td>
                </tr>
            </table>

            <!-- Deductions Header -->
            <div style="background: #002D62; color: white; font-weight: bold; padding: 5px 10px; font-size: 10.5px; border-radius: 4px 4px 0 0; display: flex; justify-content: space-between; margin-top: 6px;">
                <span>DEDUCTIONS</span>
                <span style="padding-right: 5px;">CURRENT (Rs.)</span>
            </div>
            
            <!-- Deductions Table -->
            <table style="padding: 0 5px; margin-bottom: 6px;">
                {deductions_html}
                <tr>
                    <td style="padding: 6px 6px; font-weight: bold; color: #2c3e50; border-top: 2px solid #2c3e50; font-size: 11px;">Total Deductions</td>
                    <td style="text-align:right; width: 150px; padding: 6px 6px; font-weight: bold; color: #2c3e50; border-top: 2px solid #2c3e50; font-size: 11px;">{total_ded:,.2f}</td>
                </tr>
            </table>

            <!-- Net Pay Box -->
            <table style="background: #f1f4f8; border: 1.5px solid #d0d7de; border-radius: 6px; padding: 10px; margin-top: 6px; margin-bottom: 8px;">
                <tr>
                    <td style="text-align: left; vertical-align: middle; font-size: 12px; color: #002D62; font-weight: bold; text-transform: uppercase;">
                        Net Pay
                    </td>
                    <td style="text-align: right; vertical-align: middle; font-size: 17px; font-weight: bold; color: #002D62; width: 200px;">
                        Rs. {net_pay:,.2f}
                    </td>
                </tr>
            </table>

            <!-- Employer Contributions Separate Section -->
            {f'''
            <div style="background: #002D62; color: white; font-weight: bold; padding: 5px 10px; font-size: 10.5px; border-radius: 4px 4px 0 0; display: flex; justify-content: space-between; margin-top: 6px;">
                <span>EMPLOYER CONTRIBUTIONS</span>
                <span style="padding-right: 5px;">CURRENT (Rs.)</span>
            </div>
            <table style="padding: 0 5px; margin-bottom: 10px; background: #fafbfc; border: 1px solid #e1e4e8; border-top: none; border-radius: 0 0 4px 4px;">
                {employer_contrib_html}
            </table>
            ''' if employer_contrib_html else ''}
        </div>

        <!-- Bottom Fixed Area (Signatures & Footer) -->
        <div>
            <!-- Signature Section with E-Signature Image -->
            <table style="margin-top: 15px; margin-bottom: 12px; border: none;">
                <tr>
                    <td style="width: 50%; border: none;"></td>
                    <td style="width: 50%; text-align: center; border: none; vertical-align: bottom; padding-left: 20px;">
                        {sig_html}
                        <div style="border-top: 1px solid #444; width: 80%; margin: 0 auto; padding-top: 5px; font-size: 11px; font-weight: bold;">Authorized Signature / Director</div>
                    </td>
                </tr>
            </table>
            
            <!-- Footer -->
            <div style="font-size: 9px; text-align: center; color: #555; border-top: 1px solid #ddd; padding-top: 6px;">
                If you have any questions about this payslip, please contact:<br><b>Capital Bridge (Pvt) Ltd, Email: info@capitalbridge.lk | Tel: 081 763 2009</b>
            </div>
        </div>
    </div>

    </body>
    </html>
    """

    components.html(payslip_html, height=1150, scrolling=True)
