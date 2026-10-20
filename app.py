import streamlit as st
import os
import re
import glob
import pandas as pd
import numpy as np
import pdfplumber
import openpyxl
import tempfile

st.set_page_config(page_title="GST Audit Report Generator", page_icon="📊", layout="wide")

st.title("📊 GST Consolidated Audit Report Generator")
st.markdown("Upload your monthly GST return PDFs below to generate the complete audit workbook automatically.")

st.sidebar.header("📁 File Uploads")
uploaded_pdfs = st.sidebar.file_uploader("Upload GST PDFs (GSTR-1, 2B, 3B, 9)", type=["pdf"], accept_multiple_files=True)

if st.button("🚀 Generate Consolidated Report", type="primary"):
    template_path = "Desired Output.xlsx"
    
    if not os.path.exists(template_path):
        st.error(f"❌ Template file '{template_path}' not found in the app directory. Please place it in your project folder.")
    elif not uploaded_pdfs:
        st.error("❌ Please upload at least one GST PDF file.")
    else:
        with st.spinner("Processing files and compiling audit reports..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                local_template_path = os.path.join(tmpdir, "Desired Output.xlsx")
                with open(template_path, "rb") as src, open(local_template_path, "wb") as dst:
                    dst.write(src.read())

                for pdf_file in uploaded_pdfs:
                    pdf_path = os.path.join(tmpdir, pdf_file.name)
                    with open(pdf_path, "wb") as f:
                        f.write(pdf_file.getbuffer())

                output_path = os.path.join(tmpdir, "GST_Consolidated_Audit_Report.xlsx")

                reader_wb = openpyxl.load_workbook(local_template_path)
                sheet_names = reader_wb.sheetnames

                MONTHS = ["APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "JAN", "FEB", "MAR"]

                due_dates_map = {
                    'APR': '20/05/2024', 'MAY': '20/06/2024', 'JUN': '20/07/2024',
                    'JUL': '20/08/2024', 'AUG': '20/09/2024', 'SEP': '20/10/2024',
                    'OCT': '20/11/2024', 'NOV': '20/12/2024', 'DEC': '20/01/2025',
                    'JAN': '20/02/2025', 'FEB': '20/03/2025', 'MAR': '20/04/2025'
                }

                def clean_val(val):
                    if val is None or pd.isna(val):
                        return 0.0
                    val_str = str(val).replace(",", "").replace(" ", "").replace("\n", "").strip()
                    if val_str in ["", "-", ".", "0.00"]:
                        return 0.0
                    try:
                        return float(val_str)
                    except ValueError:
                        return val_str  # Return original string for dates

                def get_month_key_from_filename(filename):
                    match = re.search(r'_(\d{2})\d{4}', filename)
                    if match:
                        m_num = match.group(1)
                        months_map = {
                            '04': 'APR', '05': 'MAY', '06': 'JUN', '07': 'JUL',
                            '08': 'AUG', '09': 'SEP', '10': 'OCT', '11': 'NOV',
                            '12': 'DEC', '01': 'JAN', '02': 'FEB', '03': 'MAR'
                        }
                        return months_map.get(m_num, None)
                    return None

                g1_data = {}
                g2b_data = {}
                g3b_data = {}
                g9_data = {"2023-24": {}, "2024-25": {}}

                # Parse GSTR-1 PDFs
                for f_path in glob.glob(os.path.join(tmpdir, "GSTR1_*.pdf")):
                    m = get_month_key_from_filename(os.path.basename(f_path))
                    if m:
                        g1_data[m] = {
                            "B2B_Taxable": 0.0, "B2B_IGST": 0.0, "B2B_CGST": 0.0, "B2B_SGST": 0.0, "B2B_CESS": 0.0,
                            "RC_Taxable": 0.0, "RC_IGST": 0.0, "RC_CGST": 0.0, "RC_SGST": 0.0, "RC_CESS": 0.0,
                            "B2CS_Net": 0.0, "B2CS_IGST": 0.0, "B2CS_CGST": 0.0, "B2CS_SGST": 0.0, "B2CS_CESS": 0.0
                        }
                        try:
                            with pdfplumber.open(f_path) as pdf:
                                text = "\n".join([page.extract_text() or "" for page in pdf.pages])

                            b2b = re.search(r"4A.*?Total.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if b2b:
                                g1_data[m]["B2B_Taxable"] = clean_val(b2b.group(1))
                                g1_data[m]["B2B_IGST"] = clean_val(b2b.group(2))
                                g1_data[m]["B2B_CGST"] = clean_val(b2b.group(3))
                                g1_data[m]["B2B_SGST"] = clean_val(b2b.group(4))
                                g1_data[m]["B2B_CESS"] = clean_val(b2b.group(5))

                            rc_match = re.search(r"4B.*?Total.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if rc_match:
                                g1_data[m]["RC_Taxable"] = clean_val(rc_match.group(1))
                                g1_data[m]["RC_IGST"] = clean_val(rc_match.group(2))
                                g1_data[m]["RC_CGST"] = clean_val(rc_match.group(3))
                                g1_data[m]["RC_SGST"] = clean_val(rc_match.group(4))
                                g1_data[m]["RC_CESS"] = clean_val(rc_match.group(5))

                            b2cs = re.search(r"7\s*-\s*Taxable.*?Total.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if b2cs:
                                g1_data[m]["B2CS_Net"] = clean_val(b2cs.group(1))
                                g1_data[m]["B2CS_IGST"] = clean_val(b2cs.group(2))
                                g1_data[m]["B2CS_CGST"] = clean_val(b2cs.group(3))
                                g1_data[m]["B2CS_SGST"] = clean_val(b2cs.group(4))
                                g1_data[m]["B2CS_CESS"] = clean_val(b2cs.group(5))
                        except Exception as e:
                            pass

                # Parse GSTR-2B PDFs
                for f_path in glob.glob(os.path.join(tmpdir, "GSTR2B_*.pdf")):
                    m = get_month_key_from_filename(os.path.basename(f_path))
                    if m:
                        g2b_data[m] = {
                            "PartA_IGST": 0.0, "PartA_CGST": 0.0, "PartA_SGST": 0.0, "PartA_CESS": 0.0,
                            "CreditNotes_IGST": 0.0, "CreditNotes_CGST": 0.0, "CreditNotes_SGST": 0.0, "CreditNotes_CESS": 0.0,
                            "NotAvailable_IGST": 0.0, "NotAvailable_CGST": 0.0, "NotAvailable_SGST": 0.0, "NotAvailable_CESS": 0.0,
                            "Rejected_IGST": 0.0, "Rejected_CGST": 0.0, "Rejected_SGST": 0.0, "Rejected_CESS": 0.0
                        }
                        try:
                            with pdfplumber.open(f_path) as pdf:
                                text = "\n".join([page.extract_text() or "" for page in pdf.pages])

                            match_itc = re.search(r"All other ITC.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE | re.DOTALL)
                            if match_itc:
                                g2b_data[m]["PartA_IGST"] = clean_val(match_itc.group(1))
                                g2b_data[m]["PartA_CGST"] = clean_val(match_itc.group(2))
                                g2b_data[m]["PartA_SGST"] = clean_val(match_itc.group(3))
                                g2b_data[m]["PartA_CESS"] = clean_val(match_itc.group(4))

                            match_cn = re.search(r"Credit Notes.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE | re.DOTALL)
                            if match_cn:
                                g2b_data[m]["CreditNotes_IGST"] = clean_val(match_cn.group(1))
                                g2b_data[m]["CreditNotes_CGST"] = clean_val(match_cn.group(2))
                                g2b_data[m]["CreditNotes_SGST"] = clean_val(match_cn.group(3))
                                g2b_data[m]["CreditNotes_CESS"] = clean_val(match_cn.group(4))

                            match_na = re.search(r"ITC Not Available.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE | re.DOTALL)
                            if match_na:
                                g2b_data[m]["NotAvailable_IGST"] = clean_val(match_na.group(1))
                                g2b_data[m]["NotAvailable_CGST"] = clean_val(match_na.group(2))
                                g2b_data[m]["NotAvailable_SGST"] = clean_val(match_na.group(3))
                                g2b_data[m]["NotAvailable_CESS"] = clean_val(match_na.group(4))

                            match_rej = re.search(r"ITC Rejected.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE | re.DOTALL)
                            if match_rej:
                                g2b_data[m]["Rejected_IGST"] = clean_val(match_rej.group(1))
                                g2b_data[m]["Rejected_CGST"] = clean_val(match_rej.group(2))
                                g2b_data[m]["Rejected_SGST"] = clean_val(match_rej.group(3))
                                g2b_data[m]["Rejected_CESS"] = clean_val(match_rej.group(4))
                        except Exception as e:
                            pass

                # Parse GSTR-3B PDFs
                for f_path in glob.glob(os.path.join(tmpdir, "GSTR3B_*.pdf")):
                    m = get_month_key_from_filename(os.path.basename(f_path))
                    if m:
                        g3b_data[m] = {
                            "Eligible_IGST": 0.0, "Eligible_CGST": 0.0, "Eligible_SGST": 0.0, "Eligible_CESS": 0.0,
                            "ImportGoods_IGST": 0.0, "ImportGoods_CESS": 0.0,
                            "ReverseCharge_IGST": 0.0, "ReverseCharge_CGST": 0.0, "ReverseCharge_SGST": 0.0, "ReverseCharge_CESS": 0.0,
                            "ISD_IGST": 0.0, "ISD_CGST": 0.0, "ISD_SGST": 0.0, "ISD_CESS": 0.0,
                            "ReversedRules_IGST": 0.0, "ReversedRules_CGST": 0.0, "ReversedRules_SGST": 0.0, "ReversedRules_CESS": 0.0,
                            "ReversedOthers_IGST": 0.0, "ReversedOthers_CGST": 0.0, "ReversedOthers_SGST": 0.0, "ReversedOthers_CESS": 0.0,
                            "NetITC_IGST": 0.0, "NetITC_CGST": 0.0, "NetITC_SGST": 0.0, "NetITC_CESS": 0.0,
                            "Interest_System": 0.0, "Interest_Paid": 0.0,
                            "LateFee_System": 0.0, "LateFee_Paid": 0.0,
                            "Cash_CGST": 0.0, "Cash_SGST": 0.0, "Days_Late": 0.0,
                            "Filing_Date": "-"
                        }
                        try:
                            with pdfplumber.open(f_path) as pdf:
                                text = "\n".join([page.extract_text() or "" for page in pdf.pages])

                            filing_match = re.search(r"(?:Date of Filing|Date of ARN)\s*([\d\-/]+)", text, re.IGNORECASE)
                            if filing_match:
                                g3b_data[m]["Filing_Date"] = filing_match.group(1).strip()

                            itc_match = re.search(r"(?:All other ITC|4\(A\)\(5\)).*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if itc_match:
                                g3b_data[m]["Eligible_IGST"] = clean_val(itc_match.group(1))
                                g3b_data[m]["Eligible_CGST"] = clean_val(itc_match.group(2))
                                g3b_data[m]["Eligible_SGST"] = clean_val(itc_match.group(3))

                            imp_match = re.search(r"Import of goods.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if imp_match:
                                g3b_data[m]["ImportGoods_IGST"] = clean_val(imp_match.group(1))
                                g3b_data[m]["ImportGoods_CESS"] = clean_val(imp_match.group(2))

                            rc_itc = re.search(r"Inward supplies liable to reverse charge.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if rc_itc:
                                g3b_data[m]["ReverseCharge_IGST"] = clean_val(rc_itc.group(1))
                                g3b_data[m]["ReverseCharge_CGST"] = clean_val(rc_itc.group(2))
                                g3b_data[m]["ReverseCharge_SGST"] = clean_val(rc_itc.group(3))

                            isd_match = re.search(r"Inward supplies from ISD.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if isd_match:
                                g3b_data[m]["ISD_IGST"] = clean_val(isd_match.group(1))
                                g3b_data[m]["ISD_CGST"] = clean_val(isd_match.group(2))
                                g3b_data[m]["ISD_SGST"] = clean_val(isd_match.group(3))

                            rev_rules = re.search(r"As per rules 38.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if rev_rules:
                                g3b_data[m]["ReversedRules_IGST"] = clean_val(rev_rules.group(1))
                                g3b_data[m]["ReversedRules_CGST"] = clean_val(rev_rules.group(2))
                                g3b_data[m]["ReversedRules_SGST"] = clean_val(rev_rules.group(3))

                            rev_others = re.search(r"(?:Others|4\(B\)\(2\)).*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if rev_others:
                                g3b_data[m]["ReversedOthers_IGST"] = clean_val(rev_others.group(1))
                                g3b_data[m]["ReversedOthers_CGST"] = clean_val(rev_others.group(2))
                                g3b_data[m]["ReversedOthers_SGST"] = clean_val(rev_others.group(3))

                            net_itc = re.search(r"Net ITC Available.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if net_itc:
                                g3b_data[m]["NetITC_IGST"] = clean_val(net_itc.group(1))
                                g3b_data[m]["NetITC_CGST"] = clean_val(net_itc.group(2))
                                g3b_data[m]["NetITC_SGST"] = clean_val(net_itc.group(3))

                            interest_match = re.search(r"Interest.*?\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if interest_match:
                                g3b_data[m]["Interest_System"] = clean_val(interest_match.group(1))
                            late_match = re.search(r"Late fee.*?\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if late_match:
                                g3b_data[m]["LateFee_System"] = clean_val(late_match.group(1))

                            payment_cash_cgst = re.search(r"Central Tax.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if payment_cash_cgst:
                                g3b_data[m]["Cash_CGST"] = clean_val(payment_cash_cgst.group(1))
                            payment_cash_sgst = re.search(r"State/UT Tax.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.IGNORECASE)
                            if payment_cash_sgst:
                                g3b_data[m]["Cash_SGST"] = clean_val(payment_cash_sgst.group(1))

                            days_match = re.search(r"delayed.*?by.*?(\d+)\s+days", text, re.IGNORECASE)
                            if days_match:
                                g3b_data[m]["Days_Late"] = float(days_match.group(1))
                        except Exception as e:
                            pass

                # Parse GSTR-9 PDFs
                for f_path in glob.glob(os.path.join(tmpdir, "GSTR9_*.pdf")):
                    fy = "2023-24" if "032024" in f_path else "2024-25"
                    g9_data[fy] = {}
                    try:
                        with pdfplumber.open(f_path) as pdf:
                            full_text = "\n".join([p.extract_text() or "" for p in pdf.pages])
                        clean_t = re.sub(r' +', ' ', full_text)
                        patterns = {
                            "4A": r"A\s+Supplies\s+made\s+to\s+un-registered\s+persons.*?\s+([\d, .]+)",
                            "4B": r"B\s+Supplies\s+made\s+to\s+registered\s+persons.*?\s+([\d, .]+)",
                            "4C": r"C\s+Zero\s+rated\s+supply.*?\s+([\d, .]+)",
                            "4D": r"D\s+Supplies\s+to\s+SEZs.*?\s+([\d, .]+)",
                            "4E": r"E\s+Deemed\s+Exports.*?\s+([\d, .]+)",
                            "4F": r"F\s+Advances.*?\s+([\d, .]+)",
                            "4G": r"G\s+Inward\s+supplies.*?reverse.*?\s+([\d, .]+)",
                            "4I": r"I\s+Credit\s+notes.*?\s+([\d, .]+)",
                            "4J": r"J\s+Debit\s+notes.*?\s+([\d, .]+)",
                            "4K": r"K\s+Supplies.*?declared.*?Amendments.*?\s+([\d, .]+)",
                            "4L": r"L\s+Supplies.*?reduced.*?Amendments.*?\s+([\d, .]+)",
                            "5A": r"A\s+Zero\s+rated\s+supply.*?without\s+payment.*?\s+([\d, .]+)",
                            "5B": r"B\s+Supplies\s+to\s+SEZs.*?without\s+payment.*?\s+([\d, .]+)",
                            "5C": r"C\s+Supplies\s+on\s+which\s+tax\s+is\s+to\s+be\s+paid.*?recipient.*?\s+([\d, .]+)",
                            "5D": r"D\s+Exempted\s+([\d, .]+)",
                            "5E": r"E\s+Nil\s+Rated\s+([\d, .]+)",
                            "5F": r"F\s+Non-GST\s+supply.*?\s+([\d, .]+)",
                            "6A": r"A\s+Total\s+amount\s+of\s+input\s+tax\s+credit\s+availed\s+through\s+FORM\s+GSTR-3B.*?\s+([\d, .]+)",
                            "6B": r"B\s+Inward\s+supplies\s+\(other\s+than\s+imports.*?\s+([\d, .]+)",
                            "6C": r"C\s+Inward\s+supplies\s+received\s+from\s+unregistered.*?\s+([\d, .]+)",
                            "6D": r"D\s+Inward\s+supplies\s+received\s+from\s+registered.*?\s+([\d, .]+)",
                            "7I": r"Total\s+ITC\s+Reversed.*?\s+([\d, .]+)",
                            "8A": r"A\s+ITC\s+as\s+per\s+GSTR-2B.*?\s+([\d, .]+)",
                            "8F": r"F\s+ITC\s+available\s+but\s+ineligible.*?\s+([\d, .]+)"
                        }
                        for key, pattern in patterns.items():
                            match = re.search(pattern, clean_t, re.IGNORECASE | re.DOTALL)
                            if match:
                                tokens = [t for t in match.group(1).split() if '.' in t or t in ['0', '0.00']]
                                if len(tokens) >= 3:
                                    g9_data[fy][key] = {
                                        "Value": clean_val(tokens[0]) if len(tokens) >= 5 else 0.0,
                                        "CGST": clean_val(tokens[1]) if len(tokens) >= 5 else clean_val(tokens[0]),
                                        "SGST": clean_val(tokens[2]) if len(tokens) >= 5 else clean_val(tokens[1]),
                                        "IGST": clean_val(tokens[3]) if len(tokens) >= 5 else clean_val(tokens[2]),
                                        "CESS": clean_val(tokens[4]) if len(tokens) >= 5 else 0.0
                                    }
                    except Exception as e:
                        pass

                output_sheets = {}
                for sheet_name in sheet_names:
                    df = pd.read_excel(local_template_path, sheet_name=sheet_name)
                    heading_col = 'HEADING' if 'HEADING' in df.columns else df.columns[0]

                    # FIX: Coerce all month columns to object/float types to prevent integer strictness type errors
                    for m in MONTHS:
                        if m in df.columns:
                            df[m] = df[m].astype(object)

                    if sheet_name == 'GSTR-1 Summary (12M)':
                        section_col = next((c for c in df.columns if str(c).strip().lower() in ['table section', 'section', 'tablesection', 'table_section']), heading_col)
                        details_col = next((c for c in df.columns if str(c).strip().lower() in ['details / field', 'details', 'field', 'details/field', 'details_field']), heading_col)
                        for idx, row in df.iterrows():
                            section = str(row[section_col]).strip().lower()
                            field = str(row[details_col]).strip().lower()
                            for m in MONTHS:
                                if m in g1_data:
                                    if 'outward supplies other than reverse charge' in section or 'total liability' in section:
                                        if 'taxable' in field: df.at[idx, m] = g1_data[m]["B2B_Taxable"]
                                        elif 'igst' in field: df.at[idx, m] = g1_data[m]["B2B_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g1_data[m]["B2B_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g1_data[m]["B2B_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g1_data[m]["B2B_CESS"]
                                    elif '4b' in section or 'reverse charge' in section:
                                        if 'taxable' in field: df.at[idx, m] = g1_data[m]["RC_Taxable"]
                                        elif 'igst' in field: df.at[idx, m] = g1_data[m]["RC_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g1_data[m]["RC_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g1_data[m]["RC_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g1_data[m]["RC_CESS"]
                                    elif '4a' in section or 'b2b' in section:
                                        if 'taxable' in field: df.at[idx, m] = g1_data[m]["B2B_Taxable"]
                                        elif 'igst' in field: df.at[idx, m] = g1_data[m]["B2B_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g1_data[m]["B2B_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g1_data[m]["B2B_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g1_data[m]["B2B_CESS"]
                                    elif '7' in section or 'b2cs' in section:
                                        if 'net' in field or 'taxable' in field: df.at[idx, m] = g1_data[m]["B2CS_Net"]
                                        elif 'igst' in field: df.at[idx, m] = g1_data[m]["B2CS_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g1_data[m]["B2CS_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g1_data[m]["B2CS_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g1_data[m]["B2CS_CESS"]
                        for m in MONTHS:
                            df[m] = df[m].apply(clean_val)
                        numeric_only = df[MONTHS].apply(pd.to_numeric, errors='coerce').fillna(0.0)
                        df['FY TOTAL'] = numeric_only.sum(axis=1)

                    elif sheet_name == 'GSTR-2B Summary (12M)':
                        section_col = next((c for c in df.columns if str(c).strip().lower() in ['table section', 'section', 'tablesection', 'table_section']), heading_col)
                        details_col = next((c for c in df.columns if str(c).strip().lower() in ['details / field', 'details', 'field', 'details/field', 'details_field']), heading_col)
                        for idx, row in df.iterrows():
                            section = str(row[section_col]).strip().lower()
                            field = str(row[details_col]).strip().lower()
                            for m in MONTHS:
                                if m in g2b_data:
                                    if 'credit notes - others' in section or 'part b' in section:
                                        if 'igst' in field or 'integrated' in field: df.at[idx, m] = g2b_data[m]["CreditNotes_IGST"]
                                        elif 'cgst' in field or 'central' in field: df.at[idx, m] = g2b_data[m]["CreditNotes_CGST"]
                                        elif 'sgst' in field or 'state' in field: df.at[idx, m] = g2b_data[m]["CreditNotes_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g2b_data[m]["CreditNotes_CESS"]
                                    elif 'not available' in section:
                                        if 'igst' in field or 'integrated' in field: df.at[idx, m] = g2b_data[m]["NotAvailable_IGST"]
                                        elif 'cgst' in field or 'central' in field: df.at[idx, m] = g2b_data[m]["NotAvailable_CGST"]
                                        elif 'sgst' in field or 'state' in field: df.at[idx, m] = g2b_data[m]["NotAvailable_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g2b_data[m]["NotAvailable_CESS"]
                                    elif 'rejected' in section or 'ims' in section:
                                        if 'igst' in field or 'integrated' in field: df.at[idx, m] = g2b_data[m]["Rejected_IGST"]
                                        elif 'cgst' in field or 'central' in field: df.at[idx, m] = g2b_data[m]["Rejected_CGST"]
                                        elif 'sgst' in field or 'state' in field: df.at[idx, m] = g2b_data[m]["Rejected_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g2b_data[m]["Rejected_CESS"]
                                    elif 'all other' in section or '4(a)(5)' in section:
                                        if 'igst' in field or 'integrated' in field: df.at[idx, m] = g2b_data[m]["PartA_IGST"]
                                        elif 'cgst' in field or 'central' in field: df.at[idx, m] = g2b_data[m]["PartA_CGST"]
                                        elif 'sgst' in field or 'state' in field: df.at[idx, m] = g2b_data[m]["PartA_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g2b_data[m]["PartA_CESS"]
                        for m in MONTHS:
                            df[m] = df[m].apply(clean_val)
                        numeric_only = df[MONTHS].apply(pd.to_numeric, errors='coerce').fillna(0.0)
                        df['FY TOTAL'] = numeric_only.sum(axis=1)

                    elif sheet_name == 'GSTR-3B Summary (12M)':
                        section_col = next((c for c in df.columns if str(c).strip().lower() in ['table section', 'section', 'tablesection', 'table_section']), heading_col)
                        details_col = next((c for c in df.columns if str(c).strip().lower() in ['details / field', 'details', 'field', 'details/field', 'details_field']), heading_col)
                        for idx, row in df.iterrows():
                            section = str(row[section_col]).strip().lower()
                            field = str(row[details_col]).strip().lower()
                            for m in MONTHS:
                                if m in g3b_data:
                                    if 'due date' in section or 'due date' in field:
                                        df.at[idx, m] = due_dates_map.get(m, "-")
                                    elif 'filing date' in section or 'filing date' in field:
                                        df.at[idx, m] = g3b_data[m]["Filing_Date"]
                                    elif 'import of goods' in section or '(1)' in section:
                                        if 'igst' in field: df.at[idx, m] = g3b_data[m]["ImportGoods_IGST"]
                                        elif 'cess' in field: df.at[idx, m] = g3b_data[m]["ImportGoods_CESS"]
                                    elif 'liable to reverse charge' in section or '(3)' in section:
                                        if 'igst' in field: df.at[idx, m] = g3b_data[m]["ReverseCharge_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g3b_data[m]["ReverseCharge_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g3b_data[m]["ReverseCharge_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g3b_data[m]["ReverseCharge_CESS"]
                                    elif 'from isd' in section or '(4)' in section:
                                        if 'igst' in field: df.at[idx, m] = g3b_data[m]["ISD_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g3b_data[m]["ISD_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g3b_data[m]["ISD_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g3b_data[m]["ISD_CESS"]
                                    elif 'all other' in section or '(5)' in section:
                                        if 'igst' in field: df.at[idx, m] = g3b_data[m]["Eligible_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g3b_data[m]["Eligible_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g3b_data[m]["Eligible_SGST"]
                                        elif 'cess' in field: df.at[idx, m] = g3b_data[m]["Eligible_CESS"]
                                    elif 'rules 38,42' in section or 'section 17(5)' in section or '(1)' in section and 'reversed' in section:
                                        if 'igst' in field: df.at[idx, m] = g3b_data[m]["ReversedRules_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g3b_data[m]["ReversedRules_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g3b_data[m]["ReversedRules_SGST"]
                                    elif 'reversed - (2) others' in section or 'others' in section and 'reversed' in section:
                                        if 'igst' in field: df.at[idx, m] = g3b_data[m]["ReversedOthers_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g3b_data[m]["ReversedOthers_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g3b_data[m]["ReversedOthers_SGST"]
                                    elif 'net itc' in section or '4 (c)' in section:
                                        if 'igst' in field: df.at[idx, m] = g3b_data[m]["NetITC_IGST"]
                                        elif 'cgst' in field: df.at[idx, m] = g3b_data[m]["NetITC_CGST"]
                                        elif 'sgst' in field: df.at[idx, m] = g3b_data[m]["NetITC_SGST"]
                                    elif 'interest' in section or 'late fee' in section:
                                        if 'system' in field or 'computed' in field:
                                            df.at[idx, m] = g3b_data[m]["Interest_System"] + g3b_data[m]["LateFee_System"]
                                        elif 'paid' in field:
                                            df.at[idx, m] = g3b_data[m]["Interest_Paid"] + g3b_data[m]["LateFee_Paid"]
                        for m in MONTHS:
                            df[m] = df[m].apply(clean_val)
                        numeric_only = df[MONTHS].apply(pd.to_numeric, errors='coerce').fillna(0.0)
                        df['FY TOTAL'] = numeric_only.sum(axis=1)

                    elif 'gstr 9' in sheet_name.lower() or 'gstr-9' in sheet_name.lower():
                        target_fy = "2023-24" if "2023-24" in sheet_name else "2024-25"
                        for idx, row in df.iterrows():
                            text_val = str(row[heading_col]).lower()
                            matched_key = None
                            for key in ["4a", "4b", "4c", "4d", "4e", "4f", "4g", "4i", "4j", "4k", "4l", "5a", "5b", "5c", "5d", "5e", "5f", "6a", "6b", "6c", "6d", "7i", "8a", "8f"]:
                                if key in text_val:
                                    matched_key = key.upper()
                                    break
                            if matched_key and target_fy in g9_data and matched_key in g9_data[target_fy]:
                                data = g9_data[target_fy][matched_key]
                                if 'Taxable Value' in df.columns: df.at[idx, 'Taxable Value'] = data.get("Value", 0.0)
                                if 'CGST' in df.columns: df.at[idx, 'CGST'] = data.get("CGST", 0.0)
                                if 'SGST' in df.columns: df.at[idx, 'SGST'] = data.get("SGST", 0.0)
                                if 'IGST' in df.columns: df.at[idx, 'IGST'] = data.get("IGST", 0.0)
                                if 'CESS' in df.columns: df.at[idx, 'CESS'] = data.get("CESS", 0.0)

                    output_sheets[sheet_name] = df

                recon_rows = []
                for m in MONTHS:
                    g3b = g3b_data.get(m, {}).get("Eligible_IGST", 0.0)
                    g2b = g2b_data.get(m, {}).get("PartA_IGST", 0.0)
                    dev = round(max(0.0, g3b - g2b), 2)
                    recon_rows.append({
                        "Period": m, "GSTR-3B ITC Availed": g3b, "GSTR-2B ITC Available": g2b, "Deviation": dev,
                        "Risk Status": "Excess Availment Risk" if dev > 100 else "Satisfactory"
                    })
                output_sheets["1. ITC Supplier Recon"] = pd.DataFrame(recon_rows)

                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    for sheet_name, df_data in output_sheets.items():
                        df_data.to_excel(writer, sheet_name=sheet_name, index=False)

                st.success("🎉 Audit Report successfully generated!")

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Download GST Consolidated Audit Report",
                        data=f,
                        file_name="GST_Consolidated_Audit_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )