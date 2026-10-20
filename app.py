import streamlit as st
import os
import re
import glob
import pandas as pd
import numpy as np
import pdfplumber
import openpyxl
import tempfile
from datetime import datetime

st.set_page_config(page_title="GST Consolidated Working File Generator", page_icon="📊", layout="wide")

st.title("📊 GST Consolidated Working File Generator")
st.markdown("Upload your monthly/quarterly GST return PDFs below to generate the complete audit working file automatically.")

st.sidebar.header("📁 File Uploads")
uploaded_pdfs = st.sidebar.file_uploader("Upload GST PDFs (GSTR-1, 2B, 3B, 9, etc.)", type=["pdf"], accept_multiple_files=True)

if st.button("🚀 Generate Consolidated Report", type="primary"):
    template_path = "GST_Audit_Working_File (2).xlsx"
    if not os.path.exists(template_path):
        template_path = "Desired Output.xlsx"

    if not os.path.exists(template_path):
        st.error("❌ Reference template Excel file not found in the app directory. Please place it in your project folder.")
    elif not uploaded_pdfs:
        st.error("❌ Please upload at least one GST PDF file.")
    else:
        with st.spinner("Processing files and compiling audit working files..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                local_template_path = os.path.join(tmpdir, "Desired Output.xlsx")
                with open(template_path, "rb") as src, open(local_template_path, "wb") as dst:
                    dst.write(src.read())

                for pdf_file in uploaded_pdfs:
                    pdf_path = os.path.join(tmpdir, pdf_file.name)
                    with open(pdf_path, "wb") as f:
                        f.write(pdf_file.getbuffer())

                output_path = os.path.join(tmpdir, "GST_Audit_Working_File.xlsx")

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
                    if val is None or pd.isna(val) or str(val).strip() in ["", "-", "."]:
                        return 0.0
                    val_str = str(val).replace(",", "").replace(" ", "").replace("\n", "").strip()
                    if val_str in ["", "-", ".", "0.00"]:
                        return 0.0
                    try:
                        return float(val_str)
                    except ValueError:
                        return val_str

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

                # Parse GSTR-2B / GSTR-2BQ PDFs
                for f_path in glob.glob(os.path.join(tmpdir, "GSTR2B_*.pdf")) + glob.glob(os.path.join(tmpdir, "GSTR2BQ_*.pdf")):
                    filename = os.path.basename(f_path)
                    m = get_month_key_from_filename(filename)
                    if not m:
                        if "06" in filename or "Q1" in filename.upper(): m = "JUN"
                        elif "09" in filename or "Q2" in filename.upper(): m = "SEP"
                        elif "12" in filename or "Q3" in filename.upper(): m = "DEC"
                        elif "03" in filename or "Q4" in filename.upper(): m = "MAR"

                    if m:
                        if m not in g2b_data:
                            g2b_data[m] = {
                                "PartA_IGST": 0.0, "PartA_CGST": 0.0, "PartA_SGST": 0.0, "PartA_CESS": 0.0,
                                "CreditNotes_IGST": 0.0, "CreditNotes_CGST": 0.0, "CreditNotes_SGST": 0.0, "CreditNotes_CESS": 0.0,
                                "NotAvailable_IGST": 0.0, "NotAvailable_CGST": 0.0, "NotAvailable_SGST": 0.0, "NotAvailable_CESS": 0.0
                            }
                        try:
                            with pdfplumber.open(f_path) as pdf:
                                text = "\n".join([page.extract_text() or "" for page in pdf.pages])

                            # Split text into Table 3 (ITC Available) and Table 4 (ITC Not Available) blocks
                            t3_block = text
                            t4_block = ""
                            if "4." in text or "ITC Not Available" in text:
                                parts = re.split(r"(?:4\.\s*ITC Not Available|ITC Not Available)", text, flags=re.IGNORECASE)
                                if len(parts) >= 2:
                                    t3_block = parts[0]
                                    t4_block = parts[1]

                            # Table 3 - ITC Available (All other ITC)
                            match_itc = re.search(r"All other ITC.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", t3_block, re.IGNORECASE | re.DOTALL)
                            if match_itc:
                                g2b_data[m]["PartA_IGST"] = clean_val(match_itc.group(1))
                                g2b_data[m]["PartA_CGST"] = clean_val(match_itc.group(2))
                                g2b_data[m]["PartA_SGST"] = clean_val(match_itc.group(3))
                                g2b_data[m]["PartA_CESS"] = clean_val(match_itc.group(4))

                            # Table 3 Part B - Credit Notes (Others)
                            match_cn = re.search(r"Credit Notes.*?others.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", t3_block, re.IGNORECASE | re.DOTALL)
                            if not match_cn:
                                match_cn = re.search(r"Credit Notes.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", t3_block, re.IGNORECASE | re.DOTALL)
                            if match_cn:
                                g2b_data[m]["CreditNotes_IGST"] = clean_val(match_cn.group(1))
                                g2b_data[m]["CreditNotes_CGST"] = clean_val(match_cn.group(2))
                                g2b_data[m]["CreditNotes_SGST"] = clean_val(match_cn.group(3))
                                g2b_data[m]["CreditNotes_CESS"] = clean_val(match_cn.group(4))

                            # Table 4 - ITC Not Available (Searched strictly within t4_block)
                            if t4_block:
                                match_na = re.search(r"All other ITC.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", t4_block, re.IGNORECASE | re.DOTALL)
                                if not match_na:
                                    match_na = re.search(r"(\d[\d,]*\.\d{2})\s+(\d[\d,]*\.\d{2})\s+(\d[\d,]*\.\d{2})\s+(\d[\d,]*\.\d{2})", t4_block)
                                if match_na:
                                    g2b_data[m]["NotAvailable_IGST"] = clean_val(match_na.group(1))
                                    g2b_data[m]["NotAvailable_CGST"] = clean_val(match_na.group(2))
                                    g2b_data[m]["NotAvailable_SGST"] = clean_val(match_na.group(3))
                                    g2b_data[m]["NotAvailable_CESS"] = clean_val(match_na.group(4))
                        except Exception as e:
                            pass

                # Parse GSTR-3B PDFs
                for f_path in glob.glob(os.path.join(tmpdir, "GSTR3B_*.pdf")):
                    m = get_month_key_from_filename(os.path.basename(f_path))
                    if m:
                        g3b_data[m] = {
                            "T3A_Taxable": 0.0, "T3A_IGST": 0.0, "T3A_CGST": 0.0, "T3A_SGST": 0.0, "T3A_CESS": 0.0,
                            "T3D_Taxable": 0.0, "T3D_IGST": 0.0, "T3D_CGST": 0.0, "T3D_SGST": 0.0, "T3D_CESS": 0.0,
                            "Eligible_IGST": 0.0, "Eligible_CGST": 0.0, "Eligible_SGST": 0.0, "Eligible_CESS": 0.0,
                            "RC_IGST": 0.0, "RC_CGST": 0.0, "RC_SGST": 0.0, "RC_CESS": 0.0,
                            "Rev_IGST": 0.0, "Rev_CGST": 0.0, "Rev_SGST": 0.0, "Rev_CESS": 0.0,
                            "Net_IGST": 0.0, "Net_CGST": 0.0, "Net_SGST": 0.0, "Net_CESS": 0.0,
                            "Itc_Other_IGST": 0.0, "Itc_Other_CGST": 0.0, "Itc_Other_SGST": 0.0,
                            "Cash_Other_CGST": 0.0, "Cash_Other_SGST": 0.0,
                            "Cash_RC_IGST": 0.0, "Cash_RC_CGST": 0.0, "Cash_RC_SGST": 0.0,
                            "Days_Late": 0.0, "Filing_Date": "-"
                        }
                        try:
                            with pdfplumber.open(f_path) as pdf:
                                text = "\n".join([page.extract_text() or "" for page in pdf.pages])

                            filing_match = re.search(r"(?:Date of Filing|Date of ARN)\s*([\d\-/]+)", text, re.IGNORECASE)
                            if filing_match:
                                f_date_str = filing_match.group(1).strip()
                                g3b_data[m]["Filing_Date"] = f_date_str
                                due_date_str = due_dates_map.get(m, f"20/05/{f_date_str[-4:]}")
                                try:
                                    f_dt = pd.to_datetime(f_date_str, format='%d/%m/%Y')
                                    d_dt = pd.to_datetime(due_date_str, format='%d/%m/%Y')
                                    diff_days = (f_dt - d_dt).days
                                    g3b_data[m]["Days_Late"] = float(max(0, diff_days))
                                except Exception:
                                    g3b_data[m]["Days_Late"] = 0.0

                            m_31a = re.search(r"\(a\)\s+Outward taxable supplies.*?([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if m_31a:
                                g3b_data[m]["T3A_Taxable"] = clean_val(m_31a.group(1))
                                g3b_data[m]["T3A_IGST"] = clean_val(m_31a.group(2))
                                g3b_data[m]["T3A_CGST"] = clean_val(m_31a.group(3))
                                g3b_data[m]["T3A_SGST"] = clean_val(m_31a.group(4))
                                g3b_data[m]["T3A_CESS"] = clean_val(m_31a.group(5))

                            m_31d = re.search(r"\(d\)\s+Inward supplies \(liable to reverse charge\).*?([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if m_31d:
                                g3b_data[m]["T3D_Taxable"] = clean_val(m_31d.group(1))
                                g3b_data[m]["T3D_IGST"] = clean_val(m_31d.group(2))
                                g3b_data[m]["T3D_CGST"] = clean_val(m_31d.group(3))
                                g3b_data[m]["T3D_SGST"] = clean_val(m_31d.group(4))
                                g3b_data[m]["T3D_CESS"] = clean_val(m_31d.group(5))

                            rc_match = re.search(r"\(3\)\s+Inward supplies liable to reverse charge.*?([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if rc_match:
                                g3b_data[m]["RC_IGST"] = clean_val(rc_match.group(1))
                                g3b_data[m]["RC_CGST"] = clean_val(rc_match.group(2))
                                g3b_data[m]["RC_SGST"] = clean_val(rc_match.group(3))
                                g3b_data[m]["RC_CESS"] = clean_val(rc_match.group(4))

                            rev_match = re.search(r"\(1\)\s+As per rules 38,42.*?([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if rev_match:
                                g3b_data[m]["Rev_IGST"] = clean_val(rev_match.group(1))
                                g3b_data[m]["Rev_CGST"] = clean_val(rev_match.group(2))
                                g3b_data[m]["Rev_SGST"] = clean_val(rev_match.group(3))
                                g3b_data[m]["Rev_CESS"] = clean_val(rev_match.group(4))

                            net_match = re.search(r"C\.\s+Net ITC available.*?([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if net_match:
                                g3b_data[m]["Net_IGST"] = clean_val(net_match.group(1))
                                g3b_data[m]["Net_CGST"] = clean_val(net_match.group(2))
                                g3b_data[m]["Net_SGST"] = clean_val(net_match.group(3))
                                g3b_data[m]["Net_CESS"] = clean_val(net_match.group(4))

                            itc_match = re.search(r"All other ITC.*?([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", text, re.DOTALL | re.IGNORECASE)
                            if itc_match:
                                g3b_data[m]["Eligible_IGST"] = clean_val(itc_match.group(1))
                                g3b_data[m]["Eligible_CGST"] = clean_val(itc_match.group(2))
                                g3b_data[m]["Eligible_SGST"] = clean_val(itc_match.group(3))
                                g3b_data[m]["Eligible_CESS"] = clean_val(itc_match.group(4))

                            m_61 = re.search(r"6\.1\s+Payment of tax.*?(?=Breakup of tax liability)", text, re.DOTALL | re.IGNORECASE)
                            if m_61:
                                block_61 = m_61.group(0)
                                t6_central = re.search(r"Central\s+tax\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+|-)\s+([\d,]+\.\d+|-)\s+([\d,]+\.\d+|-)\s+([\d,]+\.\d+)", block_61)
                                if t6_central:
                                    g3b_data[m]["Itc_Other_IGST"] = clean_val(t6_central.group(2))
                                    g3b_data[m]["Itc_Other_CGST"] = clean_val(t6_central.group(3))
                                    g3b_data[m]["Cash_Other_CGST"] = clean_val(t6_central.group(6))

                                t6_state = re.search(r"State/UT\s+tax\s+([\d,]+\.\d+)\s+([\d,]+\.\d+|-)\s+([\d,]+\.\d+|-)\s+([\d,]+\.\d+|-)\s+([\d,]+\.\d+|-)\s+([\d,]+\.\d+)", block_61)
                                if t6_state:
                                    g3b_data[m]["Itc_Other_SGST"] = clean_val(t6_state.group(4))
                                    g3b_data[m]["Cash_Other_SGST"] = clean_val(t6_state.group(6))

                                rc_block = re.search(r"\(B\)\s+Reverse charge.*", block_61, re.DOTALL)
                                if rc_block:
                                    rc_b = rc_block.group(0)
                                    m_rc_igst = re.search(r"Integrated\s+tax\s+([\d,]+\.\d+)\s+.*?\s+([\d,]+\.\d+)", rc_b, re.DOTALL)
                                    if m_rc_igst: g3b_data[m]["Cash_RC_IGST"] = clean_val(m_rc_igst.group(2))
                                    
                                    m_rc_cgst = re.search(r"Central\s+tax\s+([\d,]+\.\d+)\s+.*?\s+([\d,]+\.\d+)", rc_b, re.DOTALL)
                                    if m_rc_cgst: g3b_data[m]["Cash_RC_CGST"] = clean_val(m_rc_cgst.group(2))

                                    m_rc_sgst = re.search(r"State/UT\s+tax\s+([\d,]+\.\d+)\s+.*?\s+([\d,]+\.\d+)", rc_b, re.DOTALL)
                                    if m_rc_sgst: g3b_data[m]["Cash_RC_SGST"] = clean_val(m_rc_sgst.group(2))

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

                    for m in MONTHS:
                        if m in df.columns:
                            df[m] = None
                            df[m] = df[m].astype(object)

                    if sheet_name == 'GSTR-3B Summary (12M)':
                        section_col = next((c for c in df.columns if str(c).strip().lower() in ['table section', 'section', 'tablesection', 'table_section']), heading_col)
                        details_col = next((c for c in df.columns if str(c).strip().lower() in ['details / field', 'details', 'field', 'details/field', 'details_field']), heading_col)
                        for idx, row in df.iterrows():
                            section = str(row[section_col]).strip().lower()
                            field = str(row[details_col]).strip().lower()
                            for m in MONTHS:
                                if m in g3b_data:
                                    val = None
                                    if 'due date' in section or 'due date' in field:
                                        val = due_dates_map.get(m, "-")
                                    elif 'filing date' in section or 'filing date' in field:
                                        val = g3b_data[m]["Filing_Date"]
                                    elif 'no. of days late' in field or 'days late' in section:
                                        val = g3b_data[m]["Days_Late"]
                                    elif '3.1 (a)' in section or ('3.1' in section and 'taxable supplies' in section and 'other than' in section):
                                        if 'taxable' in field: val = g3b_data[m]["T3A_Taxable"]
                                        elif 'igst' in field: val = g3b_data[m]["T3A_IGST"]
                                        elif 'cgst' in field: val = g3b_data[m]["T3A_CGST"]
                                        elif 'sgst' in field: val = g3b_data[m]["T3A_SGST"]
                                        elif 'cess' in field: val = g3b_data[m]["T3A_CESS"]
                                    elif '3.1 (d)' in section or ('3.1' in section and 'reverse charge' in section):
                                        if 'taxable' in field: val = g3b_data[m]["T3D_Taxable"]
                                        elif 'igst' in field: val = g3b_data[m]["T3D_IGST"]
                                        elif 'cgst' in field: val = g3b_data[m]["T3D_CGST"]
                                        elif 'sgst' in field: val = g3b_data[m]["T3D_SGST"]
                                        elif 'cess' in field: val = g3b_data[m]["T3D_CESS"]
                                    elif 'inward supplies liable to reverse charge' in section or '4(a)(3)' in section:
                                        if 'igst' in field: val = g3b_data[m]["RC_IGST"]
                                        elif 'cgst' in field: val = g3b_data[m]["RC_CGST"]
                                        elif 'sgst' in field: val = g3b_data[m]["RC_SGST"]
                                        elif 'cess' in field: val = g3b_data[m]["RC_CESS"]
                                    elif 'as per rules 38' in section or '4(b)(1)' in section:
                                        if 'igst' in field: val = g3b_data[m]["Rev_IGST"]
                                        elif 'cgst' in field: val = g3b_data[m]["Rev_CGST"]
                                        elif 'sgst' in field: val = g3b_data[m]["Rev_SGST"]
                                        elif 'cess' in field: val = g3b_data[m]["Rev_CESS"]
                                    elif 'net itc available' in section or '4(c)' in section:
                                        if 'igst' in field: val = g3b_data[m]["Net_IGST"]
                                        elif 'cgst' in field: val = g3b_data[m]["Net_CGST"]
                                        elif 'sgst' in field: val = g3b_data[m]["Net_SGST"]
                                        elif 'cess' in field: val = g3b_data[m]["Net_CESS"]
                                    elif 'other than reverse charge' in section and 'itc' in field:
                                        if 'igst' in field: val = g3b_data[m]["Itc_Other_IGST"]
                                        elif 'cgst' in field: val = g3b_data[m]["Itc_Other_CGST"]
                                        elif 'sgst' in field: val = g3b_data[m]["Itc_Other_SGST"]
                                    elif 'other than reverse charge' in section and 'cash' in field:
                                        if 'cgst' in field: val = g3b_data[m]["Cash_Other_CGST"]
                                        elif 'sgst' in field: val = g3b_data[m]["Cash_Other_SGST"]
                                    elif 'reverse charge' in section and 'cash' in field:
                                        if 'igst' in field: val = g3b_data[m]["Cash_RC_IGST"]
                                        elif 'cgst' in field: val = g3b_data[m]["Cash_RC_CGST"]
                                        elif 'sgst' in field: val = g3b_data[m]["Cash_RC_SGST"]
                                    
                                    if val is not None:
                                        df.at[idx, m] = val

                        numeric_only = df[MONTHS].apply(pd.to_numeric, errors='coerce').fillna(0.0)
                        df['FY TOTAL'] = numeric_only.sum(axis=1)

                    elif sheet_name == 'GSTR-1 Summary (12M)':
                        section_col = next((c for c in df.columns if str(c).strip().lower() in ['table section', 'section', 'tablesection', 'table_section']), heading_col)
                        details_col = next((c for c in df.columns if str(c).strip().lower() in ['details / field', 'details', 'field', 'details/field', 'details_field']), heading_col)
                        for idx, row in df.iterrows():
                            section = str(row[section_col]).strip().lower()
                            field = str(row[details_col]).strip().lower()
                            for m in MONTHS:
                                if m in g1_data:
                                    val = None
                                    if '4a' in section or 'b2b regular' in section:
                                        if 'taxable' in field: val = g1_data[m]["B2B_Taxable"]
                                        elif 'igst' in field: val = g1_data[m]["B2B_IGST"]
                                        elif 'cgst' in field: val = g1_data[m]["B2B_CGST"]
                                        elif 'sgst' in field: val = g1_data[m]["B2B_SGST"]
                                        elif 'cess' in field: val = g1_data[m]["B2B_CESS"]
                                    elif '4b' in section or 'b2b reverse charge' in section:
                                        if 'taxable' in field: val = g1_data[m]["RC_Taxable"]
                                        elif 'igst' in field: val = g1_data[m]["RC_IGST"]
                                        elif 'cgst' in field: val = g1_data[m]["RC_CGST"]
                                        elif 'sgst' in field: val = g1_data[m]["RC_SGST"]
                                        elif 'cess' in field: val = g1_data[m]["RC_CESS"]
                                    elif '7' in section or 'b2cs' in section:
                                        if 'net' in field or 'taxable' in field: val = g1_data[m]["B2CS_Net"]
                                        elif 'igst' in field: val = g1_data[m]["B2CS_IGST"]
                                        elif 'cgst' in field: val = g1_data[m]["B2CS_CGST"]
                                        elif 'sgst' in field: val = g1_data[m]["B2CS_SGST"]
                                        elif 'cess' in field: val = g1_data[m]["B2CS_CESS"]
                                    
                                    if val is not None:
                                        df.at[idx, m] = val

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
                                    val = None
                                    if 'all other' in section or '4(a)(5)' in section or '3. itc available - part a' in section:
                                        if 'igst' in field or 'integrated' in field: val = g2b_data[m]["PartA_IGST"]
                                        elif 'cgst' in field or 'central' in field: val = g2b_data[m]["PartA_CGST"]
                                        elif 'sgst' in field or 'state' in field: val = g2b_data[m]["PartA_SGST"]
                                        elif 'cess' in field: val = g2b_data[m]["PartA_CESS"]
                                    elif 'credit notes' in section or 'part b' in section:
                                        if 'igst' in field or 'integrated' in field: val = g2b_data[m]["CreditNotes_IGST"]
                                        elif 'cgst' in field or 'central' in field: val = g2b_data[m]["CreditNotes_CGST"]
                                        elif 'sgst' in field or 'state' in field: val = g2b_data[m]["CreditNotes_SGST"]
                                        elif 'cess' in field: val = g2b_data[m]["CreditNotes_CESS"]
                                    elif 'not available' in section or '4. itc not available' in section:
                                        if 'igst' in field or 'integrated' in field: val = g2b_data[m]["NotAvailable_IGST"]
                                        elif 'cgst' in field or 'central' in field: val = g2b_data[m]["NotAvailable_CGST"]
                                        elif 'sgst' in field or 'state' in field: val = g2b_data[m]["NotAvailable_SGST"]
                                        elif 'cess' in field: val = g2b_data[m]["NotAvailable_CESS"]
                                    
                                    if val is not None:
                                        df.at[idx, m] = val

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

                interest_rows = []
                for m in MONTHS:
                    days_late = g3b_data.get(m, {}).get("Days_Late", 0.0)
                    cash_tax = g3b_data.get(m, {}).get("Cash_RC_CGST", 0.0) + g3b_data.get(m, {}).get("Cash_RC_SGST", 0.0)
                    calc_interest = round(cash_tax * 0.18 * (days_late / 365.0), 2) if days_late > 0 else 0.0
                    interest_rows.append({
                        "Period": m, "Days Delayed": days_late, "Net Cash Tax Paid": cash_tax,
                        "Interest Payable (18%)": calc_interest, "Status": "Action Needed" if calc_interest > 0 else "Satisfactory"
                    })
                output_sheets["10. Delayed Interest Audit"] = pd.DataFrame(interest_rows)

                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    for sheet_name, df_data in output_sheets.items():
                        df_data.to_excel(writer, sheet_name=sheet_name, index=False)

                st.success("🎉 GST_Audit_Working_File.xlsx successfully generated!")

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Download GST_Audit_Working_File.xlsx",
                        data=f,
                        file_name="GST_Audit_Working_File.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )