import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from PyPDF2 import PdfReader
import tempfile
import io

st.title("REDPAY Reconciliation Tool")

# ===== FILE UPLOAD =====
excel_file = st.file_uploader("Upload Excel/CSV file", type=["xlsx", "csv"])
pdf_file = st.file_uploader("Upload PDF statement", type=["pdf"])

# ===== RUN BUTTON =====
if st.button("Run Reconciliation"):

    if excel_file and pdf_file:

        st.info("Processing...")

        # ===== LOAD FILE =====
        if excel_file.name.endswith(".csv"):
            df = pd.read_csv(excel_file)
        else:
            df = pd.read_excel(excel_file)

        # ===== COLUMN SETUP =====
        action_col = df.columns[0]
        amount_col = df.columns[1]
        txn_col = df.columns[11]

        # ===== FILTER DEBITS =====
        debits = df[df[action_col].astype(str).str.strip().str.upper() == "DEBIT"]

        amounts = debits[amount_col].dropna().tolist()
        txn_ids = debits[txn_col].dropna().astype(str).str.strip().tolist()

        # ===== READ PDF =====
        pdf_file.seek(0)
        reader = PdfReader(pdf_file)

        pdf_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                pdf_text += text

        pdf_text_clean = pdf_text.replace(",", "").replace(".00", "")

        # ===== MATCH =====
        found_amounts = []
        found_txns = []

        for amt in amounts:
            try:
                clean = str(int(float(amt)))
                if clean in pdf_text_clean:
                    found_amounts.append(amt)
            except:
                pass

        for txn in txn_ids:
            if txn in pdf_text:
                found_txns.append(txn)

        # ===== REPORT DATA (FIXED STRUCTURE) =====

        # FOUND
        found_amount_df = pd.DataFrame({
            "Matched Amounts": sorted(set(int(float(a)) for a in found_amounts))
        })

        found_txn_df = pd.DataFrame({
            "Matched Transaction IDs": sorted(set(found_txns))
        })

        # MISSING
        all_amounts = [int(float(a)) for a in amounts if pd.notna(a)]
