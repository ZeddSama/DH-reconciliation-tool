import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
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

        # ===== READ PDF TEXT =====
        pdf_file.seek(0)
        reader = PdfReader(pdf_file)

        pdf_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                pdf_text += text

        pdf_text_clean = pdf_text.replace(",", "").replace(".00", "")

        # ===== MATCH (PAIRED LOGIC ✅) =====
        matched_rows = []
        missing_rows = []

        for _, row in debits.iterrows():
            try:
                amt = row[amount_col]
                txn = str(row[txn_col]).strip()

                amt_clean = str(int(float(amt)))

                amount_found = amt_clean in pdf_text_clean
                txn_found = txn in pdf_text

                if amount_found and txn_found:
                    matched_rows.append({
                        "Amount": int(float(amt)),
                        "Transaction ID": txn
                    })
                else:
                    missing_rows.append({
                        "Amount": int(float(amt)),
                        "Transaction ID": txn
                    })

            except:
                continue

        # ===== CREATE DATAFRAMES =====
        matched_df = pd.DataFrame(matched_rows)
        missing_df = pd.DataFrame(missing_rows)

        # ===== CREATE EXCEL OUTPUT =====
        output_excel = io.BytesIO()

        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            matched_df.to_excel(writer, sheet_name="MATCHED", index=False)
            missing_df.to_excel(writer, sheet_name="MISSING", index=False)

        output_excel.seek(0)

        # ===== DISPLAY =====
        st.success("Reconciliation Complete ✅")

        st.write("Confirmed Settlements:", len(matched_df))
        st.write("Missing Settlements:", len(missing_df))

        # ===== DOWNLOAD EXCEL =====
        st.download_button(
            label="Download Excel Report",
            data=output_excel,
            file_name="reconciliation_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Please upload both files.")
