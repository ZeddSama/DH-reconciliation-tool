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

        # ===== READ PDF TEXT =====
        pdf_file.seek(0)
        reader = PdfReader(pdf_file)

        pdf_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                pdf_text += text

        # Normalize text
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

        # ===== CREATE DATAFRAMES FOR REPORT =====

        # FOUND
        found_df = pd.DataFrame({
            "Matched Amount": [int(float(a)) for a in found_amounts],
        })

        if found_txns:
            found_df["Matched Transaction ID"] = list(set(found_txns))

        # MISSING
        all_amounts = [int(float(a)) for a in amounts if pd.notna(a)]
        missing_amounts = list(set(all_amounts) - set(found_df["Matched Amount"]))

        missing_txns = list(set(txn_ids) - set(found_txns))

        missing_df = pd.DataFrame({
            "Missing Amount": missing_amounts,
        })

        if missing_txns:
            missing_df["Missing Transaction ID"] = missing_txns

        # ===== CREATE EXCEL OUTPUT =====
        output_excel = io.BytesIO()

        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            found_df.to_excel(writer, sheet_name="FOUND", index=False)
            missing_df.to_excel(writer, sheet_name="MISSING", index=False)

        output_excel.seek(0)

        # ===== SAVE PDF TEMP =====
        pdf_file.seek(0)
        pdf_bytes = pdf_file.read()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(pdf_bytes)
            temp_pdf.flush()
            temp_pdf_path = temp_pdf.name

        doc = fitz.open(temp_pdf_path)

        # ===== HIGHLIGHT MATCHES =====
        for page in doc:

            # highlight amounts
            for amt in set(found_amounts):
                try:
                    val = float(amt)
                    formats = [
                        f"{val:,.2f}",
                        f"{int(val):,}"
                    ]

                    for f in formats:
                        for inst in page.search_for(f):
                            annot = page.add_highlight_annot(inst)
                            annot.set_colors(stroke=(1, 1, 0))
                            annot.update()
                except:
                    pass

            # highlight transactions
            for txn in set(found_txns):
                for inst in page.search_for(txn):
                    annot = page.add_highlight_annot(inst)
                    annot.set_colors(stroke=(0, 1, 0))
                    annot.update()

        # ===== SAVE OUTPUT PDF =====
        output_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
        doc.save(output_pdf, deflate=True, garbage=4)

        # ===== DISPLAY =====
        st.success("Reconciliation Complete ✅")

        st.write("Matched Amounts:", len(set(found_amounts)))
        st.write("Matched Transactions:", len(set(found_txns)))

        # ===== DOWNLOAD EXCEL REPORT =====
        st.download_button(
            label="Download Excel Report",
            data=output_excel,
            file_name="reconciliation_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ===== DOWNLOAD PDF =====
        with open(output_pdf, "rb") as f:
            st.download_button(
                label="Download Highlighted PDF",
                data=f,
                file_name="highlighted_statement.pdf",
                mime="application/pdf"
            )

    else:
        st.warning("Please upload both files.")
