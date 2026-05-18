import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from PyPDF2 import PdfReader
import tempfile

st.title("DHikrulahi Reconciliation Project")

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
        amount_col = df.columns[1]   # Column 2
        txn_col = df.columns[11]     # Column 12

        # ===== FILTER DEBITS =====
        debits = df[df[action_col].astype(str).str.strip().str.upper() == "DEBIT"]

        amounts = debits[amount_col].dropna().tolist()
        txn_ids = debits[txn_col].dropna().astype(str).str.strip().tolist()

        # ===== READ PDF TEXT =====
        pdf_file.seek(0)  # IMPORTANT
        reader = PdfReader(pdf_file)

        pdf_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                pdf_text += text

        # Normalize for matching
        pdf_text_clean = pdf_text.replace(",", "").replace(".00", "")

        # ===== MATCH =====
        found_amounts = []
        found_txns = []

        for amt in amounts:
            try:
                clean = str(int(float(amt)))
                if clean in pdf_text_clean:
                    found_amounts.append(amt)  # keep original
            except:
                pass

        for txn in txn_ids:
            if txn in pdf_text:
                found_txns.append(txn)

        # ===== CREATE REPORT =====
        report_text = "RECONCILIATION REPORT\n\n"

        report_text += "FOUND AMOUNTS\n---------------------\n"
        for a in sorted(set(found_amounts)):
            report_text += str(int(float(a))) + "\n"

        report_text += "\nMISSING AMOUNTS\n---------------------\n"
        all_amounts = [str(int(float(a))) for a in amounts if pd.notna(a)]
        for a in sorted(set(all_amounts) - set(str(int(float(x))) for x in found_amounts)):
            report_text += a + "\n"

        report_text += "\nFOUND TRANSACTION IDs\n---------------------\n"
        for t in sorted(set(found_txns)):
            report_text += t + "\n"

        report_text += "\nMISSING TRANSACTION IDs\n---------------------\n"
        for t in sorted(set(txn_ids) - set(found_txns)):
            report_text += t + "\n"

        # ===== SAVE PDF TEMP =====
        pdf_file.seek(0)
        pdf_bytes = pdf_file.read()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(pdf_bytes)
            temp_pdf.flush()
            temp_pdf_path = temp_pdf.name

        doc = fitz.open(temp_pdf_path)

        # ===== HIGHLIGHT ONLY MATCHES =====
        for page in doc:

            # highlight matched amounts
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
                            annot.set_colors(stroke=(1, 1, 0))  # yellow
                            annot.update()
                except:
                    pass

            # highlight matched transaction IDs
            for txn in set(found_txns):
                for inst in page.search_for(txn):
                    annot = page.add_highlight_annot(inst)
                    annot.set_colors(stroke=(0, 1, 0))  # green
                    annot.update()

        # ===== SAVE OUTPUT PDF =====
        output_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
        doc.save(output_pdf, deflate=True, garbage=4)

        # ===== DISPLAY =====
        st.success("Done!")

        st.write("Matched Amounts:", len(set(found_amounts)))
        st.write("Matched Transactions:", len(set(found_txns)))

        # ===== DOWNLOAD REPORT =====
        st.download_button(
            label="Download Report",
            data=report_text,
            file_name="reconciliation_report.txt",
            mime="text/plain"
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
