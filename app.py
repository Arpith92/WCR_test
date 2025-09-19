import os
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import streamlit as st
import zipfile
import io
from docx2pdf import convert   # ✅ Only works on Windows/Mac with Word

TEMPLATE_DOC = "sample.docx"
OUT_DIR = "Result"
os.makedirs(OUT_DIR, exist_ok=True)

def _safe(x):
    if pd.isna(x) or x == "":
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")
    try:
        num = float(x)
        return f"{num:.2f}"
    except (ValueError, TypeError):
        return str(x).strip()

st.title("📑 Automated WCR Generator")

uploaded_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    rename_map = {
        "wo no": "wo_no", "wo_no": "wo_no",
        "wo date": "wo_date", "wo_date": "wo_date",
        "wo des": "wo_des", "wo_des": "wo_des",
        "location_code": "Location_code", "Location_code": "Location_code",
        "customername_code": "customername_code",
        "capacity_code": "Capacity_code", "Capacity_code": "Capacity_code",
        "site_incharge": "site_incharge",
        "Scada_incharge": "Scada_incharge",
        "Re_date": "Re_date",
        "Site_Name": "Site_Name",
        "Line_1_Workstatus":"Line_1_Workstatus",
        "Line_2_Workstatus":"Line_2_Workstatus",
        "Payment Terms": "Payment_Terms"
    }
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

    generated_word, generated_pdf = [], []

    for i, row in df.iterrows():
        context = {col: _safe(row[col]) for col in df.columns}

        # Word generation
        doc = DocxTemplate(TEMPLATE_DOC)
        doc.render(context)
        wo = context.get("wo_no", "") or f"Row{i+1}"
        word_path = os.path.join(OUT_DIR, f"WCR_{wo}.docx")
        doc.save(word_path)
        generated_word.append(word_path)

        # PDF conversion (Word → PDF)
        pdf_path = os.path.join(OUT_DIR, f"WCR_{wo}.pdf")
        try:
            convert(word_path, pdf_path)   # uses MS Word in background
            generated_pdf.append(pdf_path)
        except Exception as e:
            st.error(f"PDF conversion failed for {word_path}: {e}")

    # ---- ZIP Word ----
    zip_word = io.BytesIO()
    with zipfile.ZipFile(zip_word, "w") as zipf:
        for file in generated_word:
            zipf.write(file, arcname=os.path.basename(file))
    zip_word.seek(0)
    st.download_button("⬇️ Download All WCR Files (Word ZIP)", zip_word, "WCR_Word_Files.zip", "application/zip")

    # ---- ZIP PDF ----
    if generated_pdf:
        zip_pdf = io.BytesIO()
        with zipfile.ZipFile(zip_pdf, "w") as zipf:
            for file in generated_pdf:
                zipf.write(file, arcname=os.path.basename(file))
        zip_pdf.seek(0)
        st.download_button("⬇️ Download All WCR Files (PDF ZIP)", zip_pdf, "WCR_PDF_Files.zip", "application/zip")
