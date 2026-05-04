import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
from docxtpl import DocxTemplate
import io
import zipfile
from num2words import num2words
import os

# ---------------- TEMPLATE STORAGE ----------------
TEMPLATE_DIR = "templates"
os.makedirs(TEMPLATE_DIR, exist_ok=True)

st.title("📄 Data Input → File Generator")
st.write("Enter data manually or upload an Excel file")

# ---------------- SIDEBAR ----------------
option = st.sidebar.radio("Choose input method", ["Manual Form", "Upload Excel"])

st.sidebar.header("Template Settings")

uploaded_templates = st.sidebar.file_uploader(
    "Upload Word Templates (.docx)",
    type=["docx"],
    accept_multiple_files=True
)

# SAVE TEMPLATES
if uploaded_templates:
    for uploaded_file in uploaded_templates:
        file_path = os.path.join(TEMPLATE_DIR, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
    st.sidebar.success("Templates saved permanently!")

# LOAD SAVED TEMPLATES
saved_templates = [
    f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".docx")
]

selected_templates = st.sidebar.multiselect(
    "Select Templates to Use",
    saved_templates
)

# ---------------- DATA ----------------
data = {}
batch_data = []

# ---------------- MANUAL INPUT ----------------
if option == "Manual Form":
    col1, col2 = st.columns(2)

    with col1:
        data["Company"] = st.text_input("Company Name")
        data["Company Address"] = st.text_input("Company Address")
        data["Subcontractor Name"] = st.text_input("Subcontractor Name")
        data["Subcontractor Address"] = st.text_input("Subcontractor Address")
        data["POC Name"] = st.text_input("Point of Contact Name")
        data["POC Phone"] = st.text_input("POC Phone")
        data["POC Email"] = st.text_input("POC Email")

    with col2:
        data["Project Name"] = st.text_input("Project Name")
        data["Project Number"] = st.text_input("Project Number")
        data["Project Address"] = st.text_input("Project Address")
        data["Payout"] = st.number_input("Payout ($)", min_value=0.0)
        data["Completion Date"] = st.date_input("Project Completion Date")

    data["Scope"] = st.text_area("Scope of Work", height=150)

# ---------------- EXCEL INPUT ----------------
else:
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.dataframe(df, use_container_width=True)

        if not df.empty:
            selected_rows = st.multiselect(
                "Select rows to process",
                options=list(df.index),
                default=[0]
            )
            batch_data = df.loc[selected_rows].to_dict(orient="records")
            st.success(f"{len(batch_data)} row(s) selected!")

# ---------------- GENERATE ----------------
if st.button("Generate Files", type="primary"):

    if option == "Upload Excel":
        if not batch_data:
            st.error("No rows selected!")
            st.stop()
        records = batch_data
    else:
        if not data:
            st.error("No data to generate!")
            st.stop()
        records = [data]

    tab1, tab2, tab3, tab4 = st.tabs(["CSV", "PDF", "Excel", "Word"])

    # -------- CSV --------
    with tab1:
        csv = pd.DataFrame(records).to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download CSV", csv,
                           f"report_{datetime.now().strftime('%Y%m%d')}.csv",
                           "text/csv")

    # -------- PDF --------
    with tab2:
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for i, record in enumerate(records):
                pdf = FPDF()
                pdf.add_page()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.set_font("Arial", size=12)

                payout_value = record.get("Payout", 0)

                try:
                    payout_value = float(payout_value)
                    dollars = int(payout_value)
                    cents = int(round((payout_value - dollars) * 100))

                    dollars_words = num2words(dollars).replace(',', '').title()
                    payout_words = f"{dollars_words} Dollars And {cents:02d}/100"

                except:
                    payout_words = "Invalid Amount"

                contract_text = f"""
SUBCONTRACTOR AGREEMENT

Company: {record.get("Company", "")}
Project: {record.get("Project Name", "")}

Payout: ${payout_value:,.2f}
Amount in Words: {payout_words}
"""

                pdf.multi_cell(0, 8, contract_text.strip())
                pdf_bytes = bytes(pdf.output(dest='S'))

                filename = record.get("Company", f"contract_{i+1}").replace(" ", "_") + ".pdf"
                zip_file.writestr(filename, pdf_bytes)

        zip_buffer.seek(0)
        st.download_button("📦 Download PDFs (ZIP)", zip_buffer,
                           "contracts_pdf.zip", "application/zip")

    # -------- EXCEL --------
    with tab3:
        output = io.BytesIO()
        pd.DataFrame(records).to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button("📥 Download Excel", output,
                           f"report_{datetime.now().strftime('%Y%m%d')}.xlsx")

    # -------- WORD --------
    with tab4:

        if not selected_templates:
            st.error("Please select at least one saved template in the sidebar")
            st.stop()

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:

            for i, record in enumerate(records):

                payout_value = record.get("Payout", 0)

                try:
                    payout_value = float(payout_value)
                    dollars = int(payout_value)
                    cents = int(round((payout_value - dollars) * 100))

                    dollars_words = num2words(dollars).replace(',', '').title()
                    payout_words = f"{dollars_words} Dollars And {cents:02d}/100"

                except:
                    payout_words = "Invalid Amount"

                context = {
                    "Agreement_Date": datetime.now().strftime('%Y-%m-%d'),
                    "Company": record.get("Company", ""),
                    "Company_Address": record.get("Company Address", ""),
                    "Subcontractor_Name": record.get("Subcontractor Name", ""),
                    "Subcontractor_Address": record.get("Subcontractor Address", ""),
                    "POC_Name": record.get("POC Name", ""),
                    "POC_Phone": record.get("POC Phone", ""),
                    "POC_Email": record.get("POC Email", ""),
                    "Project_Name": record.get("Project Name", ""),
                    "Project_Number": record.get("Project Number", ""),
                    "Project_Address": record.get("Project Address", ""),
                    "Scope": record.get("Scope", ""),
                    "Payout": f"{payout_value:,.2f}",
                    "PayoutWords": payout_words,
                    "Completion_Date": str(record.get("Completion Date", ""))
                }

                for template_name in selected_templates:
                    template_path = os.path.join(TEMPLATE_DIR, template_name)

                    doc = DocxTemplate(template_path)
                    doc.render(context)

                    file_stream = io.BytesIO()
                    doc.save(file_stream)

                    company = record.get("Company", f"company_{i+1}").replace(" ", "_")
                    clean_template = template_name.replace(".docx", "").replace(" ", "_")

                    filename = f"{company}_{clean_template}.docx"
                    zip_file.writestr(filename, file_stream.getvalue())

        zip_buffer.seek(0)

        st.download_button(
            "📄 Download Word Docs (ZIP)",
            zip_buffer,
            "contracts_word.zip",
            "application/zip"
        )