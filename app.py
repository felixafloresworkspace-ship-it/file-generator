import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
from docxtpl import DocxTemplate
import io
import zipfile
from num2words import num2words
import os
import re

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(layout="wide")

# =========================================================
# TEMPLATE STORAGE
# =========================================================
TEMPLATE_DIR = "templates"

if os.path.exists(TEMPLATE_DIR) and not os.path.isdir(TEMPLATE_DIR):
    os.remove(TEMPLATE_DIR)

os.makedirs(TEMPLATE_DIR, exist_ok=True)

# =========================================================
# TITLE
# =========================================================
st.title("📄 Data Input → File Generator")
st.write("Enter data manually or upload an Excel file")

# =========================================================
# SESSION STATE
# =========================================================
if "selected_templates" not in st.session_state:
    st.session_state.selected_templates = []

# =========================================================
# HELPERS
# =========================================================
def clean_number_words(n):
    return num2words(n).replace(",", "").replace(" and ", " ").title()

def clean_text(val):
    """
    Prevent issues with special characters
    """
    if pd.isna(val):
        return ""

    val = str(val)

    replacements = {
        "&": "&",
        "–": "-",
        "—": "-",
        "“": '"',
        "”": '"',
        "’": "'"
    }

    for old, new in replacements.items():
        val = val.replace(old, new)

    return val.strip()

def clean_filename(text):
    """
    Remove invalid filename characters
    """
    text = str(text)
    return re.sub(r'[\\/*?:"<>|]', "", text)

def format_division(val):

    if val is None or val == "":
        return ""

    try:

        if isinstance(val, str) and "," in val:

            parts = [
                str(int(float(x.strip())))
                for x in val.split(",")
            ]

            return ", ".join(parts)

        return str(int(float(val)))

    except:
        return str(val)

def build_full_division(
    division,
    division_description
):

    division = clean_text(division)

    division_description = clean_text(
        division_description
    )

    if division and division_description:

        return (
            f"Division {division} - "
            f"{division_description}"
        )

    if division:
        return f"Division {division}"

    return ""

def build_payout_words(value):

    try:

        payout_value = float(value)

        dollars = int(payout_value)

        cents = int(
            round(
                (payout_value - dollars) * 100
            )
        )

        dollars_words = clean_number_words(
            dollars
        )

        cents_words = clean_number_words(
            cents
        )

        cents_label = (
            "Cent"
            if cents == 1
            else "Cents"
        )

        return (
            f"{dollars_words} Dollars And "
            f"{cents_words} "
            f"{cents_label} "
            f"({cents:02d}/100)"
        )

    except:
        return "Invalid Amount"

def get_val(record, *keys):

    for k in keys:

        if k in record and pd.notna(record[k]):

            return clean_text(record[k])

    return ""

# =========================================================
# SIDEBAR
# =========================================================
option = st.sidebar.radio(
    "Choose input method",
    ["Manual Form", "Upload Excel"]
)

st.sidebar.header("📂 Template Settings")

uploaded_templates = st.sidebar.file_uploader(
    "Upload Word Templates (.docx)",
    type=["docx"],
    accept_multiple_files=True
)

# =========================================================
# SAVE TEMPLATES
# =========================================================
if uploaded_templates:

    for uploaded_file in uploaded_templates:

        file_path = os.path.join(
            TEMPLATE_DIR,
            uploaded_file.name
        )

        if os.path.exists(file_path):

            st.sidebar.warning(
                f"{uploaded_file.name} "
                f"already exists and was overwritten."
            )

        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

    st.sidebar.success(
        "Templates saved successfully!"
    )

# =========================================================
# LOAD SAVED TEMPLATES
# =========================================================
saved_templates = []

if os.path.exists(TEMPLATE_DIR):

    saved_templates = sorted([
        f for f in os.listdir(TEMPLATE_DIR)
        if f.endswith(".docx")
    ])

# =========================================================
# DISPLAY SAVED TEMPLATES
# =========================================================
if saved_templates:

    st.sidebar.write("### Saved Templates")

    for t in saved_templates:
        st.sidebar.write(f"✅ {t}")

selected_templates = st.sidebar.multiselect(
    "Select Templates to Use",
    options=saved_templates,
    default=st.session_state.selected_templates
)

st.session_state.selected_templates = (
    selected_templates
)

# =========================================================
# DATA STORAGE
# =========================================================
data = {}
batch_data = []

# =========================================================
# MANUAL FORM INPUT
# =========================================================
if option == "Manual Form":

    col1, col2 = st.columns(2)

    with col1:

        data["Company"] = st.text_input(
            "Company Name"
        )

        data["Company Address"] = st.text_input(
            "Company Address"
        )

        data["Subcontractor Name"] = st.text_input(
            "Subcontractor Name"
        )

        data["Subcontractor Address"] = st.text_input(
            "Subcontractor Address"
        )

        data["Subcontractor City"] = st.text_input(
            "Subcontractor City"
        )

        data["POC Name"] = st.text_input(
            "Point of Contact Name"
        )

        data["POC Phone"] = st.text_input(
            "POC Phone"
        )

        data["POC Email"] = st.text_input(
            "POC Email"
        )

    with col2:

        data["Project Name"] = st.text_input(
            "Project Name"
        )

        data["Project Number"] = st.text_input(
            "Project Number"
        )

        data["Project Address"] = st.text_input(
            "Project Address"
        )

        data["Division"] = st.text_input(
            "Division Number"
        )

        data["Division Description"] = st.text_input(
            "Division Description"
        )

        data["Payout"] = st.number_input(
            "Payout ($)",
            min_value=0.0
        )

        data["Completion Date"] = st.date_input(
            "Project Completion Date"
        )

    data["Scope"] = st.text_area(
        "Scope of Work",
        height=150
    )

    data["Addendums"] = st.text_area(
        "Addendums"
    )

    data["alternate"] = st.text_area(
        "Alternates"
    )

# =========================================================
# EXCEL INPUT
# =========================================================
else:

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"]
    )

    if uploaded_file:

        df = pd.read_excel(uploaded_file)

        df.columns = (
            df.columns.str.strip()
        )

        # =====================================================
        # INSERT SELECT COLUMN FIRST
        # =====================================================
        if "Select" not in df.columns:
            df.insert(0, "Select", False)

        # =====================================================
        # DISPLAY TABLE
        # =====================================================
        st.subheader("📋 Uploaded Records")

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            hide_index=False,
            num_rows="dynamic"
        )

        # =====================================================
        # SELECTED ROWS
        # =====================================================
        selected_rows = edited_df[
            edited_df["Select"] == True
        ]

        if not selected_rows.empty:

            batch_data = selected_rows.drop(
                columns=["Select"]
            ).to_dict(
                orient="records"
            )

            st.success(
                f"{len(batch_data)} row(s) selected!"
            )

            # =================================================
            # SHOW SELECTED COMPANIES
            # =================================================
            st.write(
                "### ✅ Selected Companies"
            )

            selected_display = []

            for row in batch_data:

                project = row.get(
                    "Project Name",
                    ""
                )

                company = row.get(
                    "Company",
                    ""
                )

                selected_display.append(
                    f"{project} | {company}"
                )

            st.write(selected_display)

# =========================================================
# GENERATE FILES
# =========================================================
if st.button(
    "🚀 Generate Files",
    type="primary"
):

    # =====================================================
    # VALIDATION
    # =====================================================
    if option == "Upload Excel":

        if not batch_data:

            st.error(
                "No rows selected!"
            )

            st.stop()

        records = batch_data

    else:

        if not data:

            st.error(
                "No data to generate!"
            )

            st.stop()

        records = [data]

    # =====================================================
    # TABS
    # =====================================================
    tab1, tab2, tab3, tab4 = st.tabs([
        "CSV",
        "PDF",
        "Excel",
        "Word"
    ])

    # =====================================================
    # CSV EXPORT
    # =====================================================
    with tab1:

        csv = pd.DataFrame(records).to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            "📥 Download CSV",
            csv,
            f"report_"
            f"{datetime.now().strftime('%Y%m%d')}"
            f".csv"
        )

    # =====================================================
    # PDF EXPORT
    # =====================================================
    with tab2:

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            "w"
        ) as zip_file:

            for i, record in enumerate(records):

                pdf = FPDF()

                pdf.add_page()

                pdf.set_font(
                    "Arial",
                    size=12
                )

                payout_value = record.get(
                    "Payout",
                    0
                )

                payout_words = (
                    build_payout_words(
                        payout_value
                    )
                )

                division = get_val(
                    record,
                    "Division"
                )

                division_description = get_val(
                    record,
                    "Division Description",
                    "DivisionDescription"
                )

                full_division = (
                    build_full_division(
                        division,
                        division_description
                    )
                )

                contract_text = f"""
Company:
{get_val(record, "Company")}

Project:
{get_val(record, "Project Name", "ProjectName")}

{full_division}

Payout:
${float(payout_value):,.2f}

Amount in Words:
{payout_words}
"""

                pdf.multi_cell(
                    0,
                    8,
                    contract_text.strip()
                )

                pdf_bytes = pdf.output(
                    dest='S'
                ).encode('latin-1')

                company = clean_filename(
                    get_val(
                        record,
                        "Company"
                    )
                )

                filename = f"{company}.pdf"

                zip_file.writestr(
                    filename,
                    pdf_bytes
                )

        zip_buffer.seek(0)

        st.download_button(
            "📦 Download PDFs",
            zip_buffer,
            "contracts_pdf.zip"
        )

    # =====================================================
    # EXCEL EXPORT
    # =====================================================
    with tab3:

        output = io.BytesIO()

        pd.DataFrame(records).to_excel(
            output,
            index=False,
            engine='openpyxl'
        )

        output.seek(0)

        st.download_button(
            "📥 Download Excel",
            output,
            f"report_"
            f"{datetime.now().strftime('%Y%m%d')}"
            f".xlsx"
        )

    # =====================================================
    # WORD EXPORT
    # =====================================================
    with tab4:

        if not selected_templates:

            st.warning(
                "Select at least one template"
            )

            st.stop()

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            "w"
        ) as zip_file:

            for i, record in enumerate(records):

                payout_value = record.get(
                    "Payout",
                    0
                )

                payout_words = (
                    build_payout_words(
                        payout_value
                    )
                )

                division = get_val(
                    record,
                    "Division"
                )

                division_description = get_val(
                    record,
                    "Division Description",
                    "DivisionDescription"
                )

                full_division = (
                    build_full_division(
                        division,
                        division_description
                    )
                )

                context = {

                    # DATES
                    "Agreement_Date":
                        datetime.now().strftime(
                            '%Y-%m-%d'
                        ),

                    # COMPANY
                    "Company":
                        get_val(
                            record,
                            "Company"
                        ),

                    "CompanyAddress":
                        get_val(
                            record,
                            "Company Address",
                            "CompanyAddress"
                        ),

                    # SUBCONTRACTOR
                    "SubcontractorName":
                        get_val(
                            record,
                            "Subcontractor Name",
                            "SubcontractorName"
                        ),

                    "SubcontractorAddress":
                        get_val(
                            record,
                            "Subcontractor Address",
                            "SubcontractorAddress"
                        ),

                    "SubcontractorCity":
                        get_val(
                            record,
                            "Subcontractor City",
                            "SubcontractorCity"
                        ),

                    # POC
                    "POCName":
                        get_val(
                            record,
                            "POC Name",
                            "POCName"
                        ),

                    "POCPhone":
                        get_val(
                            record,
                            "POC Phone",
                            "POCPhone"
                        ),

                    "POCEmail":
                        get_val(
                            record,
                            "POC Email",
                            "POCEmail"
                        ),

                    # PROJECT
                    "ProjectName":
                        get_val(
                            record,
                            "Project Name",
                            "ProjectName"
                        ),

                    "ProjectNumber":
                        get_val(
                            record,
                            "Project Number",
                            "ProjectNumber"
                        ),

                    "ProjectAddress":
                        get_val(
                            record,
                            "Project Address",
                            "ProjectAddress"
                        ),

                    # DIVISION
                    "Division":
                        division,

                    "DivisionDescription":
                        division_description,

                    "FullDivision":
                        full_division,

                    # SCOPE
                    "Scope":
                        get_val(
                            record,
                            "Scope"
                        ),

                    "Addendums":
                        get_val(
                            record,
                            "Addendums"
                        ),

                    "alternate":
                        get_val(
                            record,
                            "alternate"
                        ),

                    # PAYOUT
                    "Payout":
                        f"{float(payout_value):,.2f}",

                    "PayoutWords":
                        payout_words,

                    # COMPLETION DATE
                    "CompletionDate":
                        str(
                            get_val(
                                record,
                                "Completion Date",
                                "CompletionDate"
                            )
                        )
                }

                # =================================================
                # CLEAN CONTEXT
                # =================================================
                context = {
                    k: clean_text(v)
                    for k, v in context.items()
                }

                # =================================================
                # GENERATE DOCUMENTS
                # =================================================
                for template_name in selected_templates:

                    template_path = os.path.join(
                        TEMPLATE_DIR,
                        template_name
                    )

                    doc = DocxTemplate(
                        template_path
                    )

                    doc.render(context)

                    file_stream = io.BytesIO()

                    doc.save(file_stream)

                    # =============================================
                    # FILE NAMING
                    # =============================================
                    project = get_val(
                        record,
                        "Project Name",
                        "ProjectName"
                    )

                    company = get_val(
                        record,
                        "Company"
                    )

                    project = clean_filename(
                        project
                    )

                    company = clean_filename(
                        company
                    )

                    filename = (
                        f"{project} "
                        f"Subcontract "
                        f"{company}.docx"
                    )

                    zip_file.writestr(
                        filename,
                        file_stream.getvalue()
                    )

        zip_buffer.seek(0)

        st.download_button(
            "📄 Download Word Docs",
            zip_buffer,
            "contracts_word.zip"
        )
