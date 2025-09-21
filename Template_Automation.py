import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook
from update_logic import (
    validate_template_logic, update_template, update_service_plan, update_service_offering,
    update_parts_pricing, update_labor_pricing, account_location_mapping,
    install_product_mapping, location_mapping_validate_fixed,
    validate_install_product_mapping, update_template_with_picklist, TEMPLATE_CONFIGS
)

st.set_page_config(page_title="Template Automation Tool", layout="wide")

# --- Sidebar Navigation ---
st.sidebar.markdown("""
<div style='background:#000;border-radius:12px;padding:18px 10px;margin-bottom:1em;text-align:center;'>
    <img src='https://img.icons8.com/fluency/96/excel.png' width='60'/><br>
    <span style='color:#bdbdbd;font-size:1.3rem;font-weight:600;'>SVC_CoE SMAX DLT Automation</span>
</div>
""", unsafe_allow_html=True)
page = st.sidebar.radio("Navigation", ["Update Template", "Mapping Template", "Picklist Values"], index=0)
st.sidebar.markdown("<hr style='border:1px solid #00008B;'>", unsafe_allow_html=True)
st.sidebar.info("Automate, validate, and map your Excel templates with a modern web interface.")

# --- Main Header ---
st.markdown(""" 
<style>
body, .main, .block-container {
    background-color: #0a174e !important; /* dark blue */
}
.main-title {
    font-size:2.2rem;font-weight:700;color:#bdbdbd;margin-bottom:0.5em; /* grey heading */
    background: linear-gradient(90deg,#0a174e 0%,#1a237e 100%);
    border-radius:16px;padding:18px 24px;margin-bottom:1em;
    box-shadow:0 2px 8px #222;
}
.section-title {
    font-size:1.3rem;font-weight:600;color:#bdbdbd;margin-top:1.5em;
    background:#1a237e;border-radius:12px;padding:10px 18px;margin-bottom:1em;
}
.stButton>button {
    background-color:#2176ae;color:white;font-weight:600;border-radius:8px;
    box-shadow:0 2px 8px #222;
}
.stDownloadButton>button {
    background-color:#21c197;color:white;font-weight:600;border-radius:8px;
    box-shadow:0 2px 8px #222;
}
.stTextInput>div>input, .stSelectbox>div>div, .stFileUploader>div, .stDataFrame {
    border-radius:8px;
    background:#0a174e;
    color:#bdbdbd;
}
.stSidebar {
    background: #000 !important; /* black sidebar */
    color: #bdbdbd !important;
}
.stRadio>div>label {
    background: #1a237e !important;
    color: #bdbdbd !important;
    border-radius: 8px;
    font-weight: 600;
    padding: 8px 16px;
    margin-bottom: 6px;
}
.stRadio>div>label[data-selected="true"] {
    background: #2176ae !important;
    color: #fff !important;
    border: 2px solid #2176ae;
}
</style>
<div class="main-title">Global SVC_CoE SMAX DLT Automation Tool</div>
""", unsafe_allow_html=True)

# --- Update Template Tab ---
if page == "Update Template":
    st.markdown('<div class="section-title">Update Template</div>', unsafe_allow_html=True)
    col1, col2 = st.columns([1,1])
    with col1:
        uploaded_file = st.file_uploader("Upload Excel Template", type=["xlsx"], key="template_file")
        sheet_name = st.text_input("Sheet Name in Template", value="INSERT", key="sheet_name")
        template_type = st.selectbox("Select Required DLT", [
            "Service Contract DLT", "Covered Product_Service Contract",
            "Warranty Service Contract", "Warranty Covered Product_Service Contract",
            "Service Plan", "Service Offering", "Parts Pricing", "Labor Pricing"
        ], key="template_type")
    with col2:
        source_file = st.file_uploader("Upload Source Data (Excel/CSV)", type=["xlsx", "csv"], key="source_file")
        st.markdown("*Note: Make sure the column names in both files are same.")
    st.markdown("---")
    if uploaded_file and source_file:
        template_path = f"temp_template_{uploaded_file.name}"
        source_path = f"temp_source_{source_file.name}"
        with open(template_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        with open(source_path, "wb") as f:
            f.write(source_file.getbuffer())
        if source_file.name.endswith('.csv'):
            source_df = pd.read_csv(source_path)
        else:
            source_df = pd.read_excel(source_path)
        st.subheader("Source Data Preview")
        st.dataframe(source_df.head(10), use_container_width=True)
        colA, colB, colC = st.columns([1,1,1])
        with colA:
            if st.button("Update Template", help="Update the template with source data and defaults", use_container_width=True):
                if template_type == "Service Plan":
                    update_result = update_service_plan(template_path, source_df, sheet_name=sheet_name)
                elif template_type == "Service Offering":
                    update_result = update_service_offering(template_path, source_df, sheet_name=sheet_name)
                elif template_type == "Parts Pricing":
                    update_result = update_parts_pricing(template_path, source_df, sheet_name=sheet_name)
                elif template_type == "Labor Pricing":
                    update_result = update_labor_pricing(template_path, source_df, sheet_name=sheet_name)
                else:
                    update_result = update_template(template_type, template_path, source_df, sheet_name=sheet_name)
                st.success(f"Update completed: {update_result}")
        with colB:
            if st.button("Validate Template", help="Run validation and see summary", use_container_width=True):
                validation_result = validate_template_logic(template_path, sheet_name=sheet_name, templates_to_validate=[template_type], source_df=source_df)
                st.markdown("#### Validation Summary")
                st.json({
                    "Total Records": validation_result.get("total_records"),
                    "Duplicate Temp ID Count": validation_result.get("duplicate_temp_id_count"),
                    "Default Mismatch Count": validation_result.get("default_mismatch_count"),
                    "Validation Passed": validation_result.get("validation_passed")
                })
                if validation_result.get("issues_df") is not None:
                    st.dataframe(validation_result["issues_df"], use_container_width=True)
        with colC:
            with open(template_path, "rb") as f:
                st.download_button("Download Updated Template", f, file_name=f"updated_{uploaded_file.name}", use_container_width=True)
        # Clean up temp files
        if os.path.exists(template_path):
            os.remove(template_path)
        if os.path.exists(source_path):
            os.remove(source_path)
    else:
        st.info("Please upload both template and source files to proceed.")

# --- Mapping Template Tab ---
elif page == "Mapping Template":
    st.markdown('<div class="section-title">Mapping Template</div>', unsafe_allow_html=True)
    col1, col2 = st.columns([1,1])
    with col1:
        mapping_source_file = st.file_uploader("Upload Source Data File (Excel/CSV)", type=["xlsx", "csv"], key="mapping_source_file")
    with col2:
        mapping_file = st.file_uploader("Upload Mapping File (Excel)", type=["xlsx"], key="mapping_file")
    mapping_type = st.selectbox("Mapping Type", ["Account and Location Mapping", "IP Mapping"], key="mapping_type")
    st.markdown("---")
    if mapping_source_file and mapping_file:
        mapping_source_path = f"temp_mapping_source_{mapping_source_file.name}"
        mapping_file_path = f"temp_mapping_file_{mapping_file.name}"
        with open(mapping_source_path, "wb") as f:
            f.write(mapping_source_file.getbuffer())
        with open(mapping_file_path, "wb") as f:
            f.write(mapping_file.getbuffer())
        if mapping_source_file.name.endswith('.csv'):
            mapping_source_df = pd.read_csv(mapping_source_path)
        else:
            mapping_source_df = pd.read_excel(mapping_source_path)
        st.subheader("Source Data Preview")
        st.dataframe(mapping_source_df.head(10), use_container_width=True)
        colA, colB, colC = st.columns([1,1,1])
        with colA:
            if st.button("Update Mapping", help="Update mapping based on selected type", use_container_width=True):
                if mapping_type == "Account and Location Mapping":
                    success = account_location_mapping(mapping_source_path, mapping_file_path)
                elif mapping_type == "IP Mapping":
                    success = install_product_mapping(mapping_source_path, mapping_file_path)
                st.success("Mapping updated successfully." if success else "Mapping update failed.")
        with colB:
            if st.button("Validate Mapping", help="Run mapping validation and see summary", use_container_width=True):
                if mapping_type == "Account and Location Mapping":
                    summary = location_mapping_validate_fixed(mapping_source_path, mapping_file_path)
                elif mapping_type == "IP Mapping":
                    summary = validate_install_product_mapping(mapping_source_path, mapping_file_path)
                st.markdown("#### Validation Summary")
                st.json(summary)
        with colC:
            pass
        # Clean up temp files
        if os.path.exists(mapping_source_path):
            os.remove(mapping_source_path)
        if os.path.exists(mapping_file_path):
            os.remove(mapping_file_path)
    else:
        st.info("Please upload both source and mapping files to proceed.")

# --- Picklist Values Tab ---
elif page == "Picklist Values":
    st.markdown('<div class="section-title">Picklist Values</div>', unsafe_allow_html=True)
    picklist_template = st.selectbox("Select Template", [""] + list(TEMPLATE_CONFIGS.keys()), key="picklist_template")
    picklist_values = {}
    if picklist_template:
        picklist_config = TEMPLATE_CONFIGS.get(picklist_template, {})
        picklist_columns = picklist_config.keys()
        col1, col2 = st.columns([1,1])
        for i, col in enumerate(picklist_columns):
            if isinstance(picklist_config[col], dict):
                continue
            values = picklist_config[col] if isinstance(picklist_config[col], list) else []
            if i % 2 == 0:
                picklist_values[col] = col1.selectbox(f"{col}", values, key=f"picklist_{col}")
            else:
                picklist_values[col] = col2.selectbox(f"{col}", values, key=f"picklist_{col}")
    uploaded_template_file = st.file_uploader("Upload Template File for Picklist", type=["xlsx"], key="picklist_template_file")
    picklist_sheet_name = st.text_input("Sheet Name", value="INSERT", key="picklist_sheet_name")
    st.markdown("---")
    colA, colB = st.columns([1,1])
    with colA:
        if st.button("Apply Picklist Values", help="Apply selected picklist values to template", use_container_width=True):
            if uploaded_template_file and picklist_template:
                picklist_template_path = f"temp_picklist_template_{uploaded_template_file.name}"
                with open(picklist_template_path, "wb") as f:
                    f.write(uploaded_template_file.getbuffer())
                success = update_template_with_picklist(picklist_template_path, picklist_sheet_name, picklist_values)
                st.success("Selected picklist values applied to template." if success else "Failed to update the template file. It may be open or invalid.")
                with open(picklist_template_path, "rb") as f:
                    st.download_button("Download Updated Template", f, file_name=f"updated_{uploaded_template_file.name}", use_container_width=True)
                if os.path.exists(picklist_template_path):
                    os.remove(picklist_template_path)
            else:
                st.warning("Please upload a template file and select a template.")
