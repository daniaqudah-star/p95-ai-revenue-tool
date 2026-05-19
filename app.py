import streamlit as st
from datetime import date

st.set_page_config(
    page_title="P95 AI Revenue Module Generator",
    layout="wide"
)

st.title("P95 AI Revenue Module Generator")
st.caption("PMO Internal Tool • Clinical Study Revenue Automation")

# Sidebar
st.sidebar.header("Study Parameters")

study_start = st.sidebar.date_input(
    "Study Start Date",
    value=date.today()
)

study_duration = st.sidebar.number_input(
    "Study Duration (months)",
    min_value=1,
    max_value=60,
    value=12
)

# Upload section
col1, col2 = st.columns(2)

with col1:
    st.subheader("Budget File")
    budget_file = st.file_uploader(
        "Upload Budget Excel (.xlsx)",
        type=["xlsx"]
    )

with col2:
    st.subheader("Unit Tracker Template")
    template_file = st.file_uploader(
        "Upload Unit Tracker (.xlsx)",
        type=["xlsx"]
    )

# Generate button
if st.button("Generate Revenue Module"):

    if not budget_file or not template_file:
        st.error("Please upload both files.")
    else:
        st.success("Files uploaded successfully.")
        st.info("Revenue engine ready.")

# Sections
st.divider()

st.subheader("Revenue Engine")
st.write("Budget → Timeline → Phase Logic → Revenue Spread")

st.subheader("AI Review")
st.write("Validation summary, confidence, PMO actions")

st.subheader("Assumptions")
st.write("Study phase logic and contract assumptions")

st.subheader("Validation Warnings")
st.write("Missing data, revenue mismatch, forecast risks")
