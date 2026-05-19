import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="P95 AI Revenue Module Generator", layout="wide")

st.title("P95 AI Revenue Module Generator")
st.caption("PMO Internal Tool • Clinical Study Revenue Automation")

# -----------------------------
# Helper functions
# -----------------------------

def month_key(dt):
    return datetime(dt.year, dt.month, 1)

def get_months_between(start, end):
    months = []
    current = datetime(start.year, start.month, 1)
    final = datetime(end.year, end.month, 1)

    while current <= final:
        months.append(current)
        if current.month == 12:
            current = datetime(current.year + 1, 1, 1)
        else:
            current = datetime(current.year, current.month + 1, 1)

    return months

def get_phase_dates(start, end):
    total_days = (end - start).days

    startup_end = start + timedelta(days=int(total_days * 0.25))
    execution_end = start + timedelta(days=int(total_days * 0.75))
    analysis_end = start + timedelta(days=int(total_days * 0.90))

    return {
        "startup": (start, startup_end),
        "execution": (startup_end, execution_end),
        "analysis": (execution_end, analysis_end),
        "closeout": (analysis_end, end),
    }

def assign_phase(activity, description):
    text = f"{activity} {description}".lower()

    if any(x in text for x in ["protocol", "sample size", "kom", "development"]):
        return "startup"

    if any(x in text for x in ["meeting", "monthly", "tc", "coordination", "internal"]):
        return "execution"

    if "review" in text:
        return "analysis"

    return "execution"

def spread_units(activity, description, total_units, start, end):
    phase_dates = get_phase_dates(start, end)
    phase = assign_phase(activity, description)

    phase_start, phase_end = phase_dates[phase]
    phase_months = get_months_between(phase_start, phase_end)

    spread = defaultdict(float)

    if total_units is None or pd.isna(total_units):
        total_units = 0

    total_units = float(total_units)
    description_text = str(description).lower()

    if len(phase_months) == 0:
        spread[month_key(start)] = total_units
        return spread

    if any(x in description_text for x in ["monthly", "bi-weekly", "biweekly", "meeting", "tc"]):
        per_month = total_units / len(phase_months)
        for m in phase_months:
            spread[month_key(m)] = per_month

    elif any(x in description_text for x in ["document", "process"]):
        spread[month_key(phase_start)] = total_units

    elif "review" in description_text:
        per_month = total_units / len(phase_months)
        for m in phase_months:
            spread[month_key(m)] = per_month

    else:
        spread[month_key(phase_start)] = total_units

    return spread

def generate_revenue_tracker(budget_file, template_file):
    budget_df = pd.read_excel(
        budget_file,
        sheet_name="Workload and Resources P1",
        header=None
    )

    wb = load_workbook(
        template_file,
        keep_links=True,
        data_only=False
    )

    if "UNIT TRACKER" not in wb.sheetnames:
        raise ValueError("UNIT TRACKER sheet not found in template.")

    ws = wb["UNIT TRACKER"]

    study_start = budget_df.iloc[23, 1]
    study_end = budget_df.iloc[23, 2]

    ws["G5"] = study_start
    ws["G6"] = study_end

    study_months = get_months_between(study_start, study_end)

    forecast_cols = {}
    start_col = 18  # R
    for i, m in enumerate(study_months):
        forecast_cols[month_key(m)] = start_col + (i * 6)

    budget_rows = budget_df.iloc[8:20]
    target_row = 40

    for _, row in budget_rows.iterrows():
        activity = row[0]
        description = row[2]
        units = row[3]
        unit_price = row[4]

        if pd.isna(activity):
            continue

        text = str(activity).lower()
        if "insert lines" in text or "sectiontotal" in text:
            continue

        ws[f"D{target_row}"] = study_start
        ws[f"E{target_row}"] = study_end
        ws[f"F{target_row}"] = activity
        ws[f"H{target_row}"] = description
        ws[f"I{target_row}"] = units
        ws[f"J{target_row}"] = unit_price

        unit_spread = spread_units(
            activity,
            description,
            units,
            study_start,
            study_end
        )

        total_forecast_units = 0
        for m, forecast_col in forecast_cols.items():
            forecast_units = unit_spread.get(m, 0)
            ws.cell(row=target_row, column=forecast_col).value = forecast_units
            total_forecast_units += forecast_units

        ws[f"M{target_row}"] = total_forecast_units

        target_row += 1

    # Validation
    validation_warnings = []
    rows_processed = 0
    total_contract_value = 0
    total_forecast_units = 0

    for row in range(40, target_row):
        activity = ws[f"F{row}"].value
        unit_description = ws[f"H{row}"].value
        contracted_units = ws[f"I{row}"].value
        unit_cost = ws[f"J{row}"].value
        forecast_units = ws[f"M{row}"].value

        rows_processed += 1

        if not activity:
            validation_warnings.append(f"Row {row}: Missing activity name")

        if not unit_description:
            validation_warnings.append(f"Row {row}: Missing unit description")

        if contracted_units in [None, "", 0]:
            validation_warnings.append(f"Row {row}: Missing or zero contracted units")

        if unit_cost in [None, "", 0]:
            validation_warnings.append(f"Row {row}: Missing or zero unit cost")

        if forecast_units and contracted_units and forecast_units > contracted_units:
            validation_warnings.append(f"Row {row}: Forecast units exceed contracted units")

        if contracted_units and unit_cost:
            total_contract_value += contracted_units * unit_cost

        if forecast_units:
            total_forecast_units += forecast_units

    high_risk = []
    medium_risk = []
    low_risk = []

    for warning in validation_warnings:
        text = warning.lower()

        if "forecast units exceed contracted units" in text:
            high_risk.append(warning)
        elif "missing or zero unit cost" in text:
            high_risk.append(warning)
        elif "missing activity name" in text:
            medium_risk.append(warning)
        elif "missing unit description" in text:
            medium_risk.append(warning)
        elif "missing or zero contracted units" in text:
            medium_risk.append(warning)
        else:
            low_risk.append(warning)

    if len(high_risk) == 0 and len(medium_risk) <= 2:
        confidence = "High"
    elif len(high_risk) <= 2:
        confidence = "Medium"
    else:
        confidence = "Low"

    actions = []

    if high_risk:
        actions.append("Review revenue-impacting rows immediately.")

    if medium_risk:
        actions.append("Validate missing data and incomplete assumptions.")

    if total_forecast_units > 0:
        actions.append("Confirm forecast spread aligns with study phases.")

    if total_contract_value > 0:
        actions.append("Verify contract totals against budget assumptions.")

    # AI PMO Review sheet
    if "AI PMO Review" in wb.sheetnames:
        del wb["AI PMO Review"]

    review_ws = wb.create_sheet("AI PMO Review")

    review_ws["A1"] = "P95 AI PMO REVIEW"
    review_ws["A3"] = "Revenue Confidence"
    review_ws["B3"] = confidence
    review_ws["A4"] = "Rows Processed"
    review_ws["B4"] = rows_processed
    review_ws["A5"] = "Warnings"
    review_ws["B5"] = len(validation_warnings)
    review_ws["A6"] = "High Risk"
    review_ws["B6"] = len(high_risk)
    review_ws["A7"] = "Medium Risk"
    review_ws["B7"] = len(medium_risk)
    review_ws["A8"] = "Low Risk"
    review_ws["B8"] = len(low_risk)
    review_ws["A9"] = "Total Contract Value"
    review_ws["B9"] = round(total_contract_value, 2)
    review_ws["A10"] = "Total Forecast Units"
    review_ws["B10"] = round(total_forecast_units, 2)

    row_num = 12
    review_ws[f"A{row_num}"] = "HIGH-RISK ITEMS"
    row_num += 1
    for item in high_risk:
        review_ws[f"A{row_num}"] = item
        row_num += 1

    row_num += 1
    review_ws[f"A{row_num}"] = "MEDIUM-RISK ITEMS"
    row_num += 1
    for item in medium_risk:
        review_ws[f"A{row_num}"] = item
        row_num += 1

    row_num += 1
    review_ws[f"A{row_num}"] = "LOW-RISK ITEMS"
    row_num += 1
    for item in low_risk:
        review_ws[f"A{row_num}"] = item
        row_num += 1

    row_num += 2
    review_ws[f"A{row_num}"] = "SUGGESTED PMO ACTIONS"
    row_num += 1
    for action in actions:
        review_ws[f"A{row_num}"] = action
        row_num += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    summary = {
        "confidence": confidence,
        "rows_processed": rows_processed,
        "warnings": len(validation_warnings),
        "high_risk": len(high_risk),
        "medium_risk": len(medium_risk),
        "low_risk": len(low_risk),
        "total_contract_value": round(total_contract_value, 2),
        "total_forecast_units": round(total_forecast_units, 2),
        "actions": actions
    }

    return output, summary

# -----------------------------
# UI
# -----------------------------

st.sidebar.header("Study Parameters")
st.sidebar.info("Study dates are currently read from the budget assumptions tab.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Budget File")
    budget_file = st.file_uploader("Upload Budget Excel (.xlsx)", type=["xlsx"])

with col2:
    st.subheader("Unit Tracker Template")
    template_file = st.file_uploader("Upload Unit Tracker Template (.xlsx)", type=["xlsx"])

if st.button("Generate Revenue Module"):
    if not budget_file or not template_file:
        st.error("Please upload both files.")
    else:
        try:
            output, summary = generate_revenue_tracker(budget_file, template_file)

            st.success("Revenue module generated successfully.")

            st.subheader("PMO Review Summary")
            st.write(f"Revenue Confidence: **{summary['confidence']}**")
            st.write(f"Rows Processed: **{summary['rows_processed']}**")
            st.write(f"Warnings: **{summary['warnings']}**")
            st.write(f"Total Contract Value: **{summary['total_contract_value']}**")
            st.write(f"Total Forecast Units: **{summary['total_forecast_units']}**")

            st.subheader("Suggested PMO Actions")
            for action in summary["actions"]:
                st.write(f"- {action}")

            st.download_button(
                label="Download Completed Revenue Tracker",
                data=output,
                file_name="P95_UNIT_TRACKER_PHASE_BASED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error("Something went wrong while generating the tracker.")
            st.exception(e)
            # Footer
st.divider()
st.caption("P95 AI Revenue Module Generator • Powered by Dania Alqudah")
