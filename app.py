
from pathlib import Path


import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook


# =========================================================
# 1. CONFIGURATION
# =========================================================

EXCEL_PATH = (
    Path(__file__).parent
    / "M1_Manpower_Model_English_Final_v18_Gap_Analysis.xlsx"
)

st.set_page_config(
    page_title="M1 Manpower Planning DSS",
    page_icon="📊",
    layout="wide"
)


# =========================================================
# LOAD CUSTOM CSS
# =========================================================

CSS_PATH = Path(__file__).parent / "style.css"

if CSS_PATH.exists():
    with open(CSS_PATH, "r", encoding="utf-8") as f:
        st.markdown(
            f"<style>{f.read()}</style>",
            unsafe_allow_html=True
        )


# =========================================================
# 2. EXCEL READING FUNCTION
# =========================================================

@st.cache_data
def load_excel_data(file_path, modified_time):

    wb = load_workbook(
        file_path,
        data_only=True,
        read_only=True
    )

    required_sheets = [
        "Area Summary",
        "AHP_Calculation",
        "SAW Calculation",
        "Criteria Normalized"
    ]

    missing = [
        sheet for sheet in required_sheets
        if sheet not in wb.sheetnames
    ]

    if missing:
        wb.close()
        raise ValueError(
            "Missing worksheets: " + ", ".join(missing)
        )

    # -----------------------------------------------------
    # A. AREA SUMMARY
    # -----------------------------------------------------

    ws = wb["Area Summary"]

    area_data = []

    for row in ws.iter_rows(
        min_row=2,
        max_row=9,
        min_col=1,
        max_col=12,
        values_only=True
    ):

        if row[0] is None:
            continue

        area_data.append({
            "Area": row[0],
            "Machine Qty": row[1],
            "OP Workload Hours / Shift": row[2],
            "Required OP": row[3],
            "Current OP": row[4],
            "OP Gap": row[5],
            "EE Workload Hours / Month": row[6],
            "Required EE": row[7],
            "Estimated Current EE": row[8],
            "EE Gap": row[9],
            "Required PE Specialist": row[10],
            "Required EE Engineer": row[11]
        })

    area_df = pd.DataFrame(area_data)

    if area_df.empty:
        wb.close()
        raise ValueError("Area Summary contains no data.")

    # -----------------------------------------------------
    # B. AHP PAIRWISE COMPARISON MATRIX
    # -----------------------------------------------------

    ws_ahp = wb["AHP_Calculation"]

    criteria_names = [
        "C1 OP Shortage",
        "C2 EE Shortage",
        "C3 Machine Workload",
        "C4 Operational Criticality"
    ]

    criteria_codes = ["C1", "C2", "C3", "C4"]

    pairwise_data = []

    for i in range(4):

        values = [
            ws_ahp.cell(
                row=4 + i,
                column=3 + j
            ).value
            for j in range(4)
        ]

        pairwise_data.append(values)

    pairwise_df = pd.DataFrame(
        pairwise_data,
        index=criteria_names,
        columns=criteria_codes
    )

    # -----------------------------------------------------
    # C. AHP NORMALIZED MATRIX AND WEIGHTS
    # -----------------------------------------------------

    normalized_data = []
    weights_data = []

    for i in range(4):

        values = [
            ws_ahp.cell(
                row=12 + i,
                column=3 + j
            ).value
            for j in range(4)
        ]

        normalized_data.append(values)

        weights_data.append(
            ws_ahp.cell(
                row=12 + i,
                column=7
            ).value
        )

    normalized_ahp_df = pd.DataFrame(
        normalized_data,
        index=criteria_names,
        columns=criteria_codes
    )

    weights_df = pd.DataFrame({
        "Criterion": criteria_codes,
        "Description": criteria_names,
        "Weight": weights_data
    })

    # -----------------------------------------------------
    # D. AHP CONSISTENCY RESULTS
    # -----------------------------------------------------

    consistency_data = {
        "Lambda Max": ws_ahp["I17"].value,
        "CI": ws_ahp["I18"].value,
        "RI": ws_ahp["I19"].value,
        "CR": ws_ahp["I20"].value,
        "Status": ws_ahp["I21"].value
    }

    # -----------------------------------------------------
    # E. CRITERIA NORMALIZED DATA
    # -----------------------------------------------------

    ws_criteria = wb["Criteria Normalized"]

    criteria_data = []

    for i in range(8):

        row = 36 + i

        criteria_data.append({
            "Area": ws_criteria.cell(row=row, column=2).value,
            "C1": ws_criteria.cell(row=row, column=3).value,
            "C2": ws_criteria.cell(row=row, column=4).value,
            "C3": ws_criteria.cell(row=row, column=5).value,
            "C4": ws_criteria.cell(row=row, column=6).value
        })

    criteria_df = pd.DataFrame(criteria_data)

    # -----------------------------------------------------
    # F. SAW CALCULATION
    # -----------------------------------------------------

    ws_saw = wb["SAW Calculation"]

    saw_data = []

    for row in range(2, 10):

        saw_data.append({
            "Area": ws_saw.cell(row=row, column=1).value,
            "C1": ws_saw.cell(row=row, column=2).value,
            "C2": ws_saw.cell(row=row, column=3).value,
            "C3": ws_saw.cell(row=row, column=4).value,
            "C4": ws_saw.cell(row=row, column=5).value,
            "SAW Score": ws_saw.cell(row=row, column=6).value,
            "Rank": ws_saw.cell(row=row, column=7).value
        })

    saw_df = pd.DataFrame(saw_data)

    saw_weights = {
        "C1": ws_saw["B12"].value,
        "C2": ws_saw["C12"].value,
        "C3": ws_saw["D12"].value,
        "C4": ws_saw["E12"].value
    }

    saw_cr = ws_saw["B14"].value
    saw_status = ws_saw["B15"].value

    wb.close()

    return (
        area_df,
        pairwise_df,
        normalized_ahp_df,
        weights_df,
        consistency_data,
        criteria_df,
        saw_df,
        saw_weights,
        saw_cr,
        saw_status
    )


# =========================================================
# 3. LOAD WORKBOOK
# =========================================================

st.title(
    "M1 Plant-Wide Manpower Planning and Decision Support System"
)

st.caption(
    "AHP-Based Criteria Weighting | "
    "SAW-Based Production Area Prioritisation"
)

if not EXCEL_PATH.exists():

    st.error("Excel workbook not found.")

    st.code(str(EXCEL_PATH))

    st.info(
        "Please check that the Excel file is inside "
        "the data folder and the filename is correct."
    )

    st.stop()


try:

    file_modified = EXCEL_PATH.stat().st_mtime

    (
        area_df,
        pairwise_df,
        normalized_ahp_df,
        weights_df,
        consistency_data,
        criteria_df,
        saw_df,
        saw_weights,
        saw_cr,
        saw_status
    ) = load_excel_data(
        str(EXCEL_PATH),
        file_modified
    )

except Exception as error:

    st.error("Unable to read Excel data.")

    st.exception(error)

    st.stop()


# =========================================================
# 4. SIDEBAR NAVIGATION
# =========================================================

st.sidebar.title("M1 DSS")

page = st.sidebar.radio(
    "Navigation",
    [
        "Manpower Overview",
        "AHP Analysis",
        "SAW Ranking"
    ]
)

st.sidebar.divider()

st.sidebar.caption("Data Source")

st.sidebar.write(EXCEL_PATH.name)

st.sidebar.caption(
    "All calculations are sourced directly from Excel."
)


# =========================================================
# PAGE 1: MANPOWER OVERVIEW
# =========================================================

if page == "Manpower Overview":

    st.header("Manpower Overview")

    # KPI values

    total_machines = area_df["Machine Qty"].sum()
    total_required_op = area_df["Required OP"].sum()
    total_current_op = area_df["Current OP"].sum()
    total_op_gap = area_df["OP Gap"].sum()

    total_required_ee = area_df["Required EE"].sum()
    total_current_ee = area_df["Estimated Current EE"].sum()
    total_ee_gap = area_df["EE Gap"].sum()

    # KPI cards

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Total Machines", f"{total_machines:,.0f}")
    c2.metric("Required OP", f"{total_required_op:,.0f}")
    c3.metric("Current OP", f"{total_current_op:,.0f}")
    c4.metric("OP Net Gap", f"{total_op_gap:,.0f}")

    c5, c6, c7 = st.columns(3)

    c5.metric("Required EE", f"{total_required_ee:,.0f}")
    c6.metric("Current EE", f"{total_current_ee:,.0f}")
    c7.metric("EE Gap", f"{total_ee_gap:,.0f}")

    st.divider()

    # Manpower table

    st.subheader("Area-Level Manpower Summary")

    display_columns = [
        "Area",
        "Machine Qty",
        "Required OP",
        "Current OP",
        "OP Gap",
        "Required EE",
        "Estimated Current EE",
        "EE Gap"
    ]

    display_df = area_df[display_columns].copy()

    st.table(display_df)

    st.divider()

    
    # OP Manpower Chart

    st.subheader("OP Manpower")

    op_chart_data = area_df[
     ["Area", "Required OP", "Current OP"]
    ].copy()

    op_chart_data = op_chart_data.set_index("Area")

    op_chart_data["Required OP"] = pd.to_numeric(
     op_chart_data["Required OP"],
     errors="coerce"
    )

    op_chart_data["Current OP"] = pd.to_numeric(
     op_chart_data["Current OP"],
     errors="coerce"
    )

    st.bar_chart(
     op_chart_data,
     height=400
    )


    
    # EE Manpower Chart

    st.subheader("EE Manpower")

    ee_chart_data = area_df[
     ["Area", "Required EE", "Estimated Current EE"]
    ].copy()

    ee_chart_data = ee_chart_data.set_index("Area")

    ee_chart_data = ee_chart_data.apply(
     pd.to_numeric,
     errors="coerce"
    )

    st.bar_chart(
     ee_chart_data,
     height=400
    )

# =========================================================
# PAGE 2: AHP ANALYSIS
# =========================================================

elif page == "AHP Analysis":

    st.header("Analytic Hierarchy Process (AHP)")

    st.write(
        "AHP is used to determine the relative importance "
        "of the four manpower prioritisation criteria."
    )

# Consistency results

    st.subheader("Consistency Test")

    a, b, c, d = st.columns(4)

    def display_value(value, digits=4):

        try:
            return f"{float(value):.{digits}f}"
        except (TypeError, ValueError):
            return "N/A"

    a.metric(
        "Lambda Max",
        display_value(consistency_data["Lambda Max"])
    )

    b.metric(
        "CI",
        display_value(consistency_data["CI"], 6)
    )

    c.metric(
        "RI",
        display_value(consistency_data["RI"])
    )

    d.metric(
        "CR",
        display_value(consistency_data["CR"], 6)
    )

    st.write(
        "Consistency Status:",
        consistency_data["Status"]
    )

    st.caption(
        "All AHP values are read directly from "
        "the AHP_Calculation worksheet."
    )


    # Pairwise matrix

    st.subheader("Pairwise Comparison Matrix")

    st.table(
        pairwise_df.round(4)
    )

    # Normalized matrix

    st.subheader("Normalized Pairwise Comparison Matrix")

    st.table(
        normalized_ahp_df.round(4)
    )

    # Criteria weights

    st.subheader("Criteria Weights")

    weight_display = weights_df.copy()

    weight_display["Weight (%)"] = (
        pd.to_numeric(
            weight_display["Weight"],
            errors="coerce"
        ) * 100
    ).round(2)

    st.table(
        weight_display[
            ["Criterion", "Description", "Weight (%)"]
        ]
    )

    
    # AHP Criteria Weight Graph

    st.subheader("AHP Criteria Weight Distribution")

    ahp_chart = weight_display[
        ["Criterion", "Weight (%)"]
    ].copy()

    ahp_chart["Weight (%)"] = pd.to_numeric(
        ahp_chart["Weight (%)"],
        errors="coerce"
    )

    st.bar_chart(
        ahp_chart.set_index("Criterion"),
        y="Weight (%)",
        height=350
    )


# =========================================================
# PAGE 3: SAW RANKING
# =========================================================

elif page == "SAW Ranking":

    st.header("Simple Additive Weighting (SAW) Ranking")

    st.write(
        "The SAW results below are directly retrieved "
        "from the Excel workbook."
    )

 # Consistency Reference

    st.subheader("AHP Consistency Reference")

    col1, col2 = st.columns(2)

    cr_value = (
        f"{float(saw_cr):.6f}"
        if saw_cr is not None
        else "N/A"
    )

    col1.metric(
        "Consistency Ratio (CR)",
        cr_value
    )

    col2.metric(
        "Consistency Status",
        str(saw_status) if saw_status is not None else "N/A"
    )

    st.caption(
        "SAW scores, criteria values, ranks and weights "
        "are read directly from the SAW Calculation worksheet. "
        "Python does not recalculate AHP or SAW."
    )

    

    # Criteria weights

    st.subheader("Criteria Weights")

    weight_table = pd.DataFrame({
        "Criterion": list(saw_weights.keys()),
        "Weight (%)": [
            (
                float(value) * 100
                if value is not None else None
            )
            for value in saw_weights.values()
        ]
    })

    weight_table["Weight (%)"] = (
        weight_table["Weight (%)"].round(2)
    )

    st.table(weight_table)

    # SAW results

    st.subheader("Normalized Criteria and SAW Results")

    ranking_display = saw_df.sort_values(
        "Rank"
    ).copy()

    numeric_columns = [
        "C1",
        "C2",
        "C3",
        "C4",
        "SAW Score"
    ]

    ranking_display[numeric_columns] = (
        ranking_display[numeric_columns].apply(
            pd.to_numeric,
            errors="coerce"
        ).round(4)
    )

    st.table(ranking_display)

    st.divider()

    
    # SAW Ranking Chart

    st.subheader("Production Area Priority Ranking")

    saw_chart_data = saw_df[
        ["Area", "SAW Score", "Rank"]
    ].copy()

    saw_chart_data["SAW Score"] = pd.to_numeric(
        saw_chart_data["SAW Score"],
        errors="coerce"
    )

    saw_chart_data["Rank"] = pd.to_numeric(
        saw_chart_data["Rank"],
        errors="coerce"
    )

    saw_chart_data = saw_chart_data.sort_values(
        "Rank",
        ascending=False
    )

    saw_chart_data = saw_chart_data.set_index("Area")

    st.bar_chart(
        saw_chart_data[["SAW Score"]],
        height=450
    )


# =========================================================
# END OF APPLICATION
# =========================================================
