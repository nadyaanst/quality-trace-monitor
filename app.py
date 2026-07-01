import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import io

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Quality Trace Monitor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# CSS
# =========================================================
st.markdown("""
<style>
.stApp{
    background:#eef2f7;
    font-family:'Segoe UI', sans-serif;
}
section[data-testid="stSidebar"]{
    background:linear-gradient(180deg,#071952 0%,#0B2A6F 100%);
}
section[data-testid="stSidebar"] *{
    color:white !important;
}
.qtm-header{
    background:white;
    padding:30px;
    border-radius:20px;
    border-left:8px solid #B22222;
    box-shadow:0 6px 20px rgba(0,0,0,0.08);
    margin-bottom:20px;
}
.qtm-title{
    font-size:50px;
    font-weight:900;
    color:#081F5C;
}
.qtm-subtitle{
    font-size:18px;
    color:#666;
}
.kpi-box{
    background:white;
    padding:20px;
    border-radius:16px;
    text-align:center;
    box-shadow:0 4px 15px rgba(0,0,0,0.08);
}
.kpi-title{
    font-size:14px;
    font-weight:700;
    color:#777;
}
.kpi-value{
    font-size:28px;
    font-weight:900;
    color:#8B0000;
}
.alert-box{
    background:#fff5f5;
    padding:15px;
    border-left:6px solid red;
    border-radius:10px;
    margin-bottom:10px;
}
.section-title{
    font-size:28px;
    font-weight:800;
    color:#081F5C;
    margin-top:20px;
    margin-bottom:15px;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# CONSTANT
# =========================================================
COLUMNS = [
    "No","Tanggal","Shift","No HP","Layer HP",
    "Kode Mold","No Lot","Keterangan",
    "Temp","Pressure","Cycle Time","Operator"
]

DEFECT_LIST = [
    "OK","Visual","Dimensi","Visual Dimensi",
    "Scratch","Bubble","Flash","Short Mold"
]

ROOT_CAUSE_MAP = {
    "Visual": ("Material contamination", 82),
    "Dimensi": ("Mold wear", 91),
    "Visual Dimensi": ("Mixed quality issue", 84),
    "Scratch": ("Handling issue", 77),
    "Bubble": ("Cooling issue", 85),
    "Flash": ("Clamping pressure low", 88),
    "Short Mold": ("Material fill issue", 86)
}

# =========================================================
# SESSION STATE
# =========================================================
if "db" not in st.session_state:
    st.session_state.db = pd.DataFrame(columns=COLUMNS)

if "audit_log" not in st.session_state:
    st.session_state.audit_log = []

# =========================================================
# HELPER
# =========================================================
def section(title):
    st.markdown(
        f'<div class="section-title">{title}</div>',
        unsafe_allow_html=True
    )

def kpi_card(title, value):
    st.markdown(f"""
    <div class="kpi-box">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)

def add_log(action):
    st.session_state.audit_log.append({
        "Action": action,
        "Time": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    })

def convert_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.markdown("# QUALITY TRACE MONITOR")
st.sidebar.markdown("AI Manufacturing Intelligence")
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Navigation",
    ["Dashboard", "Manual Input", "Data Table", "Audit Trail"]
)

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel",
    type=["xlsx"]
)

if uploaded_file:
    df_upload = pd.read_excel(uploaded_file)

    for col in COLUMNS:
        if col not in df_upload.columns:
            df_upload[col] = None

    df_upload = df_upload[COLUMNS]

    st.session_state.db = df_upload
    add_log("Upload Excel")

    st.sidebar.success("Upload berhasil")

# =========================================================
# HEADER
# =========================================================
st.markdown(f"""
<div class="qtm-header">
    <div class="qtm-title">QUALITY TRACE MONITOR</div>
    <div class="qtm-subtitle">
        AI-Powered Manufacturing Quality Intelligence
    </div>
    <br>
    <b>System Status:</b> 🟢 ONLINE &nbsp;&nbsp;
    <b>Last Update:</b> {datetime.now().strftime("%d/%m/%Y %H:%M")}
</div>
""", unsafe_allow_html=True)

df = st.session_state.db.copy()

if df.empty:
    st.warning("Upload Excel terlebih dahulu.")
    st.stop()

df["Tanggal"] = pd.to_datetime(
    df["Tanggal"],
    errors="coerce",
    dayfirst=True
)

df["Keterangan"] = df["Keterangan"].astype(str)

df["is_ok"] = (
    df["Keterangan"].str.upper().eq("OK").astype(int)
)

df["is_ng"] = (
    df["Keterangan"].str.upper().ne("OK").astype(int)
)

# =========================================================
# DASHBOARD PAGE
# =========================================================
if menu == "Dashboard":

    # =====================================================
    # FILTER
    # =====================================================
    section("FILTER DATA")

    f1, f2, f3, f4 = st.columns(4)

    with f1:
        selected_shift = st.multiselect(
            "Shift",
            sorted(df["Shift"].dropna().unique())
        )

    with f2:
        selected_hp = st.multiselect(
            "No HP",
            sorted(df["No HP"].dropna().unique())
        )

    with f3:
        selected_defect = st.multiselect(
            "Keterangan",
            sorted(df["Keterangan"].dropna().unique())
        )

    with f4:
        date_range = st.date_input(
            "Date Range",
            []
        )

    df_filtered = df.copy()

    if selected_shift:
        df_filtered = df_filtered[
            df_filtered["Shift"].isin(selected_shift)
        ]

    if selected_hp:
        df_filtered = df_filtered[
            df_filtered["No HP"].isin(selected_hp)
        ]

    if selected_defect:
        df_filtered = df_filtered[
            df_filtered["Keterangan"].isin(selected_defect)
        ]

    if len(date_range) == 2:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1])

        df_filtered = df_filtered[
            (df_filtered["Tanggal"] >= start_date) &
            (df_filtered["Tanggal"] <= end_date)
        ]

    if df_filtered.empty:
        st.warning("Tidak ada data sesuai filter.")
        st.stop()

    # =====================================================
    # KPI CALCULATION
    # =====================================================
    total_production = len(df_filtered)
    total_ok = int(df_filtered["is_ok"].sum())
    total_defect = int(df_filtered["is_ng"].sum())

    yield_rate = (
        total_ok / total_production * 100
        if total_production > 0 else 0
    )

    machine_ng = (
        df_filtered.groupby("No HP")["is_ng"]
        .sum()
        .sort_values(ascending=False)
    )

    critical_machine = (
        machine_ng.index[0]
        if len(machine_ng) > 0 else "-"
    )

    mold_ng = (
        df_filtered[df_filtered["is_ng"] == 1]
        .groupby("Kode Mold")
        .size()
        .sort_values(ascending=False)
    )

    critical_mold = (
        mold_ng.index[0]
        if len(mold_ng) > 0 else "-"
    )

    # =====================================================
    # KPI DISPLAY
    # =====================================================
    section("KEY PERFORMANCE INDICATORS")

    k1, k2, k3, k4, k5, k6 = st.columns(6)

    with k1:
        kpi_card("PRODUCTION", total_production)

    with k2:
        kpi_card("TOTAL OK", total_ok)

    with k3:
        kpi_card("DEFECT", total_defect)

    with k4:
        kpi_card("YIELD", f"{yield_rate:.2f}%")

    with k5:
        kpi_card("CRITICAL MACHINE", critical_machine)

    with k6:
        kpi_card("CRITICAL MOLD", critical_mold)

    # =====================================================
    # SMART ALERT
    # =====================================================
    section("SMART ALERT CENTER")

    alerts = []

    machine_defect_rate = (
        df_filtered.groupby("No HP")
        .apply(lambda x: x["is_ng"].sum() / len(x))
    )

    for machine, rate in machine_defect_rate.items():
        if rate >= 0.30:
            alerts.append(
                f"🔴 CRITICAL: Machine {machine} defect rate {rate:.1%}"
            )
        elif rate >= 0.20:
            alerts.append(
                f"🟡 WARNING: Machine {machine} defect rate {rate:.1%}"
            )

    if yield_rate < 80:
        alerts.append(
            f"🔴 Yield rate turun ke {yield_rate:.2f}%"
        )

    if len(mold_ng) > 0:
        top_mold_count = mold_ng.iloc[0]
        if top_mold_count >= 5:
            alerts.append(
                f"⚠ Mold {critical_mold} defect tinggi ({top_mold_count})"
            )

    if not alerts:
        st.success("🟢 Semua parameter normal.")
    else:
        for alert in alerts:
            st.markdown(
                f'<div class="alert-box">{alert}</div>',
                unsafe_allow_html=True
            )

    # =====================================================
    # MONITORING HARIAN
    # =====================================================
    section("MONITORING HARIAN")

    daily = (
        df_filtered.groupby("Tanggal")
        .agg({
            "Layer HP": "count",
            "is_ok": "sum"
        })
        .reset_index()
    )

    daily.columns = [
        "Tanggal",
        "Production",
        "OK"
    ]

    daily["Yield"] = np.where(
        daily["Production"] > 0,
        daily["OK"] / daily["Production"],
        0
    )

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=daily["Tanggal"],
            y=daily["Production"],
            name="Production"
        ),
        secondary_y=False
    )

    fig.add_trace(
        go.Bar(
            x=daily["Tanggal"],
            y=daily["OK"],
            name="OK"
        ),
        secondary_y=False
    )

    fig.add_trace(
        go.Scatter(
            x=daily["Tanggal"],
            y=daily["Yield"],
            mode="lines+markers",
            name="Yield"
        ),
        secondary_y=True
    )

    fig.update_yaxes(
        title_text="Production",
        secondary_y=False
    )

    fig.update_yaxes(
        title_text="Yield",
        tickformat=".0%",
        secondary_y=True
    )

    fig.update_layout(
        height=550,
        barmode="group"
    )

    st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # PARETO DEFECT
    # =====================================================
    section("PARETO DEFECT ANALYSIS")

    pareto = (
        df_filtered[df_filtered["is_ng"] == 1]
        .groupby("Keterangan")
        .size()
        .reset_index(name="Jumlah")
    )

    pareto = pareto.sort_values(
        by="Jumlah",
        ascending=False
    )

    if not pareto.empty:

        pareto["Cum %"] = (
            pareto["Jumlah"].cumsum()
            / pareto["Jumlah"].sum() * 100
        )

        fig_pareto = make_subplots(
            specs=[[{"secondary_y": True}]]
        )

        fig_pareto.add_trace(
            go.Bar(
                x=pareto["Keterangan"],
                y=pareto["Jumlah"],
                name="Defect Count"
            ),
            secondary_y=False
        )

        fig_pareto.add_trace(
            go.Scatter(
                x=pareto["Keterangan"],
                y=pareto["Cum %"],
                mode="lines+markers",
                name="Cumulative %"
            ),
            secondary_y=True
        )

        fig_pareto.update_yaxes(
            title_text="Jumlah",
            secondary_y=False
        )

        fig_pareto.update_yaxes(
            title_text="Cumulative %",
            secondary_y=True
        )

        fig_pareto.update_layout(height=500)

        st.plotly_chart(
            fig_pareto,
            use_container_width=True
        )
    else:
        st.info("Belum ada defect.")

    # =====================================================
    # ROOT CAUSE ANALYSIS
    # =====================================================
    section("ROOT CAUSE ANALYSIS")

    ng_df = df_filtered[
        df_filtered["is_ng"] == 1
    ]

    if not ng_df.empty:

        dominant_defect = (
            ng_df["Keterangan"]
            .value_counts()
            .index[0]
        )

        cause, confidence = ROOT_CAUSE_MAP.get(
            dominant_defect,
            ("Unknown", 50)
        )

        rc1, rc2, rc3 = st.columns(3)

        with rc1:
            kpi_card("DOMINANT DEFECT", dominant_defect)

        with rc2:
            kpi_card("ROOT CAUSE", cause)

        with rc3:
            kpi_card("CONFIDENCE", f"{confidence}%")

    else:
        st.success("Tidak ada defect.")

    # =====================================================
    # PREDICTION ENGINE
    # =====================================================
    section("PREDICTION ENGINE")

    machine_rate = (
        total_defect / total_production
        if total_production > 0 else 0
    )

    mold_rate = (
        mold_ng.iloc[0] / total_production
        if len(mold_ng) > 0 and total_production > 0
        else 0
    )

    temp_abnormal = 0
    pressure_abnormal = 0

    if "Temp" in df_filtered.columns:
        temp_abnormal = (
            (
                (pd.to_numeric(df_filtered["Temp"], errors="coerce") > 230)
                |
                (pd.to_numeric(df_filtered["Temp"], errors="coerce") < 180)
            ).sum()
        ) / total_production

    if "Pressure" in df_filtered.columns:
        pressure_abnormal = (
            (
                (pd.to_numeric(df_filtered["Pressure"], errors="coerce") > 120)
                |
                (pd.to_numeric(df_filtered["Pressure"], errors="coerce") < 80)
            ).sum()
        ) / total_production

    risk_score = (
        0.35 * machine_rate +
        0.35 * mold_rate +
        0.15 * temp_abnormal +
        0.15 * pressure_abnormal
    ) * 100

    if risk_score <= 30:
        risk_level = "SAFE"
    elif risk_score <= 60:
        risk_level = "WARNING"
    else:
        risk_level = "CRITICAL"

    p1, p2 = st.columns(2)

    with p1:
        kpi_card("RISK SCORE", f"{risk_score:.2f}")

    with p2:
        kpi_card("RISK LEVEL", risk_level)

    # =====================================================
    # AI QUALITY SCORE
    # =====================================================
    section("AI QUALITY SCORE")

    quality_score = (
        yield_rate * 0.7 +
        max(0, 100 - risk_score) * 0.3
    )

    if quality_score >= 90:
        grade = "A"
    elif quality_score >= 80:
        grade = "B"
    elif quality_score >= 70:
        grade = "C"
    else:
        grade = "D"

    q1, q2 = st.columns(2)

    with q1:
        kpi_card("QUALITY SCORE", f"{quality_score:.2f}")

    with q2:
        kpi_card("GRADE", grade)

    # =====================================================
    # RECOMMENDATION ENGINE
    # =====================================================
    section("RECOMMENDATION ENGINE")

    recommendations = []

    if risk_level == "CRITICAL":
        recommendations.append(
            "Stop machine and inspect immediately"
        )

    if total_defect > 10:
        recommendations.append(
            "Increase QC sampling frequency"
        )

    if len(mold_ng) > 0:
        recommendations.append(
            f"Inspect mold {critical_mold}"
        )

    if yield_rate < 90:
        recommendations.append(
            "Review production parameters"
        )

    if temp_abnormal > 0.2:
        recommendations.append(
            "Recalibrate machine temperature"
        )

    if pressure_abnormal > 0.2:
        recommendations.append(
            "Check pressure regulator"
        )

    if not recommendations:
        recommendations.append(
            "System stable. No corrective action needed."
        )

    for i, rec in enumerate(recommendations, start=1):
        st.markdown(
            f'<div class="alert-box">{i}. {rec}</div>',
            unsafe_allow_html=True
        )

    # =====================================================
    # HEATMAP MACHINE VS SHIFT
    # =====================================================
    section("DEFECT HEATMAP")

    heat = (
        df_filtered.groupby(["No HP", "Shift"])["is_ng"]
        .sum()
        .reset_index()
    )

    if not heat.empty:
        heatmap = heat.pivot(
            index="No HP",
            columns="Shift",
            values="is_ng"
        ).fillna(0)

        fig_heat = px.imshow(
            heatmap,
            text_auto=True,
            aspect="auto"
        )

        fig_heat.update_layout(height=500)

        st.plotly_chart(
            fig_heat,
            use_container_width=True
        )

    # =====================================================
    # TOP 10 DEFECT MOLD
    # =====================================================
    section("TOP 10 DEFECT MOLD")

    if not mold_ng.empty:
        mold_top = mold_ng.head(10).reset_index()
        mold_top.columns = [
            "Kode Mold",
            "Defect Count"
        ]

        fig_mold = px.bar(
            mold_top,
            x="Kode Mold",
            y="Defect Count",
            text="Defect Count"
        )

        fig_mold.update_layout(height=500)

        st.plotly_chart(
            fig_mold,
            use_container_width=True
        )
    else:
        st.info("Belum ada defect mold.")

# =========================================================
# MANUAL INPUT PAGE
# =========================================================
elif menu == "Manual Input":

    section("MANUAL DATA INPUT")

    with st.form("manual_input", clear_on_submit=True):

        c1, c2, c3 = st.columns(3)

        with c1:
            tanggal = st.date_input("Tanggal")

            shift = st.selectbox(
                "Shift",
                [1, 2, 3]
            )

            no_hp = st.selectbox(
                "No HP",
                [f"HP{i:02d}" for i in range(1, 31)]
            )

        with c2:
            layer_hp = st.selectbox(
                "Layer HP",
                [1, 2, 3, 4, 5]
            )

            kode_mold = st.text_input("Kode Mold")
            no_lot = st.text_input("No Lot")

        with c3:
            keterangan = st.selectbox(
                "Keterangan",
                DEFECT_LIST
            )

            temp = st.number_input(
                "Temp",
                value=200.0
            )

            pressure = st.number_input(
                "Pressure",
                value=100.0
            )

        cycle_time = st.number_input(
            "Cycle Time",
            value=35.0
        )

        operator = st.text_input("Operator")

        submit = st.form_submit_button(
            "Tambah Data"
        )

        if submit:

            new_row = pd.DataFrame([{
                "No": len(st.session_state.db) + 1,
                "Tanggal": tanggal,
                "Shift": shift,
                "No HP": no_hp,
                "Layer HP": layer_hp,
                "Kode Mold": kode_mold,
                "No Lot": no_lot,
                "Keterangan": keterangan,
                "Temp": temp,
                "Pressure": pressure,
                "Cycle Time": cycle_time,
                "Operator": operator
            }])

            st.session_state.db = pd.concat(
                [st.session_state.db, new_row],
                ignore_index=True
            )

            add_log("Manual input data")

            st.success("Data berhasil ditambahkan")
            st.rerun()

# =========================================================
# DATA TABLE PAGE
# =========================================================
elif menu == "Data Table":

    section("DATA TABLE")

    display_df = st.session_state.db.copy()

    st.dataframe(
        display_df,
        use_container_width=True,
        height=500
    )

    # =====================================================
    # DELETE ROW
    # =====================================================
    section("DELETE DATA")

    if len(display_df) > 0:

        delete_no = st.number_input(
            "Masukkan No data yang akan dihapus",
            min_value=1,
            max_value=len(display_df),
            step=1
        )

        if st.button("Delete Row"):

            db = st.session_state.db.copy()

            db = db[db["No"] != delete_no]
            db = db.reset_index(drop=True)
            db["No"] = range(1, len(db) + 1)

            st.session_state.db = db

            add_log(f"Delete row {delete_no}")

            st.success("Row berhasil dihapus")
            st.rerun()

    # =====================================================
    # EXPORT REPORT
    # =====================================================
    section("EXPORT REPORT")

    excel_data = convert_excel(
        st.session_state.db
    )

    st.download_button(
        label="Download Excel Report",
        data=excel_data,
        file_name="quality_trace_report.xlsx",
        mime="application/vnd.ms-excel"
    )

# =========================================================
# AUDIT TRAIL PAGE
# =========================================================
elif menu == "Audit Trail":

    section("AUDIT TRAIL")

    if len(st.session_state.audit_log) > 0:

        audit_df = pd.DataFrame(
            st.session_state.audit_log
        )

        st.dataframe(
            audit_df,
            use_container_width=True,
            height=500
        )

    else:
        st.info("Belum ada audit log.")
