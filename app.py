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
# CUSTOM CSS
# =========================================================
st.markdown("""
<style>

/* MAIN */
.stApp{
    background:#eef2f7;
    font-family:'Segoe UI', sans-serif;
}

/* SIDEBAR */
section[data-testid="stSidebar"]{
    background: linear-gradient(180deg,#071952 0%,#0B2A6F 100%);
    border-right:4px solid #B22222;
}

section[data-testid="stSidebar"] *{
    color:white !important;
}

/* HEADER */
.qtm-header{
    background:white;
    padding:30px;
    border-radius:24px;
    border-left:8px solid #B22222;
    box-shadow:0 8px 25px rgba(0,0,0,0.08);
    margin-bottom:20px;
}

.qtm-title{
    font-size:52px;
    font-weight:900;
    color:#081F5C;
}

.qtm-subtitle{
    font-size:18px;
    color:#666;
}

/* KPI */
.kpi-box{
    background:white;
    padding:25px;
    border-radius:20px;
    box-shadow:0 4px 15px rgba(0,0,0,0.08);
    border-top:5px solid #081F5C;
    text-align:center;
}

.kpi-title{
    font-size:14px;
    color:#777;
    font-weight:700;
}

.kpi-value{
    font-size:34px;
    color:#8B0000;
    font-weight:900;
}

/* ALERT */
.alert-box{
    background:#fff5f5;
    border-left:8px solid #ff4d4f;
    padding:18px;
    border-radius:14px;
    margin-bottom:10px;
}

/* SECTION */
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
# CONSTANTS
# =========================================================
COLUMNS = [
    "No",
    "Tanggal",
    "Shift",
    "No HP",
    "Layer HP",
    "Kode Mold",
    "No Lot",
    "Keterangan",
    "Temp",
    "Pressure",
    "Cycle Time",
    "Operator"
]

DEFECT_LIST = [
    "OK",
    "Visual",
    "Dimensi",
    "Visual Dimensi",
    "Scratch",
    "Bubble",
    "Flash",
    "Short Mold"
]

USERS = {
    "admin": {
        "password": "admin123",
        "role": "Admin"
    },
    "operator": {
        "password": "operator123",
        "role": "Operator"
    },
    "qc": {
        "password": "qc123",
        "role": "QC Inspector"
    },
    "supervisor": {
        "password": "super123",
        "role": "Supervisor"
    }
}

# =========================================================
# SESSION STATE
# =========================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = None

if "role" not in st.session_state:
    st.session_state.role = None

if "db" not in st.session_state:
    st.session_state.db = pd.DataFrame(columns=COLUMNS)

if "audit_log" not in st.session_state:
    st.session_state.audit_log = []

# =========================================================
# AUDIT FUNCTION
# =========================================================
def add_log(action):
    st.session_state.audit_log.append({
        "User": st.session_state.username,
        "Role": st.session_state.role,
        "Action": action,
        "Time": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    })

# =========================================================
# LOGIN PAGE
# =========================================================
if not st.session_state.logged_in:

    st.title("🔐 Login - Quality Trace Monitor")

    user = st.text_input("Username")
    pw = st.text_input("Password", type="password")

    if st.button("Login"):

        if user in USERS and USERS[user]["password"] == pw:
            st.session_state.logged_in = True
            st.session_state.username = user
            st.session_state.role = USERS[user]["role"]
            st.rerun()
        else:
            st.error("Username / password salah")

    st.stop()

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.markdown("# QUALITY TRACE MONITOR")
st.sidebar.markdown(f"""
User: **{st.session_state.username}**  
Role: **{st.session_state.role}**
""")

if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.rerun()

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel",
    type=["xlsx"]
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        for col in COLUMNS:
            if col not in df.columns:
                df[col] = None

        df = df[COLUMNS]
        st.session_state.db = df

        add_log("Upload Excel")

        st.sidebar.success("Upload berhasil")

    except Exception as e:
        st.sidebar.error(e)

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

# =========================================================
# IF NO DATA
# =========================================================
df = st.session_state.db.copy()

if df.empty:
    st.warning("Belum ada data. Upload file Excel terlebih dahulu.")
    st.stop()

# =========================================================
# DATA CLEANING
# =========================================================
df["Tanggal"] = pd.to_datetime(
    df["Tanggal"],
    errors="coerce",
    dayfirst=True
)

df["Keterangan"] = df["Keterangan"].astype(str)

df["is_ok"] = (
    df["Keterangan"]
    .str.upper()
    .eq("OK")
    .astype(int)
)

df["is_ng"] = (
    df["Keterangan"]
    .str.upper()
    .ne("OK")
    .astype(int)
)

# =========================================================
# FILTER SECTION
# =========================================================
st.markdown(
    '<div class="section-title">FILTER DATA</div>',
    unsafe_allow_html=True
)

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

# =========================================================
# KPI CALCULATION
# =========================================================
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

# =========================================================
# KPI DISPLAY
# =========================================================
st.markdown(
    '<div class="section-title">KEY PERFORMANCE INDICATORS</div>',
    unsafe_allow_html=True
)

k1, k2, k3, k4, k5, k6 = st.columns(6)

cards = [
    ("TOTAL PRODUCTION", total_production),
    ("TOTAL OK", total_ok),
    ("TOTAL DEFECT", total_defect),
    ("YIELD RATE", f"{yield_rate:.2f}%"),
    ("CRITICAL MACHINE", critical_machine),
    ("CRITICAL MOLD", critical_mold),
]

for col, (title, value) in zip([k1,k2,k3,k4,k5,k6], cards):
    with col:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """, unsafe_allow_html=True)

# =========================================================
# SMART ALERT CENTER
# =========================================================
st.markdown(
    '<div class="section-title">SMART ALERT CENTER</div>',
    unsafe_allow_html=True
)

alerts = []

machine_defect_rate = (
    df_filtered.groupby("No HP")
    .apply(lambda x: x["is_ng"].sum()/len(x) if len(x)>0 else 0)
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
            f"⚠ Mold {critical_mold} menghasilkan defect tinggi ({top_mold_count})"
        )

if not alerts:
    st.success("🟢 Semua parameter dalam kondisi normal.")
else:
    for alert in alerts:
        st.markdown(f"""
        <div class="alert-box">
            {alert}
        </div>
        """, unsafe_allow_html=True)

# =========================================================
# MONITORING HARIAN
# =========================================================
st.markdown(
    '<div class="section-title">MONITORING HARIAN</div>',
    unsafe_allow_html=True
)

daily = (
    df_filtered.groupby("Tanggal")
    .agg({
        "Layer HP":"count",
        "is_ok":"sum"
    })
    .reset_index()
)

daily.columns = [
    "Tanggal",
    "Production",
    "OK"
]

daily["Yield"] = daily["OK"] / daily["Production"]

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

st.plotly_chart(fig, width="stretch")

# =========================================================
# PARETO DEFECT
# =========================================================
st.markdown(
    '<div class="section-title">PARETO DEFECT ANALYSIS</div>',
    unsafe_allow_html=True
)

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

    fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])

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

    st.plotly_chart(fig_pareto, width="stretch")
else:
    st.info("Belum ada data defect.")

# =========================================================
# ROOT CAUSE ANALYSIS
# =========================================================
st.markdown(
    '<div class="section-title">ROOT CAUSE ANALYSIS</div>',
    unsafe_allow_html=True
)

def get_root_cause(defect_name):
    defect = str(defect_name).lower()

    if "visual" in defect:
        return ("Material contamination", 82)
    elif "dimensi" in defect:
        return ("Mold wear / pressure unstable", 91)
    elif "scratch" in defect:
        return ("Handling issue", 77)
    elif "bubble" in defect:
        return ("Air trapped / cooling issue", 85)
    elif "flash" in defect:
        return ("Clamping pressure low", 88)
    elif "short" in defect:
        return ("Insufficient material fill", 86)
    else:
        return ("Unknown", 50)

ng_df = df_filtered[df_filtered["is_ng"] == 1]

if not ng_df.empty:
    dominant_defect = (
        ng_df["Keterangan"]
        .value_counts()
        .index[0]
    )

    cause, confidence = get_root_cause(dominant_defect)

    rc1, rc2, rc3 = st.columns(3)

    with rc1:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">DOMINANT DEFECT</div>
            <div class="kpi-value">{dominant_defect}</div>
        </div>
        """, unsafe_allow_html=True)

    with rc2:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">ROOT CAUSE</div>
            <div class="kpi-value" style="font-size:20px">
                {cause}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with rc3:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">CONFIDENCE</div>
            <div class="kpi-value">{confidence}%</div>
        </div>
        """, unsafe_allow_html=True)

else:
    st.success("Tidak ada defect. Root cause analysis tidak diperlukan.")

# =========================================================
# PREDICTION ENGINE
# =========================================================
st.markdown(
    '<div class="section-title">PREDICTION ENGINE</div>',
    unsafe_allow_html=True
)

machine_rate = (
    total_defect / total_production
    if total_production > 0 else 0
)

mold_rate = (
    mold_ng.iloc[0] / total_production
    if len(mold_ng) > 0 and total_production > 0 else 0
)

temp_abnormal = 0
pressure_abnormal = 0

if "Temp" in df_filtered.columns:
    temp_abnormal = (
        ((pd.to_numeric(df_filtered["Temp"], errors="coerce") > 230) |
         (pd.to_numeric(df_filtered["Temp"], errors="coerce") < 180))
        .sum()
    ) / total_production if total_production > 0 else 0

if "Pressure" in df_filtered.columns:
    pressure_abnormal = (
        ((pd.to_numeric(df_filtered["Pressure"], errors="coerce") > 120) |
         (pd.to_numeric(df_filtered["Pressure"], errors="coerce") < 80))
        .sum()
    ) / total_production if total_production > 0 else 0

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

pr1, pr2 = st.columns(2)

with pr1:
    st.metric("Risk Score", f"{risk_score:.2f}")

with pr2:
    st.metric("Risk Level", risk_level)

# =========================================================
# AI QUALITY SCORE
# =========================================================
st.markdown(
    '<div class="section-title">AI QUALITY SCORE</div>',
    unsafe_allow_html=True
)

quality_score = (
    yield_rate * 0.7 +
    max(0, 100-risk_score) * 0.3
)

if quality_score >= 90:
    quality_grade = "A"
elif quality_score >= 80:
    quality_grade = "B"
elif quality_score >= 70:
    quality_grade = "C"
else:
    quality_grade = "D"

q1, q2 = st.columns(2)

with q1:
    st.metric("Quality Score", f"{quality_score:.2f}")

with q2:
    st.metric("Grade", quality_grade)

# =========================================================
# RECOMMENDATION ENGINE
# =========================================================
st.markdown(
    '<div class="section-title">RECOMMENDATION ENGINE</div>',
    unsafe_allow_html=True
)

recommendations = []

if risk_level == "CRITICAL":
    recommendations.append("Stop machine and perform inspection immediately")

if total_defect > 10:
    recommendations.append("Increase QC sampling frequency")

if len(mold_ng) > 0:
    recommendations.append(f"Inspect mold {critical_mold}")

if yield_rate < 90:
    recommendations.append("Review production parameters")

if temp_abnormal > 0.2:
    recommendations.append("Recalibrate machine temperature")

if pressure_abnormal > 0.2:
    recommendations.append("Check pressure regulator")

if not recommendations:
    recommendations.append("System stable. No corrective action needed.")

for i, rec in enumerate(recommendations, start=1):
    st.markdown(f"""
    <div class="alert-box">
        {i}. {rec}
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# HEATMAP MACHINE VS SHIFT
# =========================================================
st.markdown(
    '<div class="section-title">DEFECT HEATMAP</div>',
    unsafe_allow_html=True
)

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

    st.plotly_chart(fig_heat, width="stretch")

# =========================================================
# TOP 10 DEFECT MOLD
# =========================================================
st.markdown(
    '<div class="section-title">TOP 10 DEFECT MOLD</div>',
    unsafe_allow_html=True
)

if not mold_ng.empty:
    mold_top = mold_ng.head(10).reset_index()
    mold_top.columns = ["Kode Mold", "Defect Count"]

    fig_mold = px.bar(
        mold_top,
        x="Kode Mold",
        y="Defect Count",
        text="Defect Count"
    )

    fig_mold.update_layout(height=500)

    st.plotly_chart(fig_mold, width="stretch")
else:
    st.info("Belum ada defect mold.")

# =========================================================
# MANUAL INPUT FORM
# =========================================================
st.markdown(
    '<div class="section-title">MANUAL DATA INPUT</div>',
    unsafe_allow_html=True
)

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
            [f"HP{i:02d}" for i in range(1,31)]
        )

    with c2:
        layer_hp = st.selectbox(
            "Layer HP",
            [1,2,3,4,5]
        )

        kode_mold = st.text_input("Kode Mold")
        no_lot = st.text_input("No Lot")

    with c3:
        keterangan = st.selectbox(
            "Keterangan",
            DEFECT_LIST
        )

        temp = st.number_input("Temp", value=200.0)
        pressure = st.number_input("Pressure", value=100.0)

    cycle_time = st.number_input(
        "Cycle Time",
        value=35.0
    )

    operator = st.text_input("Operator")

    submit = st.form_submit_button("Tambah Data")

    if submit:

        new_row = pd.DataFrame([{
            "No": len(st.session_state.db)+1,
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
# DATA TABLE
# =========================================================
st.markdown(
    '<div class="section-title">DATA TABLE</div>',
    unsafe_allow_html=True
)

display_df = st.session_state.db.copy()

st.dataframe(
    display_df,
    use_container_width=True,
    height=400
)

# =========================================================
# DELETE DATA
# =========================================================
st.markdown(
    '<div class="section-title">DELETE DATA</div>',
    unsafe_allow_html=True
)

delete_no = st.number_input(
    "Masukkan No data yang akan dihapus",
    min_value=1,
    step=1
)

if st.button("Delete Row"):

    db = st.session_state.db.copy()

    db = db[db["No"] != delete_no]
    db["No"] = range(1, len(db)+1)

    st.session_state.db = db

    add_log(f"Delete row {delete_no}")

    st.success("Row berhasil dihapus")
    st.rerun()

# =========================================================
# AUDIT TRAIL
# =========================================================
st.markdown(
    '<div class="section-title">AUDIT TRAIL</div>',
    unsafe_allow_html=True
)

if len(st.session_state.audit_log) > 0:
    audit_df = pd.DataFrame(
        st.session_state.audit_log
    )

    st.dataframe(
        audit_df,
        use_container_width=True
    )
else:
    st.info("Belum ada audit log.")

# =========================================================
# DOWNLOAD DATA
# =========================================================
st.markdown(
    '<div class="section-title">EXPORT REPORT</div>',
    unsafe_allow_html=True
)

def convert_excel(df):
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:
        df.to_excel(
            writer,
            index=False
        )

    return output.getvalue()

excel_data = convert_excel(st.session_state.db)

st.download_button(
    label="Download Excel Report",
    data=excel_data,
    file_name="quality_trace_report.xlsx",
    mime="application/vnd.ms-excel"
)

# =========================================================
# DUMMY DATA GENERATOR
# =========================================================
st.markdown(
    '<div class="section-title">GENERATE DUMMY DATA</div>',
    unsafe_allow_html=True
)

if st.button("Generate 300 Dummy Rows"):

    dummy = []

    defects = [
        "OK",
        "Visual",
        "Dimensi",
        "Flash",
        "Bubble",
        "Scratch",
        "Short Mold"
    ]

    operators = [
        "Andi",
        "Budi",
        "Sinta",
        "Rina",
        "Dewi"
    ]

    for i in range(300):

        defect = np.random.choice(
            defects,
            p=[0.75,0.06,0.06,0.04,0.03,0.03,0.03]
        )

        dummy.append({
            "No": i+1,
            "Tanggal": datetime.today(),
            "Shift": np.random.randint(1,4),
            "No HP": f"HP{np.random.randint(1,30):02d}",
            "Layer HP": np.random.randint(1,6),
            "Kode Mold": f"M{np.random.randint(100,999)}",
            "No Lot": f"LOT-{np.random.randint(1000,9999)}",
            "Keterangan": defect,
            "Temp": np.random.randint(175,235),
            "Pressure": np.random.randint(75,125),
            "Cycle Time": np.random.randint(25,45),
            "Operator": np.random.choice(operators)
        })

    st.session_state.db = pd.DataFrame(dummy)

    add_log("Generate dummy data")

    st.success("300 dummy rows generated")
    st.rerun()
