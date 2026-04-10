import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import timedelta
import html
import base64

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="P2RIS Automation Portal",
    page_icon="🤖",
    layout="wide"
)

# =========================
# CONFIG
# =========================
ADMIN_PASSWORD = "freestyler"

BASE_DIR = Path(".")
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

DB_FILE = DATA_DIR / "current_database.xlsx"
HEADER_FILE = BASE_DIR / "header.png"

# =========================
# HELPERS
# =========================
def file_to_base64(file_path):
    file_path = Path(file_path)
    if not file_path.exists():
        return ""
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def find_column(columns, target_names):
    normalized = {str(col).strip().lower(): col for col in columns}
    for name in target_names:
        key = name.strip().lower()
        if key in normalized:
            return normalized[key]
    return None


def safe_value(row, col_name):
    if not col_name:
        return ""
    value = row.get(col_name, "")
    if pd.isna(value):
        return ""
    return str(value)


@st.cache_data(show_spinner=False)
def load_excel_all_sheets(file_path):
    return pd.read_excel(file_path, sheet_name=None)


def save_uploaded_file(uploaded_file, save_path):
    with open(save_path, "wb") as f:
        f.write(uploaded_file.getbuffer())


def build_master_dataframe(excel_data):
    combined_rows = []

    for sheet_name, df in excel_data.items():
        if df is None or df.empty:
            continue

        df = df.copy()
        df.columns = [str(col).strip() for col in df.columns]

        col_order = find_column(df.columns, ["Order No", "Order Number", "OrderNo"])
        col_building = find_column(df.columns, ["Building"])
        col_order_type = find_column(df.columns, ["Order Type", "OrderType"])
        col_state = find_column(df.columns, ["State"])
        col_iris = find_column(df.columns, ["Iris Ticket", "IRIS Ticket", "Iris"])
        col_created = find_column(
            df.columns,
            ["Created Date", "CreatedDate", "Create Date", "Open Time", "Date Created"]
        )

        if not col_order:
            continue

        for _, row in df.iterrows():
            combined_rows.append({
                "Order No": safe_value(row, col_order),
                "Building": safe_value(row, col_building),
                "Order Type": safe_value(row, col_order_type),
                "State": safe_value(row, col_state),
                "Iris Ticket": safe_value(row, col_iris),
                "Created Date": safe_value(row, col_created),
                "Issue": str(sheet_name)
            })

    if combined_rows:
        df_master = pd.DataFrame(combined_rows)
    else:
        df_master = pd.DataFrame(columns=[
            "Order No", "Building", "Order Type", "State",
            "Iris Ticket", "Created Date", "Issue"
        ])

    if "Created Date" in df_master.columns:
        df_master["Created Date Parsed"] = pd.to_datetime(df_master["Created Date"], errors="coerce")
    else:
        df_master["Created Date Parsed"] = pd.NaT

    return df_master


def get_latest_month_df(df):
    if df.empty or "Created Date Parsed" not in df.columns:
        return df.copy(), None

    temp = df.copy().dropna(subset=["Created Date Parsed"])

    if temp.empty:
        return df.head(0).copy(), None

    latest_date = temp["Created Date Parsed"].max()
    latest_year = latest_date.year
    latest_month = latest_date.month

    month_df = temp[
        (temp["Created Date Parsed"].dt.year == latest_year) &
        (temp["Created Date Parsed"].dt.month == latest_month)
    ].copy()

    return month_df, latest_date


def prepare_latest_week_task_dataset(excel_data):
    target_sheet_map = {
        "P.PROCESSING": "P.PROCESSING",
        "UNSYNC": "UNSYNC",
        "B.PASS EXTRA PORT": "B.PASS EXTRA PORT"
    }

    all_rows = []

    for sheet_name, df in excel_data.items():
        if df is None or df.empty:
            continue

        clean_sheet_name = str(sheet_name).strip().upper()

        matched_category = None
        for key in target_sheet_map:
            if key in clean_sheet_name:
                matched_category = target_sheet_map[key]
                break

        if not matched_category:
            continue

        df = df.copy()
        df.columns = [str(col).strip() for col in df.columns]

        col_created = find_column(
            df.columns,
            ["Created Date", "CreatedDate", "Create Date", "Open Time", "Date Created"]
        )
        col_state = find_column(df.columns, ["State"])

        if not col_created:
            continue

        temp = pd.DataFrame()
        temp["Created Date"] = pd.to_datetime(df[col_created], errors="coerce")
        temp["State"] = df[col_state].astype(str).str.strip() if col_state else "UNKNOWN"
        temp["Category"] = matched_category
        temp = temp.dropna(subset=["Created Date"])

        if not temp.empty:
            all_rows.append(temp)

    if not all_rows:
        return pd.DataFrame(), None, None, "Target trend sheets not found or Created Date column missing."

    trend_df = pd.concat(all_rows, ignore_index=True)

    latest_date = trend_df["Created Date"].max()
    if pd.isna(latest_date):
        return pd.DataFrame(), None, None, "No valid Created Date found in target sheets."

    latest_date = pd.Timestamp(latest_date).normalize()

    weekday_num = latest_date.weekday()  # Mon=0 ... Sun=6
    days_since_sunday = (weekday_num + 1) % 7
    week_start = latest_date - timedelta(days=days_since_sunday)
    week_end = week_start + timedelta(days=6)

    trend_df["DateOnly"] = trend_df["Created Date"].dt.normalize()
    trend_df = trend_df[
        (trend_df["DateOnly"] >= week_start) &
        (trend_df["DateOnly"] <= week_end)
    ].copy()

    weekday_name_map = {
        6: "Sun",
        0: "Mon",
        1: "Tue",
        2: "Wed",
        3: "Thu",
        4: "Fri",
        5: "Sat"
    }

    trend_df["WeekdayNum"] = trend_df["DateOnly"].dt.weekday
    trend_df["Day"] = trend_df["WeekdayNum"].map(weekday_name_map)

    return trend_df, week_start, week_end, "Latest current week trend based on latest available date in database."


def build_weekday_chart_df(task_df):
    if task_df.empty:
        return pd.DataFrame(columns=["Day", "P.PROCESSING", "UNSYNC", "B.PASS EXTRA PORT"])

    day_order = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    cat_order = ["P.PROCESSING", "UNSYNC", "B.PASS EXTRA PORT"]

    chart_df = (
        task_df.groupby(["Day", "Category"])
        .size()
        .unstack(fill_value=0)
    )

    chart_df = chart_df.reindex(day_order, fill_value=0)
    chart_df = chart_df.reindex(columns=cat_order, fill_value=0)
    chart_df.index.name = "Day"
    chart_df = chart_df.reset_index()

    return chart_df


def build_current_month_state_summary(master_df):
    month_df, latest_month_date = get_latest_month_df(master_df)

    if month_df.empty:
        return pd.DataFrame(columns=["State", "Total"]), latest_month_date

    state_df = (
        month_df.groupby("State")
        .size()
        .reset_index(name="Total")
        .sort_values(["Total", "State"], ascending=[False, True])
        .reset_index(drop=True)
    )

    return state_df, latest_month_date


def build_all_time_summary(master_df):
    if master_df.empty:
        return {
            "total_all_time": 0,
            "p_processing": 0,
            "unsync": 0,
            "bypass": 0,
            "states": 0
        }

    temp = master_df.copy()

    p_processing = int((temp["Issue"].astype(str).str.upper() == "P.PROCESSING").sum())
    unsync = int((temp["Issue"].astype(str).str.upper() == "UNSYNC").sum())
    bypass = int((temp["Issue"].astype(str).str.upper() == "B.PASS EXTRA PORT").sum())
    total_all_time = p_processing + unsync + bypass

    return {
        "total_all_time": total_all_time,
        "p_processing": p_processing,
        "unsync": unsync,
        "bypass": bypass,
        "states": temp["State"].nunique() if "State" in temp.columns else 0
    }


def build_current_month_summary(master_df):
    month_df, latest_month_date = get_latest_month_df(master_df)

    if month_df.empty:
        return {
            "month_df": month_df,
            "latest_month_date": latest_month_date,
            "total_month": 0,
            "states": 0,
            "peak_state": "-"
        }

    state_count = month_df.groupby("State").size().reset_index(name="Total")
    peak_state = state_count.sort_values(["Total", "State"], ascending=[False, True]).iloc[0]["State"]

    return {
        "month_df": month_df,
        "latest_month_date": latest_month_date,
        "total_month": len(month_df),
        "states": month_df["State"].nunique(),
        "peak_state": peak_state
    }


def build_receipt_text(row_dict):
    return (
        f"Order No: {row_dict.get('Order No', '')}\n"
        f"Building: {row_dict.get('Building', '')}\n"
        f"Order Type: {row_dict.get('Order Type', '')}\n"
        f"State: {row_dict.get('State', '')}\n"
        f"Iris Ticket: {row_dict.get('Iris Ticket', '')}\n"
        f"Created Date: {row_dict.get('Created Date', '')}\n"
        f"Issue: {row_dict.get('Issue', '')}"
    )


def render_copy_button(text):
    safe_text = html.escape(text).replace("\n", "\\n").replace("'", "\\'")
    st.markdown(f"""
    <div class="copy-wrap">
        <button class="copy-btn" onclick="navigator.clipboard.writeText('{safe_text}')">
            📋 Copy Result for WhatsApp
        </button>
    </div>
    """, unsafe_allow_html=True)


def render_main_result(row_dict, result_count):
    st.markdown(f"""
    <div class="result-highlight">
        <div class="result-header">Main Search Result {'' if result_count == 1 else f'(Top match from {result_count} results)'}</div>
        <div class="result-grid">
            <div class="result-item">
                <div class="result-label">Order No</div>
                <div class="result-value">{row_dict.get('Order No', '')}</div>
            </div>
            <div class="result-item">
                <div class="result-label">Building</div>
                <div class="result-value">{row_dict.get('Building', '')}</div>
            </div>
            <div class="result-item">
                <div class="result-label">Order Type</div>
                <div class="result-value">{row_dict.get('Order Type', '')}</div>
            </div>
            <div class="result-item">
                <div class="result-label">State</div>
                <div class="result-value">{row_dict.get('State', '')}</div>
            </div>
            <div class="result-item">
                <div class="result-label">Iris Ticket</div>
                <div class="result-value">{row_dict.get('Iris Ticket', '')}</div>
            </div>
            <div class="result-item">
                <div class="result-label">Created Date</div>
                <div class="result-value">{row_dict.get('Created Date', '')}</div>
            </div>
            <div class="result-item" style="grid-column: 1 / -1;">
                <div class="result-label">Issue Source</div>
                <div class="result-value">{row_dict.get('Issue', '')}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    receipt_text = build_receipt_text(row_dict)
    st.markdown("#### WhatsApp Receipt Format")
    st.code(receipt_text, language=None)
    render_copy_button(receipt_text)


def render_result_card(row_dict, title="Additional Result"):
    st.markdown(f"""
    <div class="additional-card">
        <div class="additional-title">{title}</div>
        <div class="additional-grid">
            <div><b>Order No:</b> {row_dict.get('Order No', '')}</div>
            <div><b>Building:</b> {row_dict.get('Building', '')}</div>
            <div><b>Order Type:</b> {row_dict.get('Order Type', '')}</div>
            <div><b>State:</b> {row_dict.get('State', '')}</div>
            <div><b>Iris Ticket:</b> {row_dict.get('Iris Ticket', '')}</div>
            <div><b>Created Date:</b> {row_dict.get('Created Date', '')}</div>
            <div style="grid-column: 1 / -1;"><b>Issue:</b> {row_dict.get('Issue', '')}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_centered_table(df):
    html_table = '<table class="custom-table">'
    html_table += '<thead><tr>'
    for col in df.columns:
        html_table += f'<th>{col}</th>'
    html_table += '</tr></thead><tbody>'

    for _, row in df.iterrows():
        html_table += '<tr>'
        for val in row:
            html_table += f'<td>{val}</td>'
        html_table += '</tr>'

    html_table += '</tbody></table>'
    st.markdown(html_table, unsafe_allow_html=True)


def render_weekly_summary(chart_df):
    if chart_df.empty:
        return

    total_weekly = int(chart_df[["P.PROCESSING", "UNSYNC", "B.PASS EXTRA PORT"]].sum().sum())
    p_processing_total = int(chart_df["P.PROCESSING"].sum())
    unsync_total = int(chart_df["UNSYNC"].sum())
    bypass_total = int(chart_df["B.PASS EXTRA PORT"].sum())

    peak_series = chart_df.copy()
    peak_series["Daily Total"] = peak_series[["P.PROCESSING", "UNSYNC", "B.PASS EXTRA PORT"]].sum(axis=1)
    peak_row = peak_series.sort_values("Daily Total", ascending=False).iloc[0]
    peak_day = peak_row["Day"]
    peak_value = int(peak_row["Daily Total"])

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Weekly", total_weekly)
    c2.metric("P.PROCESSING", p_processing_total)
    c3.metric("UNSYNC", unsync_total)
    c4.metric("B.PASS EXTRA PORT", bypass_total)
    c5.metric("Peak Day", f"{peak_day} ({peak_value})")


def check_admin_password():
    st.markdown('<div class="soft-card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">🔐 Admin Login</div>', unsafe_allow_html=True)
    st.caption("Admin Portal dilindungi password.")

    password = st.text_input(
        "Enter Admin Password",
        type="password",
        placeholder="Enter password"
    )

    if st.button("Login Admin", use_container_width=True):
        if password == ADMIN_PASSWORD:
            st.session_state["admin_authenticated"] = True
            st.success("Admin access granted.")
            st.rerun()
        else:
            st.error("Wrong password.")

    st.markdown("</div>", unsafe_allow_html=True)

# =========================
# CSS
# =========================
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: "Segoe UI", sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(114, 182, 255, 0.10), transparent 25%),
        radial-gradient(circle at top right, rgba(168, 214, 255, 0.10), transparent 22%),
        linear-gradient(135deg, #f7fbff 0%, #f4f9fd 50%, #f9fcff 100%);
}

.block-container {
    max-width: 1260px;
    padding-top: 1.0rem;
    padding-bottom: 2rem;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f8fbff 0%, #eef5fb 100%);
    border-right: 1px solid #dce8f2;
}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #21486B;
}

.header-full {
    position: relative;
    width: 100%;
    height: 230px;
    border-radius: 24px;
    overflow: hidden;
    margin-bottom: 0.9rem;
    box-shadow: 0 14px 30px rgba(60,100,140,0.16);
    border: 1px solid rgba(210, 228, 241, 0.6);
    background: #0f1d2b;
}

.header-full img {
    width: 100%;
    height: 100%;
    object-fit: cover;
    display: block;
}

.header-status {
    position: absolute;
    right: 20px;
    top: 16px;
    background: rgba(8,16,28,0.55);
    color: #f4fbff;
    border: 1px solid rgba(255,255,255,0.18);
    padding: 7px 12px;
    border-radius: 999px;
    font-size: 0.82rem;
    font-weight: 700;
    backdrop-filter: blur(4px);
}

.header-info-bar {
    background: rgba(255,255,255,0.94);
    border: 1px solid #E1EDF6;
    border-radius: 20px;
    padding: 16px 20px;
    box-shadow: 0 10px 24px rgba(78, 120, 155, 0.06);
    margin-bottom: 1rem;
}

.header-info-title {
    font-size: 2rem;
    font-weight: 800;
    color: #21486B;
    line-height: 1.1;
    margin-bottom: 4px;
}

.header-info-subtitle {
    font-size: 1rem;
    color: #6A8398;
    margin-bottom: 8px;
}

.header-info-badge {
    display: inline-block;
    font-size: 0.85rem;
    color: #5D7893;
    background: #EEF5FB;
    padding: 6px 12px;
    border-radius: 999px;
    border: 1px solid #DCEAF5;
}

.soft-card {
    background: rgba(255,255,255,0.93);
    border: 1px solid #E1EDF6;
    border-radius: 22px;
    padding: 1.1rem 1.1rem;
    box-shadow: 0 10px 26px rgba(78, 120, 155, 0.07);
    margin-bottom: 1rem;
}

.card-title {
    font-size: 1.2rem;
    font-weight: 700;
    color: #284C6E;
    margin-bottom: 0.9rem;
}

.result-highlight {
    background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
    border: 1px solid #DCEAF5;
    border-radius: 22px;
    padding: 1rem;
    box-shadow: 0 8px 18px rgba(80, 120, 155, 0.06);
    margin-bottom: 0.9rem;
}

.result-header {
    font-size: 1.1rem;
    font-weight: 800;
    color: #21486B;
    margin-bottom: 0.9rem;
}

.result-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(180px, 1fr));
    gap: 0.8rem;
}

.result-item {
    background: #FBFDFF;
    border: 1px solid #E6F0F8;
    border-radius: 15px;
    padding: 0.78rem 0.88rem;
}

.result-label {
    font-size: 0.8rem;
    color: #6D879D;
    font-weight: 600;
    margin-bottom: 0.15rem;
}

.result-value {
    font-size: 1rem;
    color: #284866;
    font-weight: 700;
    word-break: break-word;
}

.additional-card {
    background: #fbfdff;
    border: 1px solid #e3eef7;
    border-radius: 16px;
    padding: 0.9rem;
    margin-bottom: 0.8rem;
}

.additional-title {
    font-weight: 800;
    color: #284c6e;
    margin-bottom: 0.6rem;
}

.additional-grid {
    display: grid;
    grid-template-columns: repeat(2, minmax(180px, 1fr));
    gap: 0.5rem 1rem;
    color: #345676;
}

[data-testid="stTextInput"] input {
    border-radius: 16px !important;
    border: 1px solid #D9E8F3 !important;
    background: rgba(255,255,255,0.98) !important;
    padding: 0.95rem 1rem !important;
    font-size: 1rem !important;
    color: #23496C !important;
    box-shadow: 0 3px 10px rgba(88, 126, 154, 0.05);
}

[data-testid="stFileUploader"] {
    border-radius: 16px !important;
}

[data-testid="stMetric"] {
    background: rgba(255,255,255,0.93);
    border: 1px solid #E1EDF6;
    padding: 0.8rem;
    border-radius: 18px;
    box-shadow: 0 6px 14px rgba(78, 120, 155, 0.05);
}

.copy-wrap {
    margin-top: 0.7rem;
    margin-bottom: 0.3rem;
}

.copy-btn {
    background: linear-gradient(180deg, #EFF7FD 0%, #E4F1FB 100%);
    color: #21486B;
    border: 1px solid #CFE3F3;
    border-radius: 12px;
    padding: 10px 14px;
    font-weight: 700;
    cursor: pointer;
    width: 100%;
}

.copy-btn:hover {
    background: linear-gradient(180deg, #E4F1FB 0%, #DAECFA 100%);
}

.custom-table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    border-radius: 14px;
    overflow: hidden;
    font-size: 14px;
}

.custom-table th {
    background-color: #F4F9FD;
    color: #345676;
    font-weight: 700;
    text-align: center;
    padding: 10px;
    border: 1px solid #E3EEF7;
}

.custom-table td {
    padding: 10px;
    border: 1px solid #E3EEF7;
}

.custom-table td:first-child {
    text-align: left;
    font-weight: 600;
    color: #284866;
}

.custom-table td:not(:first-child) {
    text-align: center;
    font-weight: 700;
    color: #284866;
}

.table-title {
    font-size: 1rem;
    font-weight: 700;
    color: #345676;
    margin: 0.2rem 0 0.6rem 0;
}

.dashboard-note {
    color: #4F6D88;
    font-size: 0.92rem;
}

@media (max-width: 900px) {
    .header-full {
        height: 170px;
    }

    .header-info-title {
        font-size: 1.5rem;
    }

    .result-grid,
    .additional-grid {
        grid-template-columns: 1fr;
    }
}
</style>
""", unsafe_allow_html=True)

# =========================
# SESSION
# =========================
if "admin_authenticated" not in st.session_state:
    st.session_state["admin_authenticated"] = False

# =========================
# SIDEBAR
# =========================
with st.sidebar:
    st.markdown("## P2RIS Portal")
    page = st.radio(
        "Navigation",
        ["User Portal", "Admin Portal"],
        label_visibility="collapsed"
    )

    st.markdown("---")
    st.markdown("### Database Status")
    if DB_FILE.exists():
        st.success("Database available")
        st.caption(f"File: {DB_FILE.name}")
    else:
        st.warning("No database uploaded yet")

    if page == "Admin Portal" and st.session_state.get("admin_authenticated"):
        if st.button("Logout Admin", use_container_width=True):
            st.session_state["admin_authenticated"] = False
            st.rerun()

# =========================
# HEADER
# =========================
header_b64 = file_to_base64(HEADER_FILE)
header_status = "🟢 LIVE DATABASE" if DB_FILE.exists() else "🟡 NO DATABASE"

if header_b64:
    st.markdown(f"""
    <div class="header-full">
        <img src="data:image/png;base64,{header_b64}">
        <div class="header-status">{header_status}</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="header-info-bar">
        <div class="header-info-title">P2RIS AUTOMATION PORTAL</div>
        <div class="header-info-subtitle">AI-assisted Order Tracking System</div>
        <div class="header-info-badge">Automation • Search • Intelligence</div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="soft-card">
        <div class="card-title">P2RIS AUTOMATION PORTAL</div>
        <div class="dashboard-note">AI-assisted Order Tracking System</div>
    </div>
    """, unsafe_allow_html=True)

# =========================
# USER PORTAL
# =========================
if page == "User Portal":
    st.markdown('<div class="soft-card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">🔎 Search Order</div>', unsafe_allow_html=True)

    order_input = st.text_input(
        "",
        placeholder="Enter Orders No (e.g. 2512345678901234)",
        label_visibility="collapsed"
    ).strip()

    st.markdown("</div>", unsafe_allow_html=True)

    if not DB_FILE.exists():
        st.info("Database belum tersedia. Admin perlu upload Excel database dahulu.")
    else:
        try:
            excel_data = load_excel_all_sheets(DB_FILE)
            master_df = build_master_dataframe(excel_data)

            all_time_summary = build_all_time_summary(master_df)
            month_summary = build_current_month_summary(master_df)
            month_state_df, latest_month_date = build_current_month_state_summary(master_df)

            if order_input:
                result_df = master_df[
                    master_df["Order No"].astype(str).str.contains(order_input, case=False, na=False)
                ].copy()

                if not result_df.empty:
                    st.markdown('<div class="soft-card">', unsafe_allow_html=True)
                    st.markdown('<div class="card-title">✅ Search Result</div>', unsafe_allow_html=True)

                    top_result = result_df.iloc[0].to_dict()
                    render_main_result(top_result, len(result_df))

                    if len(result_df) > 1:
                        with st.expander(f"View {len(result_df) - 1} more matched result(s)"):
                            for i in range(1, len(result_df)):
                                row = result_df.iloc[i].to_dict()
                                render_result_card(row, title=f"Additional Result {i}")

                    st.markdown("</div>", unsafe_allow_html=True)
                else:
                    st.markdown('<div class="soft-card">', unsafe_allow_html=True)
                    st.warning("No matching order found.")
                    st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="soft-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">📌 All Time RPA Summary</div>', unsafe_allow_html=True)

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Total IRIS Created", all_time_summary["total_all_time"])
            c2.metric("P.PROCESSING", all_time_summary["p_processing"])
            c3.metric("UNSYNC", all_time_summary["unsync"])
            c4.metric("B.PASS EXTRA PORT", all_time_summary["bypass"])
            c5.metric("States Covered", all_time_summary["states"])

            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="soft-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">📅 Current Month Summary</div>', unsafe_allow_html=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("Current Month Records", month_summary["total_month"])
            c2.metric("States This Month", month_summary["states"])
            c3.metric("Peak State", month_summary["peak_state"])

            if month_summary["latest_month_date"] is not None:
                st.caption(f"Monthly reference: {month_summary['latest_month_date'].strftime('%B %Y')}")

            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="soft-card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">📊 Weekly IRIS Creation Dashboard</div>', unsafe_allow_html=True)

            task_df, week_start, week_end, trend_note = prepare_latest_week_task_dataset(excel_data)
            chart_df = build_weekday_chart_df(task_df)

            st.caption(trend_note)
            if week_start is not None and week_end is not None:
                st.caption(f"Week Range: {week_start.strftime('%d/%m/%Y')} - {week_end.strftime('%d/%m/%Y')}")

            if not chart_df.empty:
                render_weekly_summary(chart_df)
                st.markdown("<br>", unsafe_allow_html=True)
                st.line_chart(chart_df.set_index("Day"))

                st.markdown('<div class="table-title">State vs Total Orders (Current Month Only)</div>', unsafe_allow_html=True)
                render_centered_table(month_state_df)
            else:
                st.info("Trend belum dapat dipaparkan. Pastikan sheet P.PROCESSING, UNSYNC, dan B.PASS EXTRA PORT ada serta mempunyai Created Date.")

            st.markdown("</div>", unsafe_allow_html=True)

        except ImportError:
            st.error("Module openpyxl belum install. Run: python -m pip install openpyxl")
        except Exception as e:
            st.error(f"Error reading database: {e}")

# =========================
# ADMIN PORTAL
# =========================
elif page == "Admin Portal":
    if not st.session_state.get("admin_authenticated"):
        check_admin_password()
    else:
        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">🛠️ Admin Database Management</div>', unsafe_allow_html=True)
        st.markdown('<div class="dashboard-note">Page ini untuk admin upload dan refresh Excel database semasa.</div>', unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Upload Excel Database",
            type=["xlsx", "xls"]
        )

        col_a, col_b = st.columns([1, 1])

        with col_a:
            if uploaded_file is not None:
                try:
                    save_uploaded_file(uploaded_file, DB_FILE)
                    st.success("Database uploaded successfully.")
                    st.caption(f"Saved as: {DB_FILE}")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Upload failed: {e}")

        with col_b:
            if DB_FILE.exists():
                st.info(f"Current active database: {DB_FILE.name}")
            else:
                st.warning("No active database yet.")

        st.markdown("</div>", unsafe_allow_html=True)

        if DB_FILE.exists():
            try:
                excel_data = load_excel_all_sheets(DB_FILE)
                master_df = build_master_dataframe(excel_data)

                all_time_summary = build_all_time_summary(master_df)
                month_summary = build_current_month_summary(master_df)
                month_state_df, latest_month_date = build_current_month_state_summary(master_df)

                st.markdown('<div class="soft-card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📈 Admin All Time Summary</div>', unsafe_allow_html=True)

                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Total IRIS Created", all_time_summary["total_all_time"])
                c2.metric("P.PROCESSING", all_time_summary["p_processing"])
                c3.metric("UNSYNC", all_time_summary["unsync"])
                c4.metric("B.PASS EXTRA PORT", all_time_summary["bypass"])
                c5.metric("States Covered", all_time_summary["states"])

                st.markdown("</div>", unsafe_allow_html=True)

                st.markdown('<div class="soft-card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📅 Admin Current Month Summary</div>', unsafe_allow_html=True)

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Current Month Records", month_summary["total_month"])
                c2.metric("Total Sheets", len(excel_data))
                c3.metric("States This Month", month_summary["states"])
                c4.metric("Peak State", month_summary["peak_state"])

                if month_summary["latest_month_date"] is not None:
                    st.caption(f"Monthly reference: {month_summary['latest_month_date'].strftime('%B %Y')}")

                st.markdown("</div>", unsafe_allow_html=True)

                st.markdown('<div class="soft-card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📊 Weekly IRIS Creation Dashboard</div>', unsafe_allow_html=True)

                task_df, week_start, week_end, trend_note = prepare_latest_week_task_dataset(excel_data)
                chart_df = build_weekday_chart_df(task_df)

                st.caption(trend_note)
                if week_start is not None and week_end is not None:
                    st.caption(f"Week Range: {week_start.strftime('%d/%m/%Y')} - {week_end.strftime('%d/%m/%Y')}")

                if not chart_df.empty:
                    render_weekly_summary(chart_df)
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.line_chart(chart_df.set_index("Day"))

                    st.markdown('<div class="table-title">State vs Total Orders (Current Month Only)</div>', unsafe_allow_html=True)
                    render_centered_table(month_state_df)
                else:
                    st.info("Trend belum dapat dijana. Pastikan sheet P.PROCESSING, UNSYNC, dan B.PASS EXTRA PORT ada serta mempunyai Created Date.")

                st.markdown("</div>", unsafe_allow_html=True)

                st.markdown('<div class="soft-card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">🧾 Preview Database (Current Month Only)</div>', unsafe_allow_html=True)

                preview_df = month_summary["month_df"].copy()

                if "Created Date Parsed" in preview_df.columns:
                    preview_df = preview_df.sort_values(
                        "Created Date Parsed",
                        ascending=False,
                        na_position="last"
                    )

                preview_df = preview_df.drop(columns=["Created Date Parsed"], errors="ignore")

                st.dataframe(
                    preview_df.head(50),
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("</div>", unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Failed to load active database: {e}")

# =========================
# FOOTER
# =========================
st.markdown("""
<div style="text-align:center; color:#6f879b; margin-top:1rem; font-size:0.92rem;">
P2RIS Automation Portal • Internal Search Interface
</div>
""", unsafe_allow_html=True)