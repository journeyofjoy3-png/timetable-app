import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import os
import time
import sqlite3
import hashlib
import base64
import plotly.express as px
from streamlit_option_menu import option_menu
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils.dataframe import dataframe_to_rows

# --- Password Hashing Helper ---
def make_hashes(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

def check_hashes(password, hashed_text):
    if make_hashes(password) == hashed_text:
        return hashed_text
    return False

# --- Database Setup ---
def init_db():
    conn = sqlite3.connect('faculty_history.db')
    c = conn.cursor()
    
    # 1. Workload History Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS workload_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_date TEXT,
            week_label TEXT,
            teacher TEXT,
            class_lectures REAL,
            doubt_slots REAL,
            total_workload REAL,
            effective_capacity REAL,
            leaves REAL
        )
    ''')
    
    # 2. Users Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT,
            role TEXT
        )
    ''')
    
    # Create default admin account if table is empty
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
                  ("admin", make_hashes("matrix2026"), "admin"))
        
    conn.commit()
    conn.close()

init_db()

# --- Database User Management Functions ---
def login_user(username, password):
    conn = sqlite3.connect('faculty_history.db')
    c = conn.cursor()
    c.execute('SELECT * FROM users WHERE username =? AND password = ?', (username, make_hashes(password)))
    data = c.fetchall()
    conn.close()
    return data

def add_user(username, password, role="admin"):
    conn = sqlite3.connect('faculty_history.db')
    c = conn.cursor()
    try:
        c.execute('INSERT INTO users(username, password, role) VALUES (?,?,?)', (username, make_hashes(password), role))
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        conn.close()
        return False

def update_password(username, new_password):
    conn = sqlite3.connect('faculty_history.db')
    c = conn.cursor()
    c.execute('UPDATE users SET password = ? WHERE username = ?', (make_hashes(new_password), username))
    conn.commit()
    conn.close()

def save_to_db(df, week_label):
    conn = sqlite3.connect('faculty_history.db')
    c = conn.cursor()
    c.execute("DELETE FROM workload_history WHERE week_label = ?", (week_label,))
    
    df_to_save = df[['Teacher', 'Class Lectures', 'Doubt Slots', 'Total Workload', 'Effective Capacity', 'Leave Days Count']].copy()
    df_to_save['week_label'] = week_label
    df_to_save['upload_date'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    df_to_save.rename(columns={
        'Teacher': 'teacher',
        'Class Lectures': 'class_lectures',
        'Doubt Slots': 'doubt_slots',
        'Total Workload': 'total_workload',
        'Effective Capacity': 'effective_capacity',
        'Leave Days Count': 'leaves'
    }, inplace=True)
    
    df_to_save.to_sql('workload_history', conn, if_exists='append', index=False)
    conn.close()

# --- Page Setup ---
st.set_page_config(page_title="Matrix Net - Faculty Portal", layout="wide", page_icon="🏫")

# --- Custom CSS Styling ---
st.markdown("""
<style>
    div[data-testid="metric-container"] {
        background: linear-gradient(135deg, #1f2937 0%, #111827 100%);
        border: 1px solid #374151;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
    }
    div[data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 700 !important;
        color: #38bdf8 !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 14px !important;
        font-weight: 600 !important;
        color: #9ca3af !important;
    }
    .stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease-in-out !important;
    }
</style>
""", unsafe_allow_html=True)

# --- Authentication & Lockscreen ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

if not st.session_state["authenticated"]:
    # 1. Background Image Injection
    def set_login_background(image_file):
        if os.path.exists(image_file):
            with open(image_file, "rb") as f:
                encoded_string = base64.b64encode(f.read()).decode()
            
            css = f"""
            <style>
            .stApp {{
                background-image: url(data:image/jpeg;base64,{encoded_string});
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
                background-attachment: fixed;
            }}
            .stApp::before {{
                content: "";
                position: absolute;
                top: 0; left: 0; width: 100%; height: 100%;
                background-color: rgba(15, 23, 42, 0.6); 
                z-index: -1;
            }}
            [data-testid="column"]:nth-of-type(2) {{
                background: rgba(30, 41, 59, 0.7);
                backdrop-filter: blur(12px);
                -webkit-backdrop-filter: blur(12px);
                padding: 40px;
                border-radius: 16px;
                border: 1px solid rgba(255, 255, 255, 0.1);
                box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
            }}
            </style>
            """
            st.markdown(css, unsafe_allow_html=True)

    set_login_background("background.jpg")

    # 2. Login UI
    col1, col2, col3 = st.columns([1, 1.5, 1]) 
    with col2:
        if os.path.exists("logo.png"):
            st.markdown("""<div style="text-align: center;">""", unsafe_allow_html=True)
            st.image("logo.png", width=200)
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.markdown("<h2 style='text-align: center; color: #38bdf8;'>🟦 MATRIX NET</h2>", unsafe_allow_html=True)
            
        st.markdown("<h3 style='text-align: center;'>🔒 Secure Portal</h3>", unsafe_allow_html=True)
        st.write("")
        
        username = st.text_input("User ID")
        password = st.text_input("Password", type="password")
        st.write("")
        
        if st.button("Login", type="primary", use_container_width=True):
            result = login_user(username, password)
            if result:
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.rerun()
            else:
                st.error("Incorrect User ID or Password. Please try again.")
    st.stop()

# --- Sleek Sidebar Navigation ---
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.markdown("### 🟦 MATRIX NET")
        
    st.caption(f"Logged in as: **{st.session_state['username']}**")
    st.divider()
    
    selected = option_menu(
        menu_title="Main Menu",
        options=["Timetable Analyzer", "Schedule Calculator", "Doubt Generator", "Teacher Dashboard", "Historical Analytics", "User Management"],
        icons=["bar-chart-fill", "calendar3", "robot", "person-badge", "graph-up", "gear-fill"],
        menu_icon="cast",
        default_index=0,
        styles={"nav-link-selected": {"background-color": "#38bdf8", "color": "white"}}
    )
    
    st.divider()
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["username"] = ""
        st.rerun()

# ==========================================
# PAGE 1: TIMETABLE ANALYZER
# ==========================================
if selected == "Timetable Analyzer":
    st.header("📊 Faculty Workload Analyzer")
    
    with st.expander("📁 Upload / Update Timetables", expanded=True):
        mode = st.radio("Select Analysis Mode:", ["Class Timetable Only", "Class + Doubt Timetables"], horizontal=True)
        class_file = None
        doubt_file = None

        if mode == "Class Timetable Only":
            class_file = st.file_uploader("Upload Weekly Class Timetable (Excel)", type=["xlsx", "xls"], key="class_only")
        else:
            upload_col1, upload_col2 = st.columns(2)
            with upload_col1:
                class_file = st.file_uploader("1. Upload Class Timetable (Excel)", type=["xlsx", "xls"], key="class_file")
            with upload_col2:
                doubt_file = st.file_uploader("2. Upload Doubt Timetable (Excel)", type=["xlsx", "xls"], key="doubt_file")

    def parse_class_timetable(uploaded_file):
        df = pd.read_excel(uploaded_file, sheet_name=0, keep_default_na=False)
        days_map = {
            "Monday": range(1, 9), "Tuesday": range(9, 17), "Wednesday": range(17, 25),
            "Thursday": range(26, 34), "Friday": range(34, 42), "Saturday": range(42, 50)
        }
        records, teacher_stats = [], {}
        total_raw_assigned_slots, total_raw_teachers = 0, 0
        
        i = 2
        while i < len(df):
            teacher_name = str(df.iloc[i, 0]).strip()
            if teacher_name == '' or teacher_name.lower() == 'nan':
                i += 1; continue
                
            total_raw_teachers += 1
            subject_row = df.iloc[i, :]
            
            has_batch_row = False
            if i + 1 < len(df):
                next_teacher_col = str(df.iloc[i+1, 0]).strip()
                if next_teacher_col == '' or next_teacher_col.lower() == 'nan':
                    has_batch_row = True
                    batch_row = df.iloc[i+1, :]
                else:
                    batch_row = pd.Series([''] * len(subject_row))
            else:
                batch_row = pd.Series([''] * len(subject_row))
                
            valid_classes, total_classes, day_counts = {}, 0, {day: 0 for day in days_map}
            for day, cols in days_map.items():
                for col in cols:
                    if col >= len(subject_row): continue
                    subj = str(subject_row.iloc[col]).strip()
                    batch = str(batch_row.iloc[col]).strip()
                    if (subj != '' and subj.lower() != 'nan') and (batch != '' and batch.lower() != 'nan'):
                        total_raw_assigned_slots += 1
                        total_classes += 1
                        day_counts[day] += 1
                        if (subj, batch) not in valid_classes: valid_classes[(subj, batch)] = 0
                        valid_classes[(subj, batch)] += 1
            
            leave_days_list = [day for day, cnt in day_counts.items() if cnt == 0] if total_classes > 0 else []
            status = "Active" if total_classes > 0 else "Not Allotted"
            effective_capacity = 48 - (len(leave_days_list) * 8) if status == "Active" else 48
            leave_text = ", ".join(leave_days_list) if leave_days_list else ("None" if status == "Active" else "N/A")
                
            teacher_stats[teacher_name] = {
                'Total_Classes': total_classes, 'Effective_Capacity': effective_capacity,
                'Leave_Count': len(leave_days_list), 'Status': status, 'Leave_Days': leave_text
            }
            if not valid_classes: records.append({'Teacher': teacher_name, 'Subject': 'None', 'Batch': 'None', 'Lectures': 0})
            else:
                for (subj, batch), count in valid_classes.items():
                    records.append({'Teacher': teacher_name, 'Subject': subj, 'Batch': batch, 'Lectures': count})
            i += 2 if has_batch_row else 1
        return pd.DataFrame(records), teacher_stats, total_raw_teachers, total_raw_assigned_slots

    def parse_doubt_timetable(uploaded_file):
        df = pd.read_excel(uploaded_file, sheet_name=0, keep_default_na=False)
        doubt_counts, total_doubt_slots = {}, 0
        for index, row in df.iterrows():
            teacher_name = str(row.iloc[0]).strip()
            if teacher_name == 'nan' or not teacher_name or teacher_name == "Teacher's Name": continue
            slots_count = 0
            for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']:
                if day in df.columns:
                    slots_str = str(row[day]).strip()
                    if slots_str != 'nan' and slots_str:
                        slots = [s for s in [s.strip() for s in slots_str.split(',')] if s] 
                        slots_count += len(slots)
            doubt_counts[teacher_name] = slots_count
            total_doubt_slots += slots_count
        return doubt_counts, total_doubt_slots

    if class_file is not None:
        if mode == "Class + Doubt Timetables" and doubt_file is None:
            st.warning("Please upload the Doubt Timetable to proceed.")
        else:
            with st.spinner("Analyzing Timetables..."):
                df_tidy, class_teacher_stats, total_teachers, total_class_slots = parse_class_timetable(class_file)
                doubt_slots_map, total_doubt_slots = {}, 0
                if mode == "Class + Doubt Timetables" and doubt_file is not None:
                    doubt_slots_map, total_doubt_slots = parse_doubt_timetable(doubt_file)
                        
                summary_rows = []
                for teacher, info in class_teacher_stats.items():
                    class_lecs = info['Total_Classes']
                    doubt_lecs = doubt_slots_map.get(teacher, 0)
                    total_workload = class_lecs + doubt_lecs
                    effective_capacity = info['Effective_Capacity']
                    final_free_slots = max(0, effective_capacity - total_workload) 
                    class_util = class_lecs / effective_capacity if effective_capacity > 0 else 0
                    total_util = total_workload / effective_capacity if effective_capacity > 0 else 0
                        
                    summary_rows.append({
                        'Teacher': teacher, 'Class Lectures': class_lecs, 'Doubt Slots': doubt_lecs,
                        'Total Workload': total_workload, 'Leave / Off Days': info['Leave_Days'],
                        'Leave Days Count': info['Leave_Count'], 'Effective Capacity': effective_capacity,
                        'Net Free Slots': final_free_slots, 'True Class Util.': class_util,
                        'True Total Util.': total_util, 'Status': info['Status']
                    })
                    
                df_summary = pd.DataFrame(summary_rows).sort_values(by='Total Workload', ascending=False)
                st.session_state["df_summary_cached"] = df_summary
                
            st.markdown("##### 📈 Key Workload Overview")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Faculty", total_teachers)
            m2.metric("Class Lectures", total_class_slots)
            m3.metric("Doubt Slots", total_doubt_slots if mode == "Class + Doubt Timetables" else "N/A")
            m4.metric("Leaves Taken", df_summary['Leave Days Count'].sum())
            
            st.divider()
            st.markdown("##### 💾 Save to Historical Database")
            sc1, sc2 = st.columns([3, 1])
            with sc1:
                week_label = st.text_input("Timetable Label (e.g., 'Week of Aug 10')", placeholder="Week of Aug 10 - Aug 15")
            with sc2:
                st.write("")
                st.write("")
                if st.button("Save to Archive", use_container_width=True):
                    if week_label:
                        save_to_db(df_summary, week_label)
                        st.success(f"Data saved to database under '{week_label}'!")
                        st.toast("Saved to History!", icon="💾")
                    else:
                        st.error("Please provide a week label to save.")
                        
            st.divider()
            st.markdown("##### 📊 Workload Distribution (Top 15 Busiest)")
            chart_df = df_summary.head(15)
            fig = px.bar(chart_df, x='Teacher', y=['Class Lectures', 'Doubt Slots'], barmode='stack', color_discrete_sequence=['#38bdf8', '#8b5cf6'])
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc", margin=dict(t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("##### 📋 Faculty Workload Summary Table")
            display_df = df_summary.drop(columns=['Leave Days Count', 'Doubt Slots', 'True Total Util.', 'Total Workload']) if mode == "Class Timetable Only" else df_summary.drop(columns=['Leave Days Count'])
            st.dataframe(display_df.style.format({'True Class Util.': '{:.1%}', 'True Total Util.': '{:.1%}'}), use_container_width=True)

# ==========================================
# PAGE 2: SCHEDULE CALCULATOR 
# ==========================================
elif selected == "Schedule Calculator":
    st.header("📅 Syllabus Timeline Calculator")
    calc_option = st.selectbox("Calculate:", ["End Date", "Total Number of Lectures", "Lectures per Week", "Start Date"])
    col_a, col_b = st.columns(2)
    if calc_option == "End Date":
        with col_a:
            start_date = st.date_input("Start Date")
            total_lec = st.number_input("Total Number of Lectures", min_value=1, value=120)
        with col_b:
            lec_per_week = st.number_input("Lectures per Week", min_value=0.5, value=6.0, step=0.5)
            leave_days = st.number_input("Planned Holidays / Off Days", min_value=0, value=0, step=1)
        if st.button("Calculate", type="primary"):
            days = ((total_lec / lec_per_week) * 7) + leave_days
            st.success(f"### 🎯 Course Completion Date: **{(start_date + datetime.timedelta(days=days)).strftime('%B %d, %Y')}**")

    elif calc_option == "Total Number of Lectures":
        with col_a:
            start_date = st.date_input("Start Date")
            end_date = st.date_input("End Date", value=datetime.date.today() + datetime.timedelta(days=90))
        with col_b:
            lec_per_week = st.number_input("Lectures per Week", min_value=0.5, value=6.0, step=0.5)
            leave_days = st.number_input("Planned Holidays", min_value=0, value=0, step=1)
        if st.button("Calculate", type="primary"):
            effective_days = (end_date - start_date).days - leave_days
            if effective_days > 0: st.success(f"### 🎯 Total Possible Lectures: **{int((effective_days / 7) * lec_per_week)}**")

    elif calc_option == "Lectures per Week":
        with col_a:
            start_date = st.date_input("Start Date")
            end_date = st.date_input("End Date", value=datetime.date.today() + datetime.timedelta(days=90))
        with col_b:
            total_lec = st.number_input("Total Required Lectures", min_value=1, value=120)
            leave_days = st.number_input("Planned Holidays", min_value=0, value=0, step=1)
        if st.button("Calculate", type="primary"):
            effective_days = (end_date - start_date).days - leave_days
            if effective_days > 0: st.success(f"### 🎯 Required Weekly Pace: **{total_lec / (effective_days / 7):.1f}** lec/week")

    elif calc_option == "Start Date":
        with col_a:
            end_date = st.date_input("Target End Date")
            total_lec = st.number_input("Total Required Lectures", min_value=1, value=120)
        with col_b:
            lec_per_week = st.number_input("Lectures per Week", min_value=0.5, value=6.0, step=0.5)
            leave_days = st.number_input("Planned Holidays", min_value=0, value=0, step=1)
        if st.button("Calculate", type="primary"):
            days = ((total_lec / lec_per_week) * 7) + leave_days
            st.success(f"### 🎯 Required Course Start Date: **{(end_date - datetime.timedelta(days=days)).strftime('%B %d, %Y')}**")

# ==========================================
# PAGE 3: DOUBT GENERATOR
# ==========================================
elif selected == "Doubt Generator":
    st.header("🤖 Auto Doubt Generator")
    with st.expander("⚙️ Upload & Config", expanded=True):
        generator_class_file = st.file_uploader("Upload Class Timetable (Excel)", type=["xlsx", "xls"], key="gen_class")
        teacher_input = st.text_area("Enter Faculty Names (Comma Separated)", height=100)
    
    if st.button("✨ Generate Schedule", type="primary"):
        if generator_class_file is not None and teacher_input:
            with st.spinner("Generating conflict-free schedule..."):
                df_class = pd.read_excel(generator_class_file, sheet_name=0, keep_default_na=False)
                days_map = {"Monday": 1, "Tuesday": 9, "Wednesday": 17, "Thursday": 26, "Friday": 34, "Saturday": 42}
                
                def has_class(row, day_start_col, slot_idx):
                    col = day_start_col + slot_idx
                    if col >= len(row): return False
                    val = str(row.iloc[col]).strip()
                    return (val != '' and val.lower() != 'nan')

                target_teachers = [t.strip() for t in teacher_input.split(",") if t.strip()]
                generated_doubts = []

                for teacher in target_teachers:
                    teacher_idx = df_class[df_class.iloc[:, 0] == teacher].index
                    teacher_schedule = {"Teacher": teacher}
                    row = df_class.iloc[teacher_idx[0], :] if not teacher_idx.empty else pd.Series([''] * 50)

                    for day, start_col in days_map.items():
                        c0 = has_class(row, start_col, 0)
                        c1 = has_class(row, start_col, 1)
                        c2 = has_class(row, start_col, 2)
                        c3 = has_class(row, start_col, 3)
                        c4 = has_class(row, start_col, 4)
                        c5 = has_class(row, start_col, 5)
                        c6 = has_class(row, start_col, 6)
                        c7 = has_class(row, start_col, 7)
                        
                        class_count = sum([c0, c1, c2, c3, c4, c5, c6, c7])
                        d1_avail = not c2  
                        d2_avail = not (c4 and c5 and c6) 
                        d3_avail = not c6  
                        
                        doubts_needed = max(0, 5 - class_count)
                        if class_count == 0: doubts_needed = 3
                        
                        preference = []
                        if d2_avail: preference.append("D2")
                        if d3_avail: preference.append("D3")
                        if d1_avail: preference.append("D1")
                        
                        assigned_doubts = sorted(preference[:doubts_needed])
                        teacher_schedule[day] = ", ".join(assigned_doubts) if assigned_doubts else "None"
                    generated_doubts.append(teacher_schedule)
                    
                df_gen = pd.DataFrame(generated_doubts)
                st.session_state["df_doubt_gen"] = df_gen
            st.success("✨ Process Complete!")
            st.dataframe(df_gen, use_container_width=True)

# ==========================================
# PAGE 4: TEACHER DASHBOARD
# ==========================================
elif selected == "Teacher Dashboard":
    st.header("👤 Individual Faculty Dashboard")
    if "df_summary_cached" in st.session_state:
        df_sum = st.session_state["df_summary_cached"]
        teacher_list = df_sum["Teacher"].tolist()
        selected_teacher = st.selectbox("Select Faculty Member:", teacher_list)
        if selected_teacher:
            t_data = df_sum[df_sum["Teacher"] == selected_teacher].iloc[0]
            st.markdown(f"### 📌 Profile: **{selected_teacher}**")
            sc1, sc2, sc3, sc4 = st.columns(4)
            sc1.metric("Class Lectures", t_data["Class Lectures"])
            sc2.metric("Doubt Slots", t_data.get("Doubt Slots", 0))
            sc3.metric("Effective Capacity", f"{t_data['Effective Capacity']} Slots")
            sc4.metric("Free Slots Left", t_data["Net Free Slots"])
    else:
        st.info("💡 Please upload a timetable in the Analyzer menu first.")

# ==========================================
# PAGE 5: HISTORICAL ANALYTICS
# ==========================================
elif selected == "Historical Analytics":
    st.header("📈 Historical Workload Analytics")
    conn = sqlite3.connect('faculty_history.db')
    df_hist = pd.read_sql('SELECT * FROM workload_history', conn)
    conn.close()
    
    if df_hist.empty:
        st.info("No historical data found. Please analyze a timetable and click 'Save to Archive' first.")
    else:
        teacher_list = df_hist['teacher'].unique().tolist()
        selected_teacher_hist = st.selectbox("Select Faculty to View Trends:", teacher_list)
        df_teacher_hist = df_hist[df_hist['teacher'] == selected_teacher_hist]
        
        st.markdown(f"##### Workload Trend for {selected_teacher_hist}")
        fig = px.line(df_teacher_hist, x='week_label', y=['class_lectures', 'doubt_slots', 'total_workload'], markers=True)
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(df_teacher_hist[['week_label', 'class_lectures', 'doubt_slots', 'total_workload', 'leaves']], use_container_width=True)

# ==========================================
# PAGE 6: USER MANAGEMENT
# ==========================================
elif selected == "User Management":
    st.header("⚙️ User Management & Security")
    st.caption("Manage portal accounts, add new administrator users, or update your password.")
    
    tab_user1, tab_user2 = st.tabs(["🔑 Change Password", "➕ Add New User"])
    
    with tab_user1:
        st.markdown(f"##### Change Password for **{st.session_state['username']}**")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            with st.form("change_password_form"):
                current_pass = st.text_input("Current Password", type="password")
                new_pass = st.text_input("New Password", type="password")
                confirm_pass = st.text_input("Confirm New Password", type="password")
                
                submitted = st.form_submit_button("Update Password", type="primary")
                
                if submitted:
                    check_user = login_user(st.session_state['username'], current_pass)
                    if not check_user:
                        st.error("Current password is incorrect.")
                    elif new_pass.strip() == "":
                        st.error("New password cannot be empty.")
                    elif new_pass != confirm_pass:
                        st.error("New passwords do not match.")
                    else:
                        update_password(st.session_state['username'], new_pass)
                        st.success("Password updated successfully!")

    with tab_user2:
        st.markdown("##### Create a New Portal Account")
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            with st.form("add_new_user_form"):
                new_username = st.text_input("New User ID / Username")
                new_user_pass = st.text_input("New Account Password", type="password")
                confirm_user_pass = st.text_input("Confirm Account Password", type="password")
                
                submitted_new = st.form_submit_button("Create Account", type="primary")
                
                if submitted_new:
                    if new_username.strip() == "":
                        st.error("Username cannot be empty.")
                    elif new_user_pass.strip() == "":
                        st.error("Password cannot be empty.")
                    elif new_user_pass != confirm_user_pass:
                        st.error("Passwords do not match.")
                    else:
                        success = add_user(new_username, new_user_pass)
                        if success:
                            st.success(f"Account for '{new_username}' created successfully!")
                        else:
                            st.error(f"Username '{new_username}' already exists. Choose a different username.")