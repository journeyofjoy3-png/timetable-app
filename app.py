import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils.dataframe import dataframe_to_rows

# --- Page Setup ---
st.set_page_config(page_title="Faculty Operations Portal", layout="wide")

# --- Authentication ---
USER_ID = "admin"
PASSWORD = "matrix2026"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 Secure Portal Login")
    username = st.text_input("User ID")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username == USER_ID and password == PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect User ID or Password. Please try again.")
    st.stop()

# --- Main App Interface & Logout ---
col1, col2 = st.columns([8, 1])
with col1:
    st.title("🏫 Faculty Operations Portal")
with col2:
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# --- Create Navigation Tabs ---
tab1, tab2, tab3 = st.tabs(["📊 Timetable Analyzer", "📅 Schedule Calculator", "🤖 Doubt Generator"])

# ==========================================
# TAB 1: TIMETABLE ANALYZER
# ==========================================
with tab1:
    st.markdown("Upload your weekly Class Timetable and optional Doubt Timetable to analyze workload and utilization.")
    mode = st.radio("Select Analysis Mode:", ["Class Timetable Only", "Class + Doubt Timetables"], horizontal=True)
    st.divider()

    class_file = None
    doubt_file = None

    if mode == "Class Timetable Only":
        class_file = st.file_uploader("Upload Class Timetable", type=["xlsx", "xls"], key="class_only")
    else:
        upload_col1, upload_col2 = st.columns(2)
        with upload_col1:
            class_file = st.file_uploader("1. Upload Class Timetable", type=["xlsx", "xls"], key="class_file")
        with upload_col2:
            doubt_file = st.file_uploader("2. Upload Doubt Timetable", type=["xlsx", "xls"], key="doubt_file")

    def parse_class_timetable(uploaded_file):
        df = pd.read_excel(uploaded_file, sheet_name=0, keep_default_na=False)
        days_map = {
            "Monday": range(1, 9), "Tuesday": range(9, 17), "Wednesday": range(17, 25),
            "Thursday": range(25, 33), "Friday": range(33, 41), "Saturday": range(41, 49)
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
                    subj_str = str(subject_row.iloc[col]).strip()
                    batch_str = str(batch_row.iloc[col]).strip()
                    is_subj = (subj_str != '' and subj_str.lower() != 'nan')
                    is_batch = (batch_str != '' and batch_str.lower() != 'nan')
                    
                    if is_subj and is_batch:
                        total_raw_assigned_slots += 1
                        total_classes += 1
                        day_counts[day] += 1
                        if (subj_str, batch_str) not in valid_classes: valid_classes[(subj_str, batch_str)] = 0
                        valid_classes[(subj_str, batch_str)] += 1
            
            leave_days_list = [day for day, cnt in day_counts.items() if cnt == 0] if total_classes > 0 else []
            status = "Active" if total_classes > 0 else "Not Allotted"
            leave_text = ", ".join(leave_days_list) if leave_days_list else ("None" if status == "Active" else "N/A (All Slots Free)")
            effective_capacity = 48 - (len(leave_days_list) * 8) if status == "Active" else 48
                
            teacher_stats[teacher_name] = {
                'Total_Classes': total_classes, 'Effective_Capacity': effective_capacity,
                'Leave_Count': len(leave_days_list), 'Status': status, 'Leave_Days': leave_text
            }
                    
            if not valid_classes:
                records.append({'Teacher': teacher_name, 'Subject': 'None', 'Batch': 'None', 'Lectures': 0})
            else:
                for (subj, batch), count in valid_classes.items():
                    records.append({'Teacher': teacher_name, 'Subject': subj, 'Batch': batch, 'Lectures': count})
            
            i += 2 if has_batch_row else 1

        return pd.DataFrame(records), teacher_stats, total_raw_teachers, total_raw_assigned_slots

    def parse_doubt_timetable(uploaded_file):
        df = pd.read_excel(uploaded_file, sheet_name=0, keep_default_na=False)
        doubt_counts, total_doubt_slots = {}, 0
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        
        for index, row in df.iterrows():
            if "Teacher's Name" in df.columns: teacher_name = str(row["Teacher's Name"]).strip()
            else: teacher_name = str(row.iloc[0]).strip()
            if teacher_name == 'nan' or not teacher_name or teacher_name == "Teacher's Name": continue
                
            slots_count = 0
            for day in days:
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
            st.warning("Please upload the Doubt Timetable to proceed, or switch to 'Class Timetable Only' mode.")
        else:
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
            
            st.subheader("📊 Key Workload Metrics")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Faculty", total_teachers)
            m2.metric("Total Class Lectures", total_class_slots)
            m3.metric("Total Doubt Slots", total_doubt_slots if mode == "Class + Doubt Timetables" else "N/A")
            m4.metric("Total Leave Days Taken", df_summary['Leave Days Count'].sum())
            
            st.divider()
            st.subheader("📋 Faculty Workload & True Utilization Summary")
            
            if mode == "Class Timetable Only":
                display_df = df_summary.drop(columns=['Leave Days Count', 'Doubt Slots', 'True Total Util.', 'Total Workload'])
                st.dataframe(display_df.style.format({'True Class Util.': '{:.1%}'}), use_container_width=True)
            else:
                display_df = df_summary.drop(columns=['Leave Days Count'])
                st.dataframe(display_df.style.format({'True Class Util.': '{:.1%}', 'True Total Util.': '{:.1%}'}), use_container_width=True)
            
            def generate_excel(summary_df, detail_df, app_mode):
                wb = Workbook()
                header_font, header_fill, center_align = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="2F4F4F"), Alignment(horizontal="center", vertical="center")
                border = Border(left=Side(border_style="thin", color="D3D3D3"), right=Side(border_style="thin", color="D3D3D3"), top=Side(border_style="thin", color="D3D3D3"), bottom=Side(border_style="thin", color="D3D3D3"))

                def apply_styling(ws):
                    for cell in ws[1]: cell.font, cell.fill, cell.alignment = header_font, header_fill, center_align
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                        for cell in row:
                            cell.border = border
                            if isinstance(cell.value, (int, float)): cell.alignment = center_align
                    ws.freeze_panes = "A2"
                    ws.auto_filter.ref = ws.dimensions
                    for col in ws.columns:
                        max_len = max([len(str(cell.value)) for cell in col if cell.value is not None] + [0])
                        ws.column_dimensions[col[0].column_letter].width = max_len + 4

                ws_summary = wb.active
                ws_summary.title = "Workload & True Utilization"
                
                cols_to_drop = ['Leave Days Count', 'Doubt Slots', 'True Total Util.', 'Total Workload'] if app_mode == "Class Timetable Only" else ['Leave Days Count']
                df_export = summary_df.drop(columns=cols_to_drop)
                for r in dataframe_to_rows(df_export, index=False, header=True): ws_summary.append(r)
                apply_styling(ws_summary)
                
                format_cols = ['F'] if app_mode == "Class Timetable Only" else ['I', 'J']
                for col_letter in format_cols:
                    for cell in ws_summary[col_letter][1:]: cell.number_format = '0.0%'
                
                rule = ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B')
                ws_summary.conditional_formatting.add(f"{format_cols[-1]}2:{format_cols[-1]}{ws_summary.max_row}", rule)

                ws_detail = wb.create_sheet(title="Detailed Class Mapping")
                for r in dataframe_to_rows(detail_df, index=False, header=True): ws_detail.append(r)
                apply_styling(ws_detail)
                
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                return output

            excel_file = generate_excel(df_summary, df_tidy, mode)
            st.download_button(label="📥 Download Complete Excel Analysis", data=excel_file, file_name="Faculty_Timetable_Analysis.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# TAB 2: SCHEDULE CALCULATOR 
# ==========================================
with tab2:
    st.markdown("Calculate syllabus timelines by providing any 3 variables.")
    calc_option = st.selectbox("What do you want to calculate?", ["End Date", "Total Number of Lectures", "Lectures per Week", "Start Date"])
    
    col_a, col_b = st.columns(2)
    if calc_option == "End Date":
        with col_a:
            start_date = st.date_input("Start Date")
            total_lec = st.number_input("Total Number of Lectures", min_value=1, value=120)
        with col_b:
            lec_per_week = st.number_input("Lectures per Week", min_value=0.5, value=6.0)
            leave_days = st.number_input("Holidays / Leave Days", min_value=0, value=0)
        if st.button("Calculate End Date"):
            days = ((total_lec / lec_per_week) * 7) + leave_days
            st.success(f"### 🎯 Course ends on: **{(start_date + datetime.timedelta(days=days)).strftime('%B %d, %Y')}**")

    elif calc_option == "Total Number of Lectures":
        with col_a:
            start_date = st.date_input("Start Date")
            end_date = st.date_input("End Date", value=datetime.date.today() + datetime.timedelta(days=90))
        with col_b:
            lec_per_week = st.number_input("Lectures per Week", min_value=0.5, value=6.0, step=0.5)
            leave_days = st.number_input("Holidays / Leave Days", min_value=0, value=0, step=1)
        if st.button("Calculate Total Lectures"):
            effective_days = (end_date - start_date).days - leave_days
            if effective_days > 0:
                st.success(f"### 🎯 Total Lectures possible: **{int((effective_days / 7) * lec_per_week)}** lectures")
            else:
                st.error("Invalid dates or too many leave days.")

    elif calc_option == "Lectures per Week":
        with col_a:
            start_date = st.date_input("Start Date")
            end_date = st.date_input("End Date", value=datetime.date.today() + datetime.timedelta(days=90))
        with col_b:
            total_lec = st.number_input("Total Number of Lectures required", min_value=1, value=120)
            leave_days = st.number_input("Holidays / Leave Days", min_value=0, value=0, step=1)
        if st.button("Calculate Lectures per Week"):
            effective_days = (end_date - start_date).days - leave_days
            if effective_days > 0:
                st.success(f"### 🎯 You need to schedule: **{total_lec / (effective_days / 7):.1f}** lectures per week")
            else:
                st.error("Invalid dates or too many leave days.")

    elif calc_option == "Start Date":
        with col_a:
            end_date = st.date_input("Target End Date")
            total_lec = st.number_input("Total Number of Lectures", min_value=1, value=120)
        with col_b:
            lec_per_week = st.number_input("Lectures per Week", min_value=0.5, value=6.0, step=0.5)
            leave_days = st.number_input("Holidays / Leave Days", min_value=0, value=0, step=1)
        if st.button("Calculate Start Date"):
            days = ((total_lec / lec_per_week) * 7) + leave_days
            st.success(f"### 🎯 The course must start on: **{(end_date - datetime.timedelta(days=days)).strftime('%B %d, %Y')}**")


# ==========================================
# TAB 3: DOUBT GENERATOR (NEW FEATURE)
# ==========================================
with tab3:
    st.markdown("### 🤖 Auto-Generate Doubt Timetable")
    st.info("Upload your **Class Timetable** and a list of **Teachers** to auto-assign non-conflicting D1, D2, and D3 slots based on exact class timings.")
    
    generator_class_file = st.file_uploader("1. Upload Current Class Timetable (to check overlaps)", type=["xlsx", "xls"], key="gen_class_file")
    teacher_input = st.text_area("2. Enter Teacher Names (comma separated, e.g., MH, SD, AA)", height=100)
    
    if st.button("Generate Doubt Schedule", type="primary"):
        if generator_class_file is None:
            st.error("Please upload the class timetable to check constraints.")
        elif not teacher_input:
            st.error("Please enter at least one teacher name.")
        else:
            df_class = pd.read_excel(generator_class_file, sheet_name=0, keep_default_na=False)
            days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
            
            # Helper to check if a class exists at a specific slot index
            def has_class(row, day_offset, slot_idx):
                col = day_offset + slot_idx
                return (str(row.iloc[col]).strip() != '' and str(row.iloc[col]).strip().lower() != 'nan')

            target_teachers = [t.strip() for t in teacher_input.split(",") if t.strip()]
            generated_doubts = []

            for teacher in target_teachers:
                teacher_idx = df_class[df_class.iloc[:, 0] == teacher].index
                teacher_schedule = {"Teacher": teacher}
                
                # Fetch row or create empty row if teacher has no classes
                if not teacher_idx.empty:
                    row = df_class.iloc[teacher_idx[0], 1:] 
                else:
                    row = pd.Series([''] * 48)

                for day_idx, day in enumerate(days):
                    offset = day_idx * 8
                    
                    # Read all 8 slots for the day
                    c0 = has_class(row, offset, 0) # M1 (06:30 - 08:00)
                    c1 = has_class(row, offset, 1) # M2 (08:10 - 09:40)
                    c2 = has_class(row, offset, 2) # M3 (09:50 - 11:20)
                    c3 = has_class(row, offset, 3) # M4 (11:30 - 13:00)
                    c4 = has_class(row, offset, 4) # E1 (14:00 - 15:30)
                    c5 = has_class(row, offset, 5) # E2 (15:40 - 17:10)
                    c6 = has_class(row, offset, 6) # E3 (17:20 - 18:50)
                    c7 = has_class(row, offset, 7) # E4 (19:00 - 20:30)
                    
                    assigned_doubts = []
                    
                    # D1 (09:30 - 11:00) Logic
                    # Overlaps M2 (ends 09:40) and M3 (starts 09:50)
                    if not c1 and not c2:
                        assigned_doubts.append("D1")
                            
                    # D2 (12:30 - 14:00) Logic
                    # Overlaps M4 (11:30 - 13:00)
                    if not c3:
                        # Prevent 4 continuous blocks: If D2 is added, and E1, E2, E3 exist 
                        # That creates a continuous block from 12:30 to 18:50
                        if not (c4 and c5 and c6):
                            assigned_doubts.append("D2")
                                
                    # D3 (17:30 - 19:00) Logic
                    # Overlaps E3 (17:20 - 18:50)
                    if not c6:
                        # "Just after class" rule: E2 ends at 17:10. D3 starts at 17:30.
                        # This 20-min gap triggers the ban. E2 MUST be free.
                        if not c5:
                            assigned_doubts.append("D3")
                            
                    teacher_schedule[day] = ", ".join(assigned_doubts) if assigned_doubts else "None"
                
                generated_doubts.append(teacher_schedule)
                
            st.success("Doubt Schedule Generated Successfully!")
            df_gen = pd.DataFrame(generated_doubts)
            st.dataframe(df_gen, use_container_width=True)

            csv = df_gen.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download Generated Doubt Schedule (CSV)", csv, "Generated_Doubt_Schedule.csv", "text/csv")