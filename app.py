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
# TAB 1: TIMETABLE ANALYZER (Kept Intact)
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

    if class_file is not None:
        if mode == "Class + Doubt Timetables" and doubt_file is None:
            st.warning("Please upload the Doubt Timetable to proceed.")
        else:
            df_tidy, class_teacher_stats, total_teachers, total_class_slots = parse_class_timetable(class_file)
            st.success("Successfully analyzed!")

# ==========================================
# TAB 2: SCHEDULE CALCULATOR (Kept Intact)
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
            leave_days = st.number_input("Holidays", min_value=0, value=0)
        if st.button("Calculate End Date"):
            days = ((total_lec / lec_per_week) * 7) + leave_days
            st.success(f"### 🎯 Course ends on: **{(start_date + datetime.timedelta(days=days)).strftime('%B %d, %Y')}**")

# ==========================================
# TAB 3: DOUBT GENERATOR (NEW FEATURE)
# ==========================================
with tab3:
    st.markdown("### 🤖 Auto-Generate Doubt Timetable")
    st.info("Upload your **Class Timetable** and a list of **Teachers** to auto-assign non-conflicting D1, D2, and D3 slots based on your constraints.")
    
    # 1. Upload Class Timetable for reference
    generator_class_file = st.file_uploader("1. Upload Current Class Timetable (to check overlaps)", type=["xlsx", "xls"], key="gen_class_file")
    
    # 2. Input Teachers
    teacher_input = st.text_area("2. Enter Teacher Names (comma separated, e.g., MH, SD, AA)", height=100)
    
    if st.button("Generate Doubt Schedule", type="primary"):
        if generator_class_file is None:
            st.error("Please upload the class timetable to check constraints.")
        elif not teacher_input:
            st.error("Please enter at least one teacher name.")
        else:
            # Read Class Timetable
            df_class = pd.read_excel(generator_class_file, sheet_name=0, keep_default_na=False)
            days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
            
            # Map slots to indices in the excel sheet (Assuming standard M1-M4, E1-E4 order)
            # D1 = 9:30 - 11:00. Overlaps M2 (idx 1). Just After M1 (idx 0).
            # D2 = 12:30 - 2:00. Overlaps M4 (idx 3). Just After M3 (idx 2).
            # D3 = 5:30 - 7:00. Overlaps E3/E4 (idx 6, 7). Just After E2 (idx 5).
            
            def has_class(row, day_offset, slot_idx):
                col = day_offset + slot_idx
                return (str(row.iloc[col]).strip() != '' and str(row.iloc[col]).strip().lower() != 'nan')

            target_teachers = [t.strip() for t in teacher_input.split(",") if t.strip()]
            generated_doubts = []

            for teacher in target_teachers:
                # Find teacher row in class timetable
                teacher_idx = df_class[df_class.iloc[:, 0] == teacher].index
                
                teacher_schedule = {"Teacher": teacher}
                
                if not teacher_idx.empty:
                    row = df_class.iloc[teacher_idx[0], 1:] # Drop first col
                else:
                    # Teacher has no classes, completely free for doubts
                    row = pd.Series([''] * 48)

                for day_idx, day in enumerate(days):
                    offset = day_idx * 8
                    assigned_doubts = []
                    
                    # Check D1 Constraint
                    if not has_class(row, offset, 1): # Not overlapping M2
                        if not has_class(row, offset, 0): # Not just after M1
                            assigned_doubts.append("D1")
                            
                    # Check D2 Constraint
                    if not has_class(row, offset, 3): # Not overlapping M4
                        if not has_class(row, offset, 2): # Not just after M3
                            # Continous 4 check: if M1, M2 both exist, and we add D2... 
                            if not (has_class(row, offset, 0) and has_class(row, offset, 1)):
                                assigned_doubts.append("D2")
                                
                    # Check D3 Constraint
                    if not has_class(row, offset, 6) and not has_class(row, offset, 7): # Not overlapping E3/E4
                        if not has_class(row, offset, 5): # Not just after E2
                            assigned_doubts.append("D3")
                            
                    teacher_schedule[day] = ", ".join(assigned_doubts) if assigned_doubts else "None"
                
                generated_doubts.append(teacher_schedule)
                
            st.success("Doubt Schedule Generated Successfully!")
            df_gen = pd.DataFrame(generated_doubts)
            st.dataframe(df_gen, use_container_width=True)

            # Convert to CSV for quick download
            csv = df_gen.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download Generated Doubt Schedule (CSV)", csv, "Generated_Doubt_Schedule.csv", "text/csv")