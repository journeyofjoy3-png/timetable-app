import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils.dataframe import dataframe_to_rows

# --- Page Setup ---
st.set_page_config(page_title="Faculty Timetable Portal", layout="wide")

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

# --- Main App Interface ---
col1, col2 = st.columns([8, 1])
with col1:
    st.title("📊 Faculty Timetable & Doubt Analyzer")
with col2:
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

st.markdown("Upload your weekly NEET Class Timetable and optional Doubt Timetable to analyze workload, exact leave days, and true utilization.")

# --- Select Mode ---
mode = st.radio(
    "Select Analysis Mode:", 
    ["Class Timetable Only", "Class + Doubt Timetables"], 
    horizontal=True
)

st.divider()

# --- File Uploaders ---
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

# --- Helper Function: Parse Timetable Sheet ---
def parse_timetable(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0, keep_default_na=False)
    
    days_map = {
        "Monday": range(1, 9),
        "Tuesday": range(9, 17),
        "Wednesday": range(17, 25),
        "Thursday": range(25, 33),
        "Friday": range(33, 41),
        "Saturday": range(41, 49)
    }
    
    records = []
    teacher_stats = {}
    total_raw_assigned_slots = 0
    total_raw_teachers = 0
    
    i = 2
    while i < len(df):
        teacher_name = str(df.iloc[i, 0]).strip()
        if teacher_name == '' or teacher_name.lower() == 'nan':
            i += 1
            continue
            
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
            
        valid_classes = {}
        total_classes = 0
        day_counts = {day: 0 for day in days_map}
        
        for day, cols in days_map.items():
            for col in cols:
                if col >= len(subject_row):
                    continue
                subj_str = str(subject_row.iloc[col]).strip()
                batch_str = str(batch_row.iloc[col]).strip()
                
                is_subj_empty = (subj_str == '' or subj_str.lower() == 'nan')
                is_batch_empty = (batch_str == '' or batch_str.lower() == 'nan')
                
                if not is_subj_empty and not is_batch_empty:
                    total_raw_assigned_slots += 1
                    total_classes += 1
                    day_counts[day] += 1
                    if (subj_str, batch_str) not in valid_classes:
                        valid_classes[(subj_str, batch_str)] = 0
                    valid_classes[(subj_str, batch_str)] += 1
        
        # Correct Leave Logic & Free Slot Math
        leave_days_list = []
        if total_classes == 0:
            status = "Not Allotted"
            leave_text = "N/A (All Slots Free)"
            effective_capacity = 48
        else:
            status = "Active"
            leave_days_list = [day for day, cnt in day_counts.items() if cnt == 0]
            leave_text = ", ".join(leave_days_list) if leave_days_list else "None"
            
            # Deduct 8 slots from total weekly capacity for every leave day
            effective_capacity = 48 - (len(leave_days_list) * 8)
            
        net_free_slots = effective_capacity - total_classes
        
        teacher_stats[teacher_name] = {
            'Total_Classes': total_classes,
            'Effective_Capacity': effective_capacity,
            'Net_Free_Slots': max(0, net_free_slots),
            'Leave_Count': len(leave_days_list),
            'Status': status,
            'Leave_Days': leave_text
        }
                
        if not valid_classes:
            records.append({'Teacher': teacher_name, 'Subject': 'None', 'Batch': 'None', 'Lectures': 0})
        else:
            for (subj, batch), count in valid_classes.items():
                records.append({
                    'Teacher': teacher_name, 'Subject': subj, 'Batch': batch, 'Lectures': count
                })
        
        if has_batch_row:
            i += 2
        else:
            i += 1

    return pd.DataFrame(records), teacher_stats, total_raw_teachers, total_raw_assigned_slots


# --- Main Processing Logic ---
if class_file is not None:
    if mode == "Class + Doubt Timetables" and doubt_file is None:
        st.warning("Please upload the Doubt Timetable to proceed, or switch to 'Class Timetable Only' mode.")
    else:
        st.success("Timetable(s) uploaded successfully! Processing data...")
        
        df_tidy, class_teacher_stats, total_teachers, total_class_slots = parse_timetable(class_file)
        
        # Process Optional Doubt File
        doubt_slots_map = {}
        total_doubt_slots = 0
        if doubt_file is not None:
            _, doubt_stats, _, total_doubt_slots = parse_timetable(doubt_file)
            for teacher, info in doubt_stats.items():
                doubt_slots_map[teacher] = info['Total_Classes']
                
        # Build Combined Summary Table
        summary_rows = []
        for teacher, info in class_teacher_stats.items():
            class_lecs = info['Total_Classes']
            doubt_lecs = doubt_slots_map.get(teacher, 0)
            total_workload = class_lecs + doubt_lecs
            
            effective_capacity = info['Effective_Capacity']
            
            final_free_slots = effective_capacity - total_workload
            final_free_slots = max(0, final_free_slots) 
            
            if effective_capacity > 0:
                class_util = class_lecs / effective_capacity
                total_util = total_workload / effective_capacity
            else:
                class_util = 0
                total_util = 0
                
            summary_rows.append({
                'Teacher': teacher,
                'Class Lectures': class_lecs,
                'Doubt Slots': doubt_lecs,
                'Total Workload': total_workload,
                'Leave / Off Days': info['Leave_Days'],
                'Leave Days Count': info['Leave_Count'],
                'Effective Capacity': effective_capacity,
                'Net Free Slots': final_free_slots,
                'True Class Util.': class_util,
                'True Total Util.': total_util,
                'Status': info['Status']
            })
            
        df_summary = pd.DataFrame(summary_rows)
        df_summary = df_summary.sort_values(by='Total Workload', ascending=False)
        
        # --- UI Dashboard Metrics ---
        st.subheader("📊 Key Workload Metrics")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Faculty", total_teachers)
        m2.metric("Total Class Lectures", total_class_slots)
        
        if mode == "Class + Doubt Timetables":
            m3.metric("Total Doubt Slots", total_doubt_slots)
        else:
            m3.metric("Total Doubt Slots", "N/A")
            
        total_leave_days_taken = df_summary['Leave Days Count'].sum()
        m4.metric("Total Leave Days Taken", total_leave_days_taken)
        
        st.divider()
        
        st.subheader("📋 Faculty Workload & True Utilization Summary")
        
        # Hide Doubt columns if in Class Only mode for a cleaner UI
        if mode == "Class Timetable Only":
            display_df = df_summary.drop(columns=['Leave Days Count', 'Doubt Slots', 'True Total Util.', 'Total Workload'])
            st.dataframe(display_df.style.format({
                'True Class Util.': '{:.1%}'
            }), use_container_width=True)
        else:
            display_df = df_summary.drop(columns=['Leave Days Count'])
            st.dataframe(display_df.style.format({
                'True Class Util.': '{:.1%}',
                'True Total Util.': '{:.1%}'
            }), use_container_width=True)
        
        # --- Generate Downloadable Excel ---
        def generate_excel(summary_df, detail_df, app_mode):
            wb = Workbook()
            
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill("solid", fgColor="2F4F4F")
            border_side = Side(border_style="thin", color="D3D3D3")
            border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
            center_align = Alignment(horizontal="center", vertical="center")

            def apply_styling(ws):
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_align
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.border = border
                        if isinstance(cell.value, (int, float)):
                            cell.alignment = center_align
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    ws.column_dimensions[column].width = max_length + 4

            # Summary Tab
            ws_summary = wb.active
            ws_summary.title = "Workload & True Utilization"
            
            if app_mode == "Class Timetable Only":
                df_export = summary_df.drop(columns=['Leave Days Count', 'Doubt Slots', 'True Total Util.', 'Total Workload'])
                for r in dataframe_to_rows(df_export, index=False, header=True):
                    ws_summary.append(r)
                apply_styling(ws_summary)
                # Format Percentage Column (F for Class Only)
                for cell in ws_summary['F'][1:]:
                    cell.number_format = '0.0%'
                rule = ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B')
                ws_summary.conditional_formatting.add(f'F2:F{ws_summary.max_row}', rule)
            else:
                df_export = summary_df.drop(columns=['Leave Days Count'])
                for r in dataframe_to_rows(df_export, index=False, header=True):
                    ws_summary.append(r)
                apply_styling(ws_summary)
                # Format Percentage Columns (I & J for Class + Doubt)
                for col_letter in ['I', 'J']:
                    for cell in ws_summary[col_letter][1:]:
                        cell.number_format = '0.0%'
                rule = ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B')
                ws_summary.conditional_formatting.add(f'J2:J{ws_summary.max_row}', rule)

            # Details Tab
            ws_detail = wb.create_sheet(title="Detailed Class Mapping")
            for r in dataframe_to_rows(detail_df, index=False, header=True):
                ws_detail.append(r)
            apply_styling(ws_detail)
            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            return output

        excel_file = generate_excel(df_summary, df_tidy, mode)
        
        st.download_button(
            label="📥 Download Complete Excel Analysis",
            data=excel_file,
            file_name="Faculty_Timetable_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )