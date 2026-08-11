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
# Change these to your preferred login details!
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
    
    # This stops the rest of the app from loading if not logged in
    st.stop()

# --- Main App (Only visible after login) ---
col1, col2 = st.columns([8, 1])
with col1:
    st.title("📊 Faculty Timetable Analyzer")
with col2:
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

st.markdown("Upload your weekly NEET faculty Excel timetable to instantly generate a structured workload analysis and dashboard.")

# --- File Uploader ---
uploaded_file = st.file_uploader("Upload Timetable (Excel format)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success("File uploaded successfully! Processing data...")
    
    # 1. Read Data 
    df = pd.read_excel(uploaded_file, sheet_name=0, keep_default_na=False)
    
    records = []
    total_raw_assigned_slots = 0
    total_raw_teachers = 0
    
    i = 2
    while i < len(df):
        teacher_name = str(df.iloc[i, 0]).strip()
        if teacher_name == '' or teacher_name.lower() == 'nan':
            i += 1
            continue
            
        total_raw_teachers += 1
        subject_row = df.iloc[i, 1:]
        
        has_batch_row = False
        if i + 1 < len(df):
            next_teacher_col = str(df.iloc[i+1, 0]).strip()
            if next_teacher_col == '' or next_teacher_col.lower() == 'nan':
                has_batch_row = True
                batch_row = df.iloc[i+1, 1:]
            else:
                batch_row = pd.Series([''] * len(subject_row))
        else:
            batch_row = pd.Series([''] * len(subject_row))
            
        valid_classes = {}
        free_slots = 0
        
        for subj, batch in zip(subject_row, batch_row):
            subj_str = str(subj).strip()
            batch_str = str(batch).strip()
            
            is_subj_empty = (subj_str == '' or subj_str.lower() == 'nan')
            is_batch_empty = (batch_str == '' or batch_str.lower() == 'nan')
            
            if is_subj_empty and is_batch_empty:
                free_slots += 1
            elif not is_subj_empty and not is_batch_empty:
                total_raw_assigned_slots += 1
                if (subj_str, batch_str) not in valid_classes:
                    valid_classes[(subj_str, batch_str)] = 0
                valid_classes[(subj_str, batch_str)] += 1
                
        if not valid_classes:
            records.append({'Teacher': teacher_name, 'Subject': 'None', 'Batch': 'None', 'Lectures': 0, 'Free_Slots': free_slots})
        else:
            for (subj, batch), count in valid_classes.items():
                records.append({
                    'Teacher': teacher_name, 'Subject': subj, 'Batch': batch, 'Lectures': count, 'Free_Slots': free_slots
                })
        
        if has_batch_row:
            i += 2
        else:
            i += 1

    df_tidy = pd.DataFrame(records)
    
    # Create Summary
    df_summary = df_tidy.groupby('Teacher').agg(
        Total_Lectures=('Lectures', 'sum'),
        Free_Slots=('Free_Slots', 'first')
    ).reset_index()
    df_summary['Total_Capacity'] = 48
    df_summary['Utilization'] = df_summary['Total_Lectures'] / df_summary['Total_Capacity']
    df_summary = df_summary.sort_values(by='Total_Lectures', ascending=False)
    
    # --- Display Metrics on Web Page ---
    st.subheader("Data Verification")
    metric_col1, metric_col2, metric_col3 = st.columns(3)
    metric_col1.metric("Total Faculty", total_raw_teachers)
    metric_col2.metric("Total Active Lectures", total_raw_assigned_slots)
    metric_col3.metric("Structured Lectures", df_tidy['Lectures'].sum())
    
    st.divider()
    
    st.subheader("Faculty Workload Preview")
    st.dataframe(df_summary.head(10), use_container_width=True)
    
    # --- Generate Excel for Download ---
    def generate_excel(summary_df, detail_df):
        wb = Workbook()
        
        # Styles
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
        ws_summary.title = "Workload Summary"
        for r in dataframe_to_rows(summary_df, index=False, header=True):
            ws_summary.append(r)
        apply_styling(ws_summary)
        for cell in ws_summary['E'][1:]:
            cell.number_format = '0.0%'
        rule = ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B')
        ws_summary.conditional_formatting.add(f'E2:E{ws_summary.max_row}', rule)

        # Details Tab
        ws_detail = wb.create_sheet(title="Detailed Analysis")
        for r in dataframe_to_rows(detail_df, index=False, header=True):
            ws_detail.append(r)
        apply_styling(ws_detail)
        
        # Dashboard Tab
        ws_dash = wb.create_sheet(title="Dashboard", index=0) 
        ws_dash.sheet_view.showGridLines = False
        ws_dash['B2'] = "Faculty Timetable Analysis Dashboard"
        ws_dash['B2'].font = Font(size=18, bold=True, color="2F4F4F")
        ws_dash['B4'] = "Total Faculty Analyzed:"
        ws_dash['C4'] = len(summary_df)
        ws_dash['C4'].font = Font(bold=True)
        ws_dash['B5'] = "Total Active Assignments:"
        ws_dash['C5'] = len(detail_df[detail_df['Lectures'] > 0])
        ws_dash['C5'].font = Font(bold=True)

        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Top 15 Faculty by Total Weekly Lectures"
        chart.y_axis.title = "Total Lectures"
        chart.x_axis.title = "Teacher"
        chart.height = 10
        chart.width = 18

        data = Reference(ws_summary, min_col=2, min_row=1, max_row=16) 
        cats = Reference(ws_summary, min_col=1, min_row=2, max_row=16) 
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.shape = 4
        ws_dash.add_chart(chart, "B8")
        
        # Save to memory buffer
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    excel_file = generate_excel(df_summary, df_tidy)
    
    st.download_button(
        label="📥 Download Analyzed Excel Dashboard",
        data=excel_file,
        file_name="Processed_Timetable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )