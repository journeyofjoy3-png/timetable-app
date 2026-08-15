import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
import datetime
import os
import sqlite3
import hashlib
import base64
import plotly.express as px
from streamlit_option_menu import option_menu
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils.dataframe import dataframe_to_rows
import google.generativeai as genai
from PIL import Image

# --- Password Hashing Helper ---
def make_hashes(password): return hashlib.sha256(str.encode(password)).hexdigest()

def check_hashes(password, hashed_text):
    if make_hashes(password) == hashed_text: return hashed_text
    return False

# --- Database Setup ---
def init_db():
    conn = sqlite3.connect('faculty_history.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS workload_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT, upload_date TEXT, week_label TEXT,
            teacher TEXT, class_lectures REAL, doubt_slots REAL, total_workload REAL,
            effective_capacity REAL, leaves REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT, role TEXT)''')
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", ("admin", make_hashes("matrix2026"), "admin"))
    conn.commit()
    conn.close()

init_db()

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
    df_to_save.rename(columns={'Teacher': 'teacher', 'Class Lectures': 'class_lectures', 'Doubt Slots': 'doubt_slots', 'Total Workload': 'total_workload', 'Effective Capacity': 'effective_capacity', 'Leave Days Count': 'leaves'}, inplace=True)
    df_to_save.to_sql('workload_history', conn, if_exists='append', index=False)
    conn.close()

# --- AI Image Parsing Logic ---
def extract_timetable_from_image(image_file, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-pro') 
    img = Image.open(image_file)
    
    prompt = """
    You are an expert data extraction AI. I am providing an image of a faculty timetable.
    Extract this timetable into a structured CSV format.
    The columns MUST EXACTLY match this comma-separated header:
    Teacher,Monday_M1,Monday_M2,Monday_M3,Monday_M4,Monday_E1,Monday_E2,Monday_E3,Monday_E4,Tuesday_M1,Tuesday_M2,Tuesday_M3,Tuesday_M4,Tuesday_E1,Tuesday_E2,Tuesday_E3,Tuesday_E4,Wednesday_M1,Wednesday_M2,Wednesday_M3,Wednesday_M4,Wednesday_E1,Wednesday_E2,Wednesday_E3,Wednesday_E4,Thursday_M1,Thursday_M2,Thursday_M3,Thursday_M4,Thursday_E1,Thursday_E2,Thursday_E3,Thursday_E4,Friday_M1,Friday_M2,Friday_M3,Friday_M4,Friday_E1,Friday_E2,Friday_E3,Friday_E4,Saturday_M1,Saturday_M2,Saturday_M3,Saturday_M4,Saturday_E1,Saturday_E2,Saturday_E3,Saturday_E4
    
    Rules:
    1. Read each row (each teacher). 
    2. If a cell in the image has a class name/batch, put that text in the corresponding CSV column.
    3. If a cell is empty or blank, leave the CSV value empty.
    4. Do NOT use markdown code blocks like ```csv. Output raw CSV text only.
    """
    response = model.generate_content([prompt, img])
    csv_data = response.text.replace("