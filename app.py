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
    
    # Safely strip markdown backticks to prevent syntax errors
    raw_text = response.text
    raw_text = raw_text.replace('```csv', '')
    raw_text = raw_text.replace('```', '')
    csv_data = raw_text.strip()
    
    df = pd.read_csv(StringIO(csv_data))
    return df