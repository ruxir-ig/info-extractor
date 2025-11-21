import json
import os
from io import BytesIO

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import types

# Load environment variables
load_dotenv()

# --- 1. SETUP & CONFIGURATION ---
st.set_page_config(page_title="AI Doc Extractor", layout="wide")

st.title("AI-Powered Document to Excel Converter")
st.markdown("""
    **Objective:** Convert unstructured PDF data into a structured Excel file.
""")

# --- 2. LOAD API KEY FROM ENVIRONMENT ---
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    st.error(
        "GEMINI_API_KEY not found in environment variables. Please add it to your .env file."
    )
    st.stop()

# --- 3. THE BRAIN (AI FUNCTIONS) ---


def extract_with_gemini(file_bytes, api_key):
    """Sends PDF bytes directly to Gemini 2.5 Flash."""
    client = genai.Client(api_key=api_key)

    prompt = """
    You are a Data Extraction Specialist. Your goal is to convert the provided text into a strict, granular Excel-ready JSON format.
    
    **Example:**
        - *Text:* "Born in the Pink City (Jaipur), which provides regional context."
        - *Extraction:*
        {
          "Key": "Birth City",
          "Value": "Jaipur",    
          "Comments": "Born and raised in the Pink City of India, provides regional profiling context."
        }
    
    
    1. FORMATTING STANDARDS 
    * **Dates:** Convert ALL dates to **DD-Mon-YY** where Mon stands for the abbreviated month name.
        - Example: "March 15, 1989" -> "15-Mar-89"
    * **Names:** Split full names into "First Name" and "Last Name".
    * **Locations:** Split full locations into "City" and "State", if required further splitting could be used.
    * **Currency:** Keep numbers clean (e.g., "350,000 INR").

    2. "RICH COMMENT" and Extra Information
    The comments are any extra information that is not directly related to the key-value pair, keep the wording in comments as it is from the source.
    
    Comments are not necessary for all keys, but if additional information is present, it should be included as a comment, but when not present leaving it blank also works.
    
    vague feilds can also have only comments, for eg. 
    
    You can put current data in separate keys like "Current Organisation", "Current Joining Date", "Current Salary", "Current Designation" and then add respective extra information as comments.
    
    Similare for any "Previous ..." data...
    
    Then in education, segregate, High School: name, 12th Standard pass out year, 12th grade,  Undergraduate Degree, College, year, CGPA, (with Undergarduate prefix), Same for Gradutation.
    
    Then list all certificates separately with a no. following each key, each certificate should be only mentioned once, and other information should be included as comments.
    
    and then also add proficiency (not in detail, 1 feild would work)
    
    Extract information from pdf document and return it in a JSON format.
    Directly give answer, no conversation is required.
    
    data should be in #, key, value, and comments for extra context related to data
        
    any other field should be handled appropriately, by your instincts and common sense, by using previously provided information.
    keep feild names logical and humanly understandable.
    
    
    """
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"),
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json"  # This forces valid JSON output
        ),
    )
    return response.text


# --- 4. MAIN APPLICATION LOGIC ---

uploaded_file = st.file_uploader("Upload Data Input.pdf", type=["pdf"])

if uploaded_file and st.button("Start Extraction"):
    with st.spinner("Analyzing document..."):
        try:
            file_bytes = uploaded_file.getvalue()

            # Extract with Gemini
            json_result = extract_with_gemini(file_bytes, GEMINI_API_KEY)
            st.success("Extraction Finished")

            # --- 5. PARSE & DISPLAY ---
            data = json.loads(json_result)

            # Flatten if the AI returned a nested object like {"result": [...]}
            if isinstance(data, dict):
                data = list(data.values())[0]

            df = pd.DataFrame(data)

            # Reorder columns to match requirements
            if "Comments" in df.columns:
                cols = ["Key", "Value", "Comments"]
                # Ensure only these cols exist (and handle missing ones)
                df = df[[c for c in cols if c in df.columns]]

            st.dataframe(df, use_container_width=True)

            # --- 6. DOWNLOAD ---
            # Add Serial Number (#) as per request
            df.insert(0, "#", range(1, len(df) + 1))

            # Convert to Excel in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)

            st.download_button(
                label="Download Output.xlsx",
                data=output.getvalue(),
                file_name="Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")
