import json
import os
from io import BytesIO

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import types

# load environment variables
load_dotenv()

# setup page
st.set_page_config(page_title="AI Doc Extractor", layout="wide")

st.title("AI-Powered Document to Excel Converter")
st.markdown("""
    **Project Goal:** Automate the conversion of unstructured PDF documents
    into well-organized, Excel-ready structured data using AI.
""")

# get api key
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    st.error(
        "GEMINI_API_KEY not found in environment variables. Please add it to your .env file."
    )
    st.stop()

def extract_with_gemini(file_bytes, api_key):
    # setup gemini client
    client = genai.Client(api_key=api_key)

    # extraction prompt for the ai
    prompt = """
    You are a Data Extraction Specialist. Your goal is to convert the provided document
    into a strict, structured JSON format optimized for Excel analysis.

    **Example:**
        - *Source Text:* "Born in the Pink City (Jaipur), which provides regional context."
        - *Expected Extraction:*
        {
          "Key": "Birth City",
          "Value": "Jaipur",
          "Comments": "Born and raised in the Pink City of India, provides regional profiling context."
        }

    SECTION 1: DATA FORMATTING STANDARDS
    
    **Date Handling:**
    - Convert ALL dates to DD-Mon-YY format where Mon is the abbreviated month name
    - Example: "March 15, 1989" becomes "15-Mar-89"
    
    **Name Processing:**
    - Always split full names into separate "First Name" and "Last Name" fields
    
    **Location Processing:**
    - Segregate location information into "City" and "State" fields where applicable
    
    **Currency Formatting:**
    - Keep currency values clean without unnecessary decimals (e.g., "350,000 INR")

    SECTION 2: COMMENT STRATEGY & CONTEXTUAL INFORMATION
    
    Comments are designed to capture supplementary information that enhances the primary
    key-value pair without cluttering the main data. Follow these guidelines:
    
    - Include comments only when additional context is present or valuable
    - Preserve the original wording from the source document in comments
    - Empty comments are acceptable when no extra information exists
    - Vague or ambiguous fields may contain only comments if the exact value is unclear

    SECTION 3: EMPLOYMENT & ORGANIZATIONAL DATA
    
    For employment history, create separate entries for current and previous positions:
    - Current Role: "Current Organisation", "Current Joining Date", "Current Salary", "Current Designation"
    - Previous Roles: Follow similar naming pattern with "Previous" prefix
    - Attach extra information as comments to respective entries

    SECTION 4: EDUCATIONAL QUALIFICATIONS
    
    Segregate education into distinct qualification levels:
    
    High School:
    - School name, 12th Standard pass-out year, grades
    
    Undergraduate Degree:
    - College name, graduation year, CGPA (prefixed with "Undergraduate")
    
    Graduate/Post-Graduate:
    - Institution name, completion year, CGPA (prefixed with "Graduate" or "Post-Graduate")

    SECTION 5: CERTIFICATIONS & PROFICIENCIES
    
    - List each certificate as a separate keyed entry (e.g., "Certificate 1", "Certificate 2")
    - Mention each certification only once
    - Include issuing organization, date, or other details as comments
    - Add a single "Proficiencies" field summarizing technical or professional skills

    FINAL INSTRUCTIONS
    
    - Return ONLY valid JSON; no explanatory text or conversation
    - Structure the data with # (serial number), Key, Value, and Comments columns
    - Use logical, human-readable field names that are self-explanatory
    - Handle unlisted fields intelligently using common sense and context from the document
    - Maintain consistency in naming conventions throughout
    """

    # send pdf to gemini with prompt
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"),
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json"
        ),
    )
    return response.text


# file upload
uploaded_file = st.file_uploader("Upload your PDF document", type=["pdf"])

if uploaded_file and st.button("Start Extraction"):
    with st.spinner("Analyzing document..."):
        try:
            # read file and extract with gemini
            file_bytes = uploaded_file.getvalue()
            json_result = extract_with_gemini(file_bytes, GEMINI_API_KEY)
            st.success("Extraction completed successfully")

            # parse json response
            data = json.loads(json_result)

            # handle nested json
            if isinstance(data, dict):
                data = list(data.values())[0]

            # convert to dataframe
            df = pd.DataFrame(data)

            # reorder columns
            if "Comments" in df.columns:
                cols = ["Key", "Value", "Comments"]
                df = df[[c for c in cols if c in df.columns]]

            # show preview
            st.dataframe(df, width="stretch")

            # add serial numbers
            df.insert(0, "#", range(1, len(df) + 1))

            # create excel file
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)

            # download button
            st.download_button(
                label="Download Output.xlsx",
                data=output.getvalue(),
                file_name="Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"An error occurred during extraction: {e}")
