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
    # ROLE & TASK
    You are a highly precise data extraction specialist. Extract structured information from the document I provide and output it as a markdown table with extreme attention to detail and accuracy.
    
    # OUTPUT FORMAT
    Produce a markdown table with EXACTLY 5 columns:
    1. **#**: Serial number starting from 1
    2. **Key**: The field name/attribute
    3. **Value**: The extracted value
    4. **Comments**: Additional context (can keep blank)
    
    # EXTRACTION RULES
    
    ## 1. DATE HANDLING
    - Complete dates → ISO 8601: YYYY-MM-DD HH:MM:SS (e.g., "March 15, 1989" → "1989-03-15 00:00:00")
    - Partial dates → Use only available info (year-only → "2007")
    - Date ranges → Separate rows for start and end dates
    - Never invent missing day/month
    
    ## 2. NUMERIC VALUES
    - Remove commas (350,000 → 350000)
    - Separate value from unit in different rows
    - Scores: Value in Value column, maximum in Comments ("Out of 1000")
    
    ## 3. MULTI-PART ENTITIES
    Create separate rows per attribute with consistent prefixes:
    - **Current role**: Current Organization, Current Joining Date, Current Designation, Current Salary, Current Salary Currency
    - **Previous role**: Previous Organization, Previous Joining Date, Previous End Year, Previous Starting Designation
    - **First role**: Joining Date of first professional role, Designation of first professional role, Salary of first professional role, Salary currency of first professional role
    - **Education**: High School, 12th standard pass out year, 12th overall board score, Undergraduate degree, Undergraduate college, Undergraduate year, Undergraduate CGPA, Graduation degree, Graduation college, Graduation year, Graduation CGPA
    - **Certifications**: Certifications 1, Certifications 2, Certifications 3... (numbered sequentially)
    
    ## 4. INFORMATION STRATEGY
    - Read ENTIRE document first to understand context
    - Resolve all pronouns ("his", "he" → actual person)
    - Link related info across sentences
    - Calculate implicit info (age from birth date + reference year)
    
    ## 5. NAME & LOCATION PARSING
    - Split names: First Name, Last Name
    - Split locations: Birth City, Birth State
    
    ## 6. COMMENTS FIELD
    **Include**: Context from source, qualifiers ("outstanding", "with honors"), scale info ("On a 10-point scale"), rankings ("15th among 120"), temporal notes ("As of 2024", "Promoted in 2019")
    
    **Exclude**: Info that should be its own row, redundant restatements
    
    ## 7. LISTS
    - Enumerate: Certifications 1, Certifications 2...
    - Maintain chronological order
    
    ## 8. MISSING INFO
    - NEVER hallucinate or invent information
    
    ## 9. SPECIAL VALUES
    - Blood groups: Exact format (O+, A-)
    - Abbreviations: Preserve (B.Tech, M.Tech, AWS)
    - CGPA: Note scale in Comments
    
    ## 10. CONSISTENCY
    - Consecutive serial numbers starting from 1
    - Consistent Key naming throughout
    
    # QUALITY CHECKLIST
    - All serial numbers consecutive
    - All dates properly formatted
    - Currency amounts and codes in separate rows
    - Scores include maximum in Comments
    - No hallucinated information
    - ALL available information extracted
    
    Now extract all information from the following document:
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
