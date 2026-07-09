import io
import os
import pandas as pd
import streamlit as st
import google.generativeai as genai

# 1. Initialize the Gemini API (Using standard AI Studio Key, NOT Vertex AI)
# Set your API key as an environment variable or configure it directly
genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
model = genai.GenerativeModel('gemini-1.5-flash')

def clean_binary_csv(filename):
    """
    Executes the mandatory 5-step binary parsing protocol to prevent UnicodeDecodeErrors.
    """
    # Step 1: Read the file as raw bytes
    raw_bytes = open(filename, 'rb').read()
    
    # Step 2: Find the UTF-8 Byte Order Mark (BOM) index
    bom_index = raw_bytes.find(b'\xef\xbb\xbf')
    
    # Step 3: Slice byte array or fallback to ignore metadata block
    if bom_index != -1:
        cleaned_bytes = raw_bytes[bom_index:]
        text_stream = cleaned_bytes.decode('utf-8')
    else:
        text_stream = raw_bytes.decode('utf-8', errors='ignore')
        
    # Step 4: Pass the cleaned text string into pandas
    df = pd.read_csv(io.StringIO(text_stream))
    
    # Step 5: Clean up headers by stripping hidden spaces
    df.columns = [c.strip() for c in df.columns]
    
    return df

def run_extraction_pipeline(noa_pdf_path):
    st.info("Executing Data Ingestion Logic & Binary Parsing...")
    
    # Run your mandatory local verifications deterministically in Python
    sponsors_df = clean_binary_csv('Sponsors.csv')
    investigators_df = clean_binary_csv('Investigators.csv')
    orgs_df = clean_binary_csv('Rice University Orgs.csv')
    
    st.success("Reference datasets successfully parsed locally!")

    # Upload the Notice of Award PDF using the Gemini File API
    st.info("Uploading PDF to Gemini API...")
    pdf_file = genai.upload_file(path=noa_pdf_path, mime_type="application/pdf")
    
    # Prepare your systemic prompt instructions
    system_instruction = """
    You are an automated institutional entry system. Extract data according to Protocol 1.71.
    Analyze the attached PDF and cross-reference fields against the provided reference datasets.
    Return a strict JSON object mapping to the required fields.
    """
    
    # Inject your reference data directly into the text context window
    prompt = f"""
    --- REFERENCE DATASET: SPONSORS ---
    {sponsors_df.to_csv(index=False)}
    
    --- REFERENCE DATASET: INVESTIGATORS ---
    {investigators_df.to_csv(index=False)}
    
    --- REFERENCE DATASET: ORGS ---
    {orgs_df.to_csv(index=False)}
    
    Analyze the attached award PDF file against these datasets and output the JSON schema.
    """
    
    st.info("Querying Gemini Engine...")
    response = model.generate_content(
        [pdf_file, prompt],
        generation_config={"response_mime_type": "application/json"}
    )
    
    # Clean up the file from the cloud staging area after processing
    genai.delete_file(pdf_file.name)
    
    return response.text

# Simple Streamlit User Interface Layout
st.title("Award Processing & Extraction Deck")
uploaded_pdf = st.file_uploader("Upload Notice of Award (PDF)", type=["pdf"])

if uploaded_pdf and st.button("Process Document"):
    # Save uploaded file locally to pass to the API
    with open("temp_noa.pdf", "wb") as f:
        f.write(uploaded_pdf.getbuffer())
        
    json_result = run_extraction_pipeline("temp_noa.pdf")
    st.json(json_result)