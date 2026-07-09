# test_ai.py
import json
import os
from ai_engine import extract_award_data

# Ensure a sample file actually exists before hitting the API
SAMPLE_FILE = "sample_award.pdf"

if not os.path.exists(SAMPLE_FILE):
    print(f"❌ Error: Please place a test PDF named '{SAMPLE_FILE}' in this folder first!")
else:
    try:
        print("📖 Reading 1.71.md instructions and loading local CSV dataframes...")
        print("🤖 Sending 'sample_award.pdf' to gemini-2.5-flash for structured analysis...")
        
        # Run the extraction engine
        extracted_data = extract_award_data(SAMPLE_FILE)
        
        print("\n🎉 SUCCESS! Gemini successfully parsed the document under Protocol 1.71.")
        print("\n--- EXTRACTED JSON STRUCTURE ---")
        print(json.dumps(extracted_data, indent=2))
        
    except Exception as e:
        print(f"\n❌ AI Extraction failed: {e}")
        print("💡 Check that your Gemini API key is pasted correctly inside 'ai_engine.py'.")