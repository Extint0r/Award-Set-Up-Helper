import io
import os
import pandas as pd

def clean_binary_csv(filepath):
    """
    Executes the mandatory 5-step data ingestion logic from Protocol 1.71 
    to prevent UnicodeDecodeErrors from proprietary binary headers.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Mandatory file missing at: {filepath}")
        
    # 1. Read the file as raw bytes
    raw_bytes = open(filepath, 'rb').read()
    
    # 2. Find the UTF-8 Byte Order Mark (BOM) index
    bom_index = raw_bytes.find(b'\xef\xbb\xbf')
    
    # 3. Slice byte array to drop metadata block or fallback to ignore errors
    if bom_index != -1:
        cleaned_bytes = raw_bytes[bom_index:]
        text_stream = cleaned_bytes.decode('utf-8')
    else:
        text_stream = raw_bytes.decode('utf-8', errors='ignore')
        
    # 4. Pass the cleaned text string into pd.read_csv
    df = pd.read_csv(io.StringIO(text_stream))
    
    # 5. Clean up resulting dataframe headers by stripping hidden spaces
    df.columns = [c.strip() for c in df.columns]
    
    return df

def load_all_registries(data_folder="data"):
    """
    Programmatically scans the environment and returns the three cleaned dataframes.
    """
    sponsors_path = os.path.join(data_folder, "Sponsors.csv")
    investigators_path = os.path.join(data_folder, "Investigators.csv")
    orgs_path = os.path.join(data_folder, "Rice University Orgs.csv")
    
    print("🔄 Running mandatory local environment scan...")
    sponsors_df = clean_binary_csv(sponsors_path)
    investigators_df = clean_binary_csv(investigators_path)
    orgs_df = clean_binary_csv(orgs_path)
    print("✅ Local environment clean. All dataframes loaded successfully.")
    
    return sponsors_df, investigators_df, orgs_df

# Quick sanity check validation loop
if __name__ == "__main__":
    try:
        s_df, i_df, o_df = load_all_registries()
        print(print(f"Loaded: {len(s_df)} Sponsors, {len(i_df)} Investigators, {len(o_df)} Orgs."))
    except Exception as e:
        print(f"❌ Validation failed during parser initialization: {e}")