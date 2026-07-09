# test_sheets.py
from sheets_interface import initialize_workbook_tabs

# 1. Open your personal browser, create a blank Google Sheet, 
# and copy its ID from the address bar. Paste it here:
TEST_SHEET_ID = "15UPiZespLn_r2VhmCSOzlZSW1D9piwjdPHsijaqZKEU"

try:
    print("🔄 Initiating secure Google browser handshake...")
    print("👉 Look at your web browser! A login tab should open automatically.")
    
    # This function call will force gspread to look for credentials
    initialize_workbook_tabs(TEST_SHEET_ID)
    
    print("\n🎉 SUCCESS! Connection confirmed.")
    print("Go check your Google Sheet tabs in your browser. You should see all 5 operational tables created!")
except Exception as e:
    print(f"\n❌ Connection failed: {e}")