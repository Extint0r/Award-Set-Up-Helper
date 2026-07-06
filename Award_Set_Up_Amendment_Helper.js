// =========================================================================
// ENTERPRISE AI RECONCILIATION ENGINE - PROTOCOL 1.71 (CONNECTED VERSION)
// =========================================================================

function runPreSubmissionAISweep() {
  const API_KEY = "YOUR_GEMINI_API_KEY"; 
  const INTAKE_FOLDER_ID = "1KtJXTCR6ZJzw0-K9Ue4ZQEFLONQmf8vB";
  const TARGET_SHEET_ID = "1sLkVHaOq4uuy9Jcmp3q2mKw5f1BVZHWxOMoMRQmyNMw";
  
  
  // 🏢 LIVE ORACLE DAILY EXTRACT FILE IDs (FROM YOUR DRIVE)
  const PROTOCOL_MD_ID = "1pojKi0890l75E-uxepww3-qPe8QXFuhL";   // <--- Add your 1.71.md file ID here
  const SPONSORS_CSV_ID = "1NjqBQ_Oo6jQMmAom3vIIG-U6BMAYiq77";
  const INVESTIGATORS_CSV_ID = "1BOaRWsoUwcuaPrkQpy8ozVpP3izTnJCd";
  const ORGS_CSV_ID = "1Z3s-f9xpILRahnnf7-_j3HJeFGePrST8";
  
  // 1. Fetch the active validation files from Drive as raw text strings
  Logger.log("📂 Ingesting live daily reference datasets from Google Drive...");
  const protocolTxt = DriveApp.getFileById(PROTOCOL_MD_ID).getBlob().getDataAsString(); // <--- Add this line
  const sponsorsTxt = DriveApp.getFileById(SPONSORS_CSV_ID).getBlob().getDataAsString();
  const investigatorsTxt = DriveApp.getFileById(INVESTIGATORS_CSV_ID).getBlob().getDataAsString();
  const orgsTxt = DriveApp.getFileById(ORGS_CSV_ID).getBlob().getDataAsString();

  
  // 2. Fetch the incoming NoA PDF file
  const folder = DriveApp.getFolderById(INTAKE_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.PDF);
  
  if (!files.hasNext()) {
    Logger.log("ℹ️ No new award PDFs found in the intake folder.");
    return;
  }
  
  const pdfFile = files.next();
  const pdfBlob = pdfFile.getBlob();
  const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
  
  Logger.log("🚀 Analyzing document: " + pdfFile.getName());

  // 3. Define the strict output schema
  const jsonSchema = {
  "type": "OBJECT",
  "properties": {
    "award_name": {"type": "STRING"},
    "sponsor_award_number": {"type": "STRING"},
    "aln_number": {"type": "STRING"},
    "award_purpose": {"type": "STRING"},
    "amendment_sequence_table": {
      "type": "STRING",
      "description": "A pipe-delimited summary of the history chain. Example: 'Base: $100k, End: 12/24 | Mod 1: +$50k, End: 12/25 | Mod 3: +$20k, End: 12/26'"
    },
    "sequence_gap_detected": {
      "type": "BOOLEAN",
      "description": "True if a numerical modification step is mathematically missing from the chain."
    },
    "validation_flags": {"type": "STRING"}
  },
  "required": ["award_name", "sponsor_award_number", "aln_number", "award_purpose", "amendment_sequence_table", "sequence_gap_detected", "validation_flags"]
};
 

  // 4. Incorporate the full 1.71.md protocol system instructions
    
  const systemInstruction = protocolTxt; // <--- The API now reads your entire markdown ruleset dynamically! 

  const apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + API_KEY;
  
  // 5. Construct the payload, injecting the CSV texts directly into the context stream
  const payload = {
    "systemInstruction": { "parts": [{ "text": systemInstruction }] },
    "contents": [{
      "parts": [
        { "inlineData": { "mimeType": "application/pdf", "data": base64Pdf } },
        { "text": "--- REFERENCE DATASET: SPONSORS ---\n" + sponsorsTxt },
        { "text": "--- REFERENCE DATASET: INVESTIGATORS ---\n" + investigatorsTxt },
        { "text": "--- REFERENCE DATASET: RICE ORGS ---\n" + orgsTxt },
        { "text": "Analyze the attached award PDF file and cross-reference its fields against the three text reference datasets above. Return JSON matching the schema." }
      ]
    }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "responseSchema": jsonSchema,
      "temperature": 0.1
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  // 6. Fire the request
  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseData = JSON.parse(response.getContentText());
  
  // 7. Parse the structured JSON output and write back to your spreadsheet
  if (responseData.candidates && responseData.candidates[0].content.parts[0].text) {
    const aiResult = JSON.parse(responseData.candidates[0].content.parts[0].text);
    
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sheet = ss.getSheets()[0];
    
    // ⚡️ UPDATED LOGIC ARRAY: Maps perfectly to Columns A through J
    sheet.appendRow([
      new Date(),                             // Column A: TIMESTAMP
      pdfFile.getName(),                      // Column B: FILE_NAME
      aiResult.award_name,                    // Column C: AWARD_NAME
      aiResult.sponsor_award_number,          // Column D: SPONSOR_AWARD_NUMBER
      aiResult.aln_number,                    // Column E: ALN_NUMBER
      aiResult.award_purpose,                 // Column F: AWARD_PURPOSE
      aiResult.amendment_sequence_table,      // Column G: AMENDMENT_SEQUENCE_TABLE
      aiResult.sequence_gap_detected,         // Column H: SEQUENCE_GAP_DETECTED
      aiResult.validation_flags,              // Column I: VALIDATION_FLAGS
      "Pending Review"                        // Column J: REVIEW_STATUS
    ]);
    
    Logger.log("🎉 Complete! Record verified and safely logged into the staging row.");
  } else {
    Logger.log("❌ Error parsing API response: " + response.getContentText());
  }
    
}