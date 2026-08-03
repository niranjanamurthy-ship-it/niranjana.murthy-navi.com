# Install into your sheet (2 minutes)

**Sheet:** https://docs.google.com/spreadsheets/d/1_Zf-ecWOLUg_BiJu9LuE0i7wmJDT9sn8b7HAWKFi98w/edit?usp=sharing

Remote agents still hit Google’s **sign-in wall**, even with Editor link sharing — so you run this once inside the sheet.

## Single-file install (recommended)

1. Open the sheet above  
2. **Extensions → Apps Script**  
3. Delete any default code in `Code.gs`  
4. Paste the entire contents of  
   [`google-apps-script/SvaraAudit_SINGLE_FILE.gs`](./google-apps-script/SvaraAudit_SINGLE_FILE.gs)  
5. **Save** → select function **`installEverything`** → **Run** → **Allow**  
6. Reload the spreadsheet  
7. **Audit → Open Standalone Audit Form**  
8. Test UID: **`SV-10001`**

That one paste creates:

- Classification (Rotating / Permanent + Col D questions)  
- Calls Data (sample UIDs)  
- Svara Audit Form (permanent + Error/No error)  
- Parameter Classification  
- Audit Submissions  
- Audit menu + standalone form (HTML included inline)

## Optional: service account (for remote automation)

If you want agents to write to the sheet without you pasting:

1. Create a Google Cloud service account  
2. Share the sheet with that service account email as **Editor**  
3. Paste the JSON key into a secure place this environment can read  
4. Ask the agent again to populate the sheet  

Until then, the single-file installer above is the fastest path.
