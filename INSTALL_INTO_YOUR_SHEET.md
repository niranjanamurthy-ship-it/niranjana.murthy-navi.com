# Install into your sheet (2 minutes)

**Sheet:** https://docs.google.com/spreadsheets/d/1_Zf-ecWOLUg_BiJu9LuE0i7wmJDT9sn8b7HAWKFi98w/edit?usp=sharing

Remote agents still hit Google’s **sign-in wall**, even with Editor link sharing — so you run this once inside the sheet.

## Single-file install (recommended)

1. Open the sheet above  
2. **Extensions → Apps Script**  
3. Delete **all** code currently in `Code.gs`  
4. Open the **raw** file (not the GitHub preview page):  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/svara-audit-form-14be/google-apps-script/SvaraAudit_SINGLE_FILE.gs  
5. Select all → Copy → Paste into `Code.gs`  
6. **Save** → select function **`installEverything`** → **Run** → **Allow**  
7. Reload the spreadsheet  
8. **Audit → Open Standalone Audit Form**  
9. Test UID: **`SV-10001`**

> If you still see a syntax error, your paste was truncated. Use the raw link above and paste again (Ctrl/Cmd+A in the raw tab first).


## Optional: service account (for remote automation)

If you want agents to write to the sheet without you pasting:

1. Create a Google Cloud service account  
2. Share the sheet with that service account email as **Editor**  
3. Paste the JSON key into a secure place this environment can read  
4. Ask the agent again to populate the sheet  

Until then, the single-file installer above is the fastest path.
