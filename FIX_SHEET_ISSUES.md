# Fix all issues in your Google Sheet

**Your sheet:** https://docs.google.com/spreadsheets/d/1Tb-4H76nw6gyKdcQ1N4zbl1zSiM5UnDv5OuJiYal_rQ/edit

This cloud agent cannot sign into Google on your behalf, so repairs are applied by running **`fixAllSheetIssues`** inside **your** spreadsheet (Extensions → Apps Script).

## Issues fixed automatically

| Issue | Fix |
|-------|-----|
| Missing permanent parameters in classification (Audit date, QA Name, Call Quality Assessment, Penalty Timeline, Experiment Tag) | Adds rows marked **Permanent** in column B |
| Permanent fields wrongly marked **Rotating** in column B | Re-types them to **Permanent** |
| Missing / wrong Calls Data columns | Adds UID + all 16 permanent headers; restores sample UIDs if empty |
| Broken or outdated audit form layout | Rebuilds **Svara Audit Form** with Error/No error dropdowns |
| Incomplete Audit Submissions headers | Adds missing submission columns without deleting existing rows |
| Missing Audit menu | Rebuilds **Audit** menu after reload |

## Quick fix (recommended — 2 files)

### Step 1 — Update Code.gs

1. Open your sheet → **Extensions → Apps Script**
2. Replace all of `Code.gs` with the raw file:  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/google-sheet-fixes-29e4/google-apps-script/PASTE_Code.gs
3. **Save**

### Step 2 — Ensure AuditForm HTML exists

1. If you don't already have an HTML file named **`AuditForm`**: **+ → HTML**, name it `AuditForm`
2. Paste from:  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/google-sheet-fixes-29e4/google-apps-script/PASTE_AuditForm.html
3. **Save**

### Step 3 — Run the fix

1. In Apps Script, select function **`fixAllSheetIssues`** → **Run** → **Allow**
2. Read the alert — it lists every repair applied
3. **Reload** the Google Sheet
4. **Audit → Open Standalone Audit Form** → test UID **`SV-10001`**

## Fresh install instead?

If the sheet is empty or heavily broken, run **`installEverything`** instead of `fixAllSheetIssues`.  
That rebuilds all tabs (existing Calls Data / submissions on those tabs will be cleared).

## No Apps Script? Use the standalone HTML form

https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/google-sheet-fixes-29e4/Svara_Audit_Form.html

Open in Chrome — works without the Audit menu. Download CSV and paste into **Audit Submissions**.
