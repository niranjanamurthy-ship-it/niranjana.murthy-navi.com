# Fix for SyntaxError — use TWO short files (not the giant single file)

Your paste was being **cut off** before the end of the file. Use these two smaller files instead.

**Sheet:** https://docs.google.com/spreadsheets/d/1_Zf-ecWOLUg_BiJu9LuE0i7wmJDT9sn8b7HAWKFi98w/edit?usp=sharing

## Step 1 — Code.gs

1. Open sheet → **Extensions → Apps Script**
2. Delete **everything** in `Code.gs`
3. Open this **raw** link, Select All, Copy:  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/svara-audit-form-14be/google-apps-script/PASTE_Code.gs
4. Paste into `Code.gs` → **Save**

## Step 2 — AuditForm.html

1. In Apps Script: **+** next to Files → **HTML**  
2. Name it exactly: `AuditForm` (no `.html` in the name field)
3. Delete the default HTML inside it
4. Open this **raw** link, Select All, Copy:  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/svara-audit-form-14be/google-apps-script/PASTE_AuditForm.html
5. Paste into `AuditForm` → **Save**

## Step 3 — Run

1. Select function **`installEverything`** → **Run** → **Allow**
2. Reload the Google Sheet
3. **Audit → Open Standalone Audit Form**
4. Test UID: **`SV-10001`**

## Check

- `Code.gs` should end with `doGet()` (around line 439)
- If Save still shows a syntax error, the paste was truncated again — use the raw links and Ctrl/Cmd+A before copying
