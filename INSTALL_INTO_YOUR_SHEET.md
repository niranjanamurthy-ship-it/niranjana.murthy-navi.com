# Fix for SyntaxError — use TWO short files (not the giant single file)

Your paste was being **cut off** before the end of the file. Use these two smaller files instead.

**Sheet:** https://docs.google.com/spreadsheets/d/1Tb-4H76nw6gyKdcQ1N4zbl1zSiM5UnDv5OuJiYal_rQ/edit

## Already installed? Fix issues without wiping data

1. Update `Code.gs` from:  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/google-sheet-fixes-29e4/google-apps-script/PASTE_Code.gs
2. Run **`fixAllSheetIssues`** → reload → **Audit → Open Standalone Audit Form**
3. See [FIX_SHEET_ISSUES.md](./FIX_SHEET_ISSUES.md) for details

## Fresh install (two short files)

1. Open sheet → **Extensions → Apps Script**
2. Delete **everything** in `Code.gs`
3. Open this **raw** link, Select All, Copy:  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/google-sheet-fixes-29e4/google-apps-script/PASTE_Code.gs
4. Paste into `Code.gs` → **Save**

## Step 2 — AuditForm.html

1. In Apps Script: **+** next to Files → **HTML**  
2. Name it exactly: `AuditForm` (no `.html` in the name field)
3. Delete the default HTML inside it
4. Open this **raw** link, Select All, Copy:  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/google-sheet-fixes-29e4/google-apps-script/PASTE_AuditForm.html
5. Paste into `AuditForm` → **Save**

## Step 3 — Run

1. Select function **`installEverything`** → **Run** → **Allow**
2. Reload the Google Sheet
3. **Audit → Open Standalone Audit Form**
4. Test UID: **`SV-10001`**

## Check

- `Code.gs` should end with `doGet()` (around line 439)
- If Save still shows a syntax error, the paste was truncated again — use the raw links and Ctrl/Cmd+A before copying
