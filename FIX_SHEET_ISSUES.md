# Fix all issues in your Google Sheet

**Your sheet:** https://docs.google.com/spreadsheets/d/1Tb-4H76nw6gyKdcQ1N4zbl1zSiM5UnDv5OuJiYal_rQ/edit

This cloud agent cannot sign into Google on your behalf, so repairs are applied by running **`fixAllSheetIssues`** inside **your** spreadsheet (Extensions → Apps Script).

## Issues fixed automatically

| Issue | Fix |
|-------|-----|
| **Timestamp `#REF!` — circular dependency** (column C) | Run **`fixTimestampCircularDependency`** or **`fixAllSheetIssues`** — rebuilds Timestamp from Audit date / Call Date (never self-references) |
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
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/fix-timestamp-circular-5b01/google-apps-script/PASTE_Code.gs
3. **Save**

### Step 2 — Ensure AuditForm HTML exists

1. If you don't already have an HTML file named **`AuditForm`**: **+ → HTML**, name it `AuditForm`
2. Paste from:  
   https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/fix-timestamp-circular-5b01/google-apps-script/PASTE_AuditForm.html
3. **Save**

### Step 3 — Run the fix

1. In Apps Script, select function **`fixAllSheetIssues`** → **Run** → **Allow**
2. Read the alert — it lists every repair applied
3. **Reload** the Google Sheet
4. **Audit → Open Standalone Audit Form** → test UID **`SV-10001`**

Or run **`fixTimestampCircularDependency`** alone if only the Timestamp column shows `#REF!`.

### Fastest fix (single function only)

1. **Extensions → Apps Script**
2. Paste [`FIX_TIMESTAMP_ONLY.gs`](./google-apps-script/FIX_TIMESTAMP_ONLY.gs) into `Code.gs` (replace all)
3. Run **`fixTimestampCircularDependency`** → **Allow** → reload sheet

## Manual fix (edit formula directly)

If **Timestamp** (column C) shows `#REF!` / *Circular dependency detected*:

1. Click **C2** (first data row under Timestamp)
2. Delete the broken formula
3. Paste (adjust **F** / **E** if your Audit date / Call Date columns differ):

```
=ARRAYFORMULA(IF(LEN($F2:$F), $F2:$F, IF(LEN($E2:$E), $E2:$E, )))
```

4. For **Week_Num** in **B2**:

```
=ARRAYFORMULA(IF(LEN($E2:$E), WEEKNUM($E2:$E), ))
```

5. Reload the sheet

## Fresh install instead?

If the sheet is empty or heavily broken, run **`installEverything`** instead of `fixAllSheetIssues`.  
That rebuilds all tabs (existing Calls Data / submissions on those tabs will be cleared).

## No Apps Script? Use the standalone HTML form

https://github.com/niranjanamurthy-ship-it/niranjana.murthy-navi.com/raw/cursor/fix-timestamp-circular-5b01/Svara_Audit_Form.html

Open in Chrome — works without the Audit menu. Download CSV and paste into **Audit Submissions**.
