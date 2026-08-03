# Svara QA Audit Form — Google Sheets Deployment

This package adds a **macro-driven audit form** to your spreadsheet with:

- **Permanent parameters** (fixed): LAN, Call Date, Audit date, QA Name, Call Quality Assessment, Call Duration, Penalty Timeline, Experiment Tag, Recording Link, DPD, CX Preferred Language, Bot Spoken Language, Bot Disposition, Right Disposition according to Conversation, Audit Type, Bot Activity
- **Rotating parameters** (from column **D** on `Svara - New Questions & Classification`): questionnaire with **Error / No error** dropdowns
- **UID lookup**: fetch call data by Unique Identity Number from `Calls Data`
- **Toggle button**: show/hide rotating questions when not required
- **Submissions**: saved to `Audit Submissions`

---

## 1. Open your spreadsheet

[Your sheet](https://docs.google.com/spreadsheets/d/1rqFZjCFARzhIye_zWYBHe7gaQfZ3UI0TpSBH-4B4Un4/edit?gid=1884595185)

Ensure the tab is named exactly:

`Svara - New Questions & Classification`

---

## 2. Install Apps Script (5 minutes)

1. In the spreadsheet: **Extensions → Apps Script**
2. Delete any default `Code.gs` content
3. Create these script files and paste contents from `google-apps-script/`:

| Apps Script file | Source in repo |
|------------------|----------------|
| `Code.gs` | `google-apps-script/Code.gs` |
| `Config.gs` | `google-apps-script/Config.gs` |
| `Classification.gs` | `google-apps-script/Classification.gs` |
| `CallData.gs` | `google-apps-script/CallData.gs` |
| `AuditForm.gs` | `google-apps-script/AuditForm.gs` |
| `Setup.gs` | `google-apps-script/Setup.gs` |
| `AuditFormSidebar` (HTML) | `google-apps-script/AuditFormSidebar.html` |

4. **Save** the project (Ctrl/Cmd + S)
5. Run **`buildAuditMenu`** once → authorize when prompted
6. Reload the spreadsheet — you should see the **Audit** menu

---

## 3. Prepare classification sheet (Rotating vs Permanent)

On **`Svara - New Questions & Classification`**:

| Column | Purpose |
|--------|---------|
| **A** | Category (optional) |
| **B** | Type: write `Rotating` or `Permanent` |
| **D** | Parameter / question label |

Rows with **Permanent** in column B are excluded from the rotating questionnaire (they belong in the fixed header list). Rows with **Rotating** (or blank type) appear as questionnaire items with Error / No error.

Run: **Audit → Setup → Segregate Parameters (Col D → summary)**  
This creates **`Parameter Classification`** with the split.

> If your type column is **C** instead of B, change `TYPE_COLUMN: 3` in `Config.gs`.

---

## 4. Initialize the workbook

Run once: **Audit → Setup → Initialize Audit Workbook**

This creates:

- **`Calls Data`** — master call table (UID in column A + permanent fields)
- **`Svara Audit Form`** — in-sheet form with dropdowns
- **`Parameter Classification`** — rotating/permanent summary
- **`Audit Submissions`** — created on first submit

---

## 5. Load call data & audit

### Option A — Sidebar form (recommended)

1. **Audit → Open Audit Sidebar (form)**
2. Enter **UID** → **Fetch Call Data**
3. Permanent fields auto-fill from `Calls Data`
4. Answer rotating questions (Error / No error)
5. Use **Hide / Show Rotating Questions** when rotating section is not needed
6. **Submit Audit**

### Option B — In-sheet form

1. Open tab **`Svara Audit Form`**
2. Enter UID in **C4**
3. **Audit → Fetch Call Data by UID**
4. **Audit → Show/Hide Rotating Parameters**
5. **Audit → Submit Audit Record**

---

## 6. Calls Data format

| UID | LAN | Call Date | … | Bot Activity |
|-----|-----|-----------|---|--------------|
| SV-10001 | LAN-001 | 2026-08-01 | … | Active |

Header names must match permanent field names (or close match) for auto-mapping.

---

## Files in this repo

```
google-apps-script/
├── Code.gs
├── Config.gs
├── Classification.gs
├── CallData.gs
├── AuditForm.gs
├── Setup.gs
├── AuditFormSidebar.html
└── appsscript.json
```

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| No Audit menu | Run `buildAuditMenu` in Apps Script, reload sheet |
| Sheet not found | Match tab names in `Config.gs` |
| UID not found | Add row to `Calls Data` with matching UID |
| Wrong rotating list | Set column B to Rotating/Permanent, re-run Initialize or Create Audit Form |
