# Install into your Google Sheet

**Your sheet:**  
https://docs.google.com/spreadsheets/d/1_Zf-ecWOLUg_BiJu9LuE0i7wmJDT9sn8b7HAWKFi98w/edit?usp=sharing

This agent cannot write to the sheet (access is Google-login only / not public).  
Run the installer **once inside the sheet** to build all tabs + sample data.

## 5-minute install

1. Open the sheet link above.
2. **Extensions → Apps Script**
3. Delete any default code.
4. Create these files and paste from `google-apps-script/` in this repo:

| File in Apps Script | Paste from |
|---------------------|------------|
| `Code.gs` | `Code.gs` |
| `Config.gs` | `Config.gs` |
| `Classification.gs` | `Classification.gs` |
| `CallData.gs` | `CallData.gs` |
| `AuditForm.gs` | `AuditForm.gs` |
| `Setup.gs` | `Setup.gs` |
| `WebApp.gs` | `WebApp.gs` |
| `Installer.gs` | `Installer.gs` |
| `AuditFormSidebar` (HTML) | `AuditFormSidebar.html` |
| `StandaloneAuditForm` (HTML) | `StandaloneAuditForm.html` |

5. **Save** the project.
6. Select function **`installSvaraAuditWorkbook`** → **Run** → **Authorize** (Allow).
7. Reload the spreadsheet.
8. Use **Audit → Open Standalone Audit Form**
9. Test UID: **`SV-10001`**

## What the installer creates

- `Svara Questions Classification` — Col B Rotating/Permanent, Col D parameters  
- `Calls Data` — sample UIDs + permanent fields  
- `Svara Audit Form` — fixed permanent block + Error/No error rotating questions  
- `Parameter Classification` — segregated list  
- `Audit Submissions` — submission log  
- **Audit** menu (sidebar, standalone form, fetch/submit macros)

## Share setting tip

To let tools/scripts read the sheet without login, set sharing to  
**Anyone with the link → Viewer** (optional). Edit access is only needed for people who audit.
