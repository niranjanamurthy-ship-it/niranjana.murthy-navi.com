# Create your Google Sheet in ~60 seconds

This environment cannot log into your Google account, so a live `docs.google.com` link must be created under **your** Drive. The ready workbook is already built:

**Download:** [`Svara_QA_Audit_Form_READY.xlsx`](./Svara_QA_Audit_Form_READY.xlsx)

## Steps → live Google Sheet link

1. Open [Google Drive](https://drive.google.com) (signed in as your work account).
2. **New → File upload** → select `Svara_QA_Audit_Form_READY.xlsx`.
3. Right-click the uploaded file → **Open with → Google Sheets**.
4. **File → Save as Google Sheets** (if prompted).
5. Copy the browser URL — that is your sheet link (share with editors as needed).
6. Install macros: **Extensions → Apps Script** → paste files from `google-apps-script/` (see [DEPLOYMENT.md](./DEPLOYMENT.md)).
7. Run `buildAuditMenu` → authorize → reload → **Audit → Open Standalone Audit Form**.

## What’s already inside the workbook

| Tab | Contents |
|-----|----------|
| START HERE | Instructions |
| Svara Questions Classification | Col B = Rotating/Permanent, Col D = parameters |
| Calls Data | Sample UIDs `SV-10001` … `SV-10003` + permanent fields |
| Svara Audit Form | Fixed permanent block + Error/No error rotating dropdowns |
| Parameter Classification | Segregated rotating vs permanent list |
| Audit Submissions | Submission log headers |
| SETUP | Apps Script paste checklist |

## Test UIDs

`SV-10001` · `SV-10002` · `SV-10003`
