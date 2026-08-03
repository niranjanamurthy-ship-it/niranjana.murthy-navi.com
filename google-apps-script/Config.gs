/**
 * Svara QA Audit Form — configuration
 * Adjust sheet names / column letters if your workbook layout differs.
 */
var SVARA_CONFIG = {
  /** Source sheet: "Svara - New Questions & Classification" */
  CLASSIFICATION_SHEET: 'Svara - New Questions & Classification',

  /** Column with questionnaire parameter labels (user specified: column D) */
  PARAM_COLUMN: 4, // D

  /**
   * Column that marks Rotating vs Permanent.
   * Common layouts: B or C. Script reads text containing "rotat" or "perman".
   */
  TYPE_COLUMN: 2, // B — change to 3 for column C if needed

  /** Optional category / section column (for grouping questions) */
  CATEGORY_COLUMN: 1, // A

  /** First data row (below headers) */
  CLASSIFICATION_DATA_START_ROW: 2,

  /** Call master data sheet — UID lookup source */
  CALLS_DATA_SHEET: 'Calls Data',

  /** Audit submissions output sheet */
  AUDIT_SUBMISSIONS_SHEET: 'Audit Submissions',

  /** In-sheet audit form tab */
  AUDIT_FORM_SHEET: 'Svara Audit Form',

  /** UID column on Calls Data (default: A) */
  CALLS_UID_COLUMN: 1,

  /** Permanent field headers shown fixed on the form */
  PERMANENT_HEADERS: [
    'LAN',
    'Call Date',
    'Audit date',
    'QA Name',
    'Call Quality Assessment',
    'Call Duration',
    'Penalty Timeline',
    'Experiment Tag',
    'Recording Link',
    'DPD',
    'CX Preferred Language',
    'Bot Spoken Language',
    'Bot Disposition',
    'Right Disposition according to Conversation',
    'Audit Type',
    'Bot Activity'
  ],

  /** Dropdown options for rotating questionnaire */
  QUESTION_OPTIONS: ['Error', 'No error'],

  /** Row on audit form where rotating section starts (after permanent block) */
  ROTATING_SECTION_START_ROW: 22,

  /** Column used for form labels (B) and values (C) */
  LABEL_COL: 2,
  VALUE_COL: 3,
  CONTROL_COL: 4
};
