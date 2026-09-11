// ============================================================
//  SCHOOL DATABASE MANAGEMENT SYSTEM — Google Apps Script
//  Manha E-Solutions | Jaggayyapet, NTR District, AP
//  Backend: Code.gs  — v3
//
//  v3 CHANGE LOG (cross-referenced to the "Notes" sheet items)
//   1. Dashboard / fee totals now read ONLY from Tution_Fee — single
//      source of truth. Bus_Fee sheet is no longer created or read.
//   2. Add Student "Class" dropdown now numeric 1-10 (+ Nursery/LKG/UKG).
//   3. addStudent() now accepts and stores the full Student Profile set.
//   4. addStudent() now writes all 5 fee heads into Tution_Fee.
//   5. Student Profile now exposes fatherAadhar / motherAadhar;
//      Height/Weight removed from the data this function returns.
//   6. New: updateStudent(), updateStudentPhoto(), cancelPayment().
//   7. New: STUDENT_ID is the unique key everywhere — fixes the
//      duplicate-name billing collision.
//   8. Date bug already fixed pre-v3 — parsePaymentDate_ re-verified, unchanged.
//   9. getClassList() — independent class list; no longer depends on
//      Student Records having loaded first.
//  10. Bus_Fee sheet handling removed entirely (see SHEETS map below).
//  11. New: getPaymentConfirmation() + redesigned, head-wise WhatsApp templates.
//  12./13. Handled in Index.html (footer CSS fix; GAS-banner note only).
// ============================================================

var SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
var SCHOOL_NAME = "Manha E-Techno School";

// All uploaded photos (School Logo, Student Photo, Teacher/Staff Photo)
// are stored in this single Google Drive folder — Notes request: photos
// "should be stored in https://drive.google.com/drive/folders/1xj4BkiAqXh42X1bTfvj3Z9wd_rMsd3kZ
var PHOTO_FOLDER_ID = "1xj4BkiAqXh42X1bTfvj3Z9wd_rMsd3kZ";

var USERS = {
  "admin":     { password: "Manha02$",   role: "admin",     name: "Administrator" },
  "principal": { password: "Manha02$", role: "principal", name: "Principal"     },
  "clerk":     { password: "Manha02$",  role: "clerk",     name: "Clerk"  }
};

// ── FORGOT PASSWORD — recovery email per account ─────────────
// Password reset is restricted to the "admin" account only (see
// requestPasswordReset/resetPassword below) — Principal and Clerk
// cannot self-reset. Only the "admin" entry below actually needs a
// real email; the other two are kept here in case you extend this
// to other roles later.
var USER_RECOVERY_EMAIL = {
  "admin":     "manhaesolutions.citycentral@gmail.com",
  "principal": "shaikmalaika.1112@gmail.com",
  "clerk":     "abdulrafi8130@gmail.com"
};


// ── ROLE / PERMISSION MODEL ─────────────────────────────────
// Single source of truth for what each login may do. This mirrors the
// human-readable "User Roles & Permissions" sheet created automatically
// by ensureRolesSheet_() below.
//   admin     — full control: all data entry, all edits, all
//               cancellations, all deletions, and every Settings change.
//   principal — data entry + add / bulk-add students & teachers + view
//               every report. CANNOT edit, cancel, delete, or change any
//               Settings.
//   clerk     — data entry only (fee receipts, payments, expenses, marks,
//               attendance, add individual student) + view every report.
//               CANNOT bulk-add students/teachers, edit, cancel, delete,
//               or change any Settings.
var PERM_MSG = {
  ADMIN_ONLY: "This action is not permitted in this login \u2014 only the Admin can do this. Edits, cancellations, deletions and Settings changes are Admin-only.",
  MGMT_ONLY:  "This action is not permitted in this login \u2014 only the Admin or Principal can do this."
};
function requireAdmin_(sess) {
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  return null;
}
function requireManagement_(sess) {
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin" && sess.role !== "principal") return { error: PERM_MSG.MGMT_ONLY };
  return null;
}

// ── CONFIGURABLE LISTS (Fee Heads, Classes, Sections, Expense
//    Categories, Bank Accounts) ──
// v4: these are no longer hardcoded. They live in the "Settings" sheet so
// the school can add/rename/retire a class, section, fee head, expense
// category, or bank account entirely from inside the app — no code edit,
// no developer involvement. The arrays below are SEED DEFAULTS ONLY: the
// first time the app runs and the Settings sheet doesn't exist yet, these
// values are written into it once. After that, the Settings sheet is the
// single source of truth and these arrays are never read again.
var SEED_FEE_HEADS = [
  "Admission Fee", "Tuition Fee", "Bus Fee", "Books Fee", "Uniform Fee",
  "Badge Fee", "Exam Fee", "Hostel Fee", "Other"
];
var SEED_CLASSES = [
  { name: "Nursery", sections: ["A"] },
  { name: "LKG", sections: ["A"] },
  { name: "UKG", sections: ["A"] },
  { name: "1", sections: ["A"] }, { name: "2", sections: ["A"] },
  { name: "3", sections: ["A"] }, { name: "4", sections: ["A"] },
  { name: "5", sections: ["A"] }, { name: "6", sections: ["A"] },
  { name: "7", sections: ["A"] }, { name: "8", sections: ["A"] },
  { name: "9", sections: ["A"] }, { name: "10", sections: ["A"] }
];
// Seed exams: FA1-FA4 at 50 marks (pass 18 — roughly 35%), SA1-SA2 at
// 100 marks (pass 35%) — matches the user's example pattern exactly.
// Editable any time from Settings → Exam Pattern.
var SEED_EXAMS = [
  { name: "FA1", maxMarks: 50,  passMarks: 18, group: "FA" },
  { name: "FA2", maxMarks: 50,  passMarks: 18, group: "FA" },
  { name: "SA1", maxMarks: 100, passMarks: 35, group: "SA" },
  { name: "FA3", maxMarks: 50,  passMarks: 18, group: "FA" },
  { name: "FA4", maxMarks: 50,  passMarks: 18, group: "FA" },
  { name: "SA2", maxMarks: 100, passMarks: 35, group: "SA" }
];
// Seed class-subject mapping: a simpler subject set for primary classes
// (1-5) and a fuller subject set for higher classes (6-10), matching the
// user's "power classes 1-5 ... higher classes" distinction. Pre-primary
// (Nursery/LKG/UKG) intentionally left without exam subjects.
var SEED_CLASS_SUBJECTS = [
  { className: "1", subjects: ["Telugu", "English", "Maths", "EVS"] },
  { className: "2", subjects: ["Telugu", "English", "Maths", "EVS"] },
  { className: "3", subjects: ["Telugu", "English", "Maths", "Science", "Social"] },
  { className: "4", subjects: ["Telugu", "English", "Maths", "Science", "Social"] },
  { className: "5", subjects: ["Telugu", "English", "Maths", "Science", "Social"] },
  { className: "6", subjects: ["Telugu", "Hindi", "English", "Maths", "Science", "Social"] },
  { className: "7", subjects: ["Telugu", "Hindi", "English", "Maths", "Science", "Social"] },
  { className: "8", subjects: ["Telugu", "Hindi", "English", "Maths", "Science", "Social"] },
  { className: "9", subjects: ["Telugu", "Hindi", "English", "Maths", "Physical Science", "Biological Science", "Social"] },
  { className: "10", subjects: ["Telugu", "Hindi", "English", "Maths", "Physical Science", "Biological Science", "Social"] }
];
// Master subject list seed — admin can add more from Settings (Notes
// request: "subject add option... Computers, vocational, etc").
var SEED_SUBJECTS = [
  "Telugu", "Hindi", "English", "Maths", "EVS", "Science", "Social",
  "Physical Science", "Biological Science"
];
// The two subjects that combine into one "Science" result for grading
// purposes (Notes request, confirmed hardcoded by exact name rather than
// a generic configurable pairing). Each is entered and graded out of
// HALF the exam/class's normal max & pass marks (e.g. FA 50/18 per
// subject → 25/9 each for these two), but their marks SUM into a single
// Science component worth the full normal max when computing a
// student's overall total, percentage, and pass/fail.
var COMBINED_SCIENCE_SUBJECTS = ["Physical Science", "Biological Science"];

var SEED_EXPENSE_CATEGORIES = [
  { name: "Cash Deposit",          isTransfer: true  },
  { name: "Cash Withdrawal",       isTransfer: true  },
  { name: "Bank to Bank Transfer", isTransfer: true  },
  { name: "Stationery",            isTransfer: false },
  { name: "Repairs & Maintenance", isTransfer: false },
  { name: "Misc. Expenses",        isTransfer: false },
  { name: "Salaries",              isTransfer: false },
  { name: "Electricity Bill",      isTransfer: false },
  { name: "Bank Charges",          isTransfer: false },
  { name: "Others",                isTransfer: false }
];
var SEED_BANK_ACCOUNTS = []; // no defaults — added via Settings once real account details are known

// Receipt numbering: PREFIX/ACADEMIC-YEAR/SEQ, e.g. AHS/26-27/0001.
// The sequence continues automatically — clerk never type a receipt number.
var RECEIPT_PREFIX = "";   // leave blank to auto-derive from SCHOOL_NAME initials

// Student ID numbering: STU0001, STU0002, ... continues automatically.
// This is the new unique key used everywhere instead of the student's
// name — see Notes item #7 (duplicate-name billing fix).
var STUDENT_ID_PREFIX = "STU";

// Teacher ID numbering: TCH0001, TCH0002, ... — same pattern as Student ID.
var TEACHER_ID_PREFIX = "TCH";

var SHEETS = {
  PROFILE:      "Student Profile",
  TUITION:      "Tution_Fee",
  TRANSACTIONS: "Transactions",
  STAFF:        "Staff Details",       // kept for backward compatibility / historical data
  TEACHERS:     "Teacher Profile",     // new — replaces Staff Details going forward
  CONSOLIDATED: "Consilidated",
  EXPENSES:     "Expenses",
  SETTINGS:     "Settings",
  STAFF_ATTENDANCE:   "Staff Attendance",
  STUDENT_ATTENDANCE: "Student Attendance",
  MARKS:        "Marks",
  DROPOUTS:            "Dropouts",
  DROPOUT_FEE_HISTORY: "Dropout_Fee_History",
  DROPOUT_TRANSACTIONS: "Dropout_Transactions",
  ROLES:                "User Roles & Permissions"
};
var LOGIN_CRED_SHEET = "Login Credentials"; // backup record of each role's current active password
// NOTE (Notes item #10 — Bus_Fee Sheet): a separate "Bus_Fee" sheet is
// intentionally NOT listed here and is never created or read anywhere in
// this file. Bus fee now lives only as a column inside Tution_Fee. If an
// old Bus_Fee sheet still exists from before, it can be deleted manually —
// nothing in this code will ever recreate it.

// Settings sheet column map (0-based). One row per setting, grouped by
// Type. This single flexible sheet covers Classes+Sections, Fee Heads,
// Expense Categories, and Bank Accounts — see SETTINGS_TYPE below.
//   Type            | Key                | Value1      | Value2    | Value3 | Value4 | Active
//   Class           | <class name>       | "A,B,C"*    |           |        |        | TRUE
//   FeeHead         | <fee head name>    |             |           |        |        | TRUE
//   ExpenseCategory | <category name>    | TRUE/FALSE† |           |        |        | TRUE
//   BankAccount     | <account label>    | Acc No      | Bank Name | IFSC   | Branch | TRUE
//   * comma-separated section list      † isTransfer flag
var SET_COL = {
  TYPE: 0, KEY: 1, VALUE1: 2, VALUE2: 3, VALUE3: 4, VALUE4: 5, VALUE5: 6, ACTIVE: 7
};
var SETTINGS_TYPE = {
  CLASS: "Class", FEE_HEAD: "FeeHead", EXPENSE_CATEGORY: "ExpenseCategory", BANK_ACCOUNT: "BankAccount",
  SCHOOL_INFO: "SchoolInfo", EXAM: "Exam", CLASS_SUBJECTS: "ClassSubjects",
  SUBJECT: "Subject", EXAM_SUBJECT_MARK: "ExamSubjectMark", FEATURE_VISIBILITY: "FeatureVisibility",
  OPENING_BALANCE: "OpeningBalance", PAYMENT_MODE_VISIBILITY: "PaymentModeVisibility",
  ACADEMIC_CALENDAR: "AcademicCalendar"
};
// OpeningBalance: one row per account — Key="Cash" for the cash-in-hand
// opening balance, or Key=<bank account label> for each bank account's
// own opening balance (Notes request: confirmed one opening balance +
// date per bank account, entered once and editable later by an admin).
// Value1 = opening balance amount, Value2 = opening date (as a real
// Date). This is the anchor every Cash-in-Hand / Bank Balance figure is
// computed from: Opening + money in - money out, exactly like a real
// passbook.
// FeatureVisibility: single-row settings entry (Key is always "Main"),
// same single-row pattern as SchoolInfo above. Value1/2/3 hold TRUE/FALSE
// for whether Marks Entry / Staff Attendance / Student Attendance show
// up as nav items at all — Notes request: a show/hide switch in
// Settings that hides the tab itself, not just its content, when off.
// Subject: Key = subject name (e.g. "Computers", "Vocational"). Master
// list a class's subject mapping is built from — lets the school add a
// new subject (Computers, Vocational, etc.) once and reuse it across
// any class/exam, rather than retyping free text everywhere.
//
// ExamSubjectMark: Key = "Class|Exam|Subject" composite (e.g.
// "9|FA1|Physical Science"), Value1 = max marks, Value2 = pass marks
// FOR THAT SPECIFIC COMBINATION. This is what lets pass marks (and max
// marks) vary per subject per exam per class — e.g. FA exams are 50/18
// for most subjects, but Physical Science / Biological Science are
// 25/9 each within the same FA exam. A row here is only created the
// first time that combination is actually edited away from the exam's
// plain default; until then, getEffectiveSubjectMarks_() below
// transparently falls back to the exam's own maxMarks/passMarks (or
// half of it, for the two combined-science subjects specifically).
// Exam: Key = exam name (e.g. "FA1"), Value1 = max marks, Value2 = pass
// marks, Value3 = exam group ("FA" or "SA", used for report-card
// grouping/labels only).
// ClassSubjects: Key = class name, Value1 = comma-separated subject list
// for that class (e.g. "Telugu,English,Maths,Science,Social"). One row
// per class, same pattern as the Class/Sections settings.
// SchoolInfo is a single-row settings type (Key is always "Main") holding
// the letterhead details used on printed fee receipts — Notes request:
// "create print option ... with all payment related details". Stored in
// Settings rather than hardcoded so the school can fill in its own name,
// address, phone and email from inside the app with no code change.
//   Type        | Key  | Value1 (Name) | Value2 (Address) | Value3 (Phone) | Value4 (Email)

// Student Profile column map (0-based, matches header row index 1).
// Column 0 (STUDENT_ID) is new in v3. HEIGHT / WEIGHT have been dropped —
// Father/Mother Aadhar now take their place on the profile screen
// (Notes item #5). v4 adds SECTION (col 12) — see SETTINGS_TYPE.CLASS for
// how the section list per class is configured.
var SP_COL = {
  STUDENT_ID: 0, NAME: 1, AADHAR: 2, CHILD_ID: 3, PEN: 4, APAAR: 5, JOIN_DATE: 6, ADM_NO: 7,
  DOB: 8, AGE: 9, GENDER: 10, CLASS: 11, SECTION: 12, TRANSPORT: 13, ADDRESS: 14, CASTE: 15,
  SUBCASTE: 16, FATHER_NAME: 17, FATHER_AADHAR: 18, FATHER_OCC: 19, FATHER_MOBILE: 20,
  MOTHER_NAME: 21, MOTHER_AADHAR: 22, MOTHER_OCC: 23, MOTHER_MOBILE: 24,
  RATION_CARD: 25, BANK_ACC: 26, BANK_NAME: 27, IFSC: 28, BRANCH: 29,
  PHOTO: 30, STATUS: 31, UDISE: 32, REMARKS: 33
};
var SP_NUM_COLS = 34; // total columns written per student row

// Tution_Fee column map (0-based). v5: restructured into LONG FORMAT —
// one row per student PER FEE HEAD, instead of one row per student with
// a fixed column for each head. This is what lets the school create any
// number of custom fee heads in Settings (Notes request) without ever
// needing a new column here. Student ID is still the unique key tying
// each row back to Student Profile.
var TF_COL = {
  STUDENT_ID: 0, NAME: 1, CLASS: 2, SECTION: 3, FEE_HEAD: 4, BILLED: 5, DISCOUNT: 6, NET: 7, ACADEMIC_YEAR: 8
};
var TF_NUM_COLS = 9;

// Transactions column map (0-based). v3 adds Student ID (col 1) and a
// Voided flag; v4 adds a Bank Account column (col 10) recording which
// Settings-configured account a non-cash payment went into.
var TX_COL = {
  DATE: 0, STUDENT_ID: 1, STUDENT_NAME: 2, CLASS: 3, RECEIPT: 4,
  HEAD: 5, MODE: 6, AMOUNT: 7, BY: 8, VOID: 9, BANK_ACCOUNT: 10
};

// Expenses column map (0-based). Same audit-trail pattern as Transactions —
// a cancelled expense is voided (struck through, excluded from totals),
// never deleted. v4 adds BANK_ACCOUNT (which Settings bank account a
// non-cash expense, deposit, or withdrawal affected).
var EXP_COL = {
  DATE: 0, VOUCHER: 1, CATEGORY: 2, DESCRIPTION: 3, MODE: 4, AMOUNT: 5, BY: 6, VOID: 7, BANK_ACCOUNT: 8, TO_BANK_ACCOUNT: 9
};
var EXP_NUM_COLS = 10;

// Teacher Profile column map (0-based) — mirrors the school's existing
// Staff Details sheet structure exactly (Notes request: "Teacher profile
// update like student profile with all the information with add or
// modify options"), with TEACHER_ID prepended as the unique key, the
// same role Student ID plays in Student Profile.
var TEACHER_COL = {
  TEACHER_ID: 0, NAME: 1, AADHAR: 2, PAN: 3, NATIONAL_CODE: 4, GENDER: 5, DOB: 6,
  FATHER_NAME: 7, MOTHER_NAME: 8, CASTE: 9, SUBCASTE: 10, MARITAL_STATUS: 11,
  CONTACT_NO: 12, EMAIL: 13, ADDRESS: 14, ACADEMIC_QUAL: 15, PROFESSIONAL_QUAL: 16,
  EXPERIENCE: 17, DEPARTMENT: 18, SCHOOL_TYPE: 19, SUBJECT: 20, DESIGNATION: 21,
  DOJ: 22, BANK_ACC: 23, BANK_NAME: 24, IFSC: 25, BRANCH: 26, UDISE: 27,
  PHOTO: 28, UAN: 29, ESI: 30, ABHA: 31, NOMINEE_FATHER: 32, NOMINEE_MOTHER: 33,
  NOMINEE_SPOUSE: 34, STATUS: 35, REMARKS: 36
};
var TEACHER_NUM_COLS = 37;

// Staff Attendance column map (0-based). One row per teacher PER DAY.
// Status is one of "Present", "Absent", "Leave" — used for payroll to
// compute days present / leaves taken over any date range.
var SATT_COL = {
  DATE: 0, TEACHER_ID: 1, NAME: 2, STATUS: 3, REMARKS: 4, MARKED_BY: 5
};
var SATT_NUM_COLS = 6;

// Student Attendance column map (0-based). One row per student PER DAY —
// daily Present/Absent marking (like staff), replacing the old
// consolidated monthly entry. A monthly summary is computed on read.
var STATT_COL = {
  DATE: 0, STUDENT_ID: 1, NAME: 2, CLASS: 3, SECTION: 4, STATUS: 5, REMARKS: 6, MARKED_BY: 7
};
var STATT_NUM_COLS = 8;
var STATT_DAILY_HEADER = ["Date", "Student ID", "Name", "Class", "Section", "Status", "Remarks", "Marked By"];

// Marks column map (0-based). One row per student PER SUBJECT PER EXAM.
// Pass/fail and rank are computed on read, not stored, so changing the
// pass-mark setting later re-evaluates pass/fail for existing marks
// automatically rather than leaving stale data behind.
var MARKS_COL = {
  STUDENT_ID: 0, NAME: 1, CLASS: 2, SECTION: 3, EXAM: 4, SUBJECT: 5,
  MARKS_OBTAINED: 6, MAX_MARKS: 7, PASS_MARKS: 8, ACADEMIC_YEAR: 9, ENTERED_BY: 10
};
var MARKS_NUM_COLS = 11;


// ── ENTRY POINT ─────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("School DBMS — Manha E-Solutions")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── SESSION ─────────────────────────────────────────────────
// IMPORTANT: sessions are stored in CacheService (shared, fast) keyed by token,
// NOT in PropertiesService.getUserProperties(). UserProperties in a web app
// deployed as "Execute as: Me" is scoped per *effective user*, which can behave
// inconsistently across reloads/cookies and was the root cause of random
// logouts on refresh. CacheService keyed by token survives page reloads
// reliably as long as the token itself is preserved client-side (sessionStorage).
var SESSION_TTL_SEC = 8 * 60 * 60; // 8 hours

// A password reset (see resetPassword() below) is saved as an override
// in Script Properties rather than rewriting the USERS constant above
// (a running script can't safely edit its own source file). login()
// checks for that override first and falls back to the hardcoded
// default in USERS if no reset has ever happened for this account.
function getCurrentPassword_(u) {
  if (!USERS[u]) return null;
  var override = PropertiesService.getScriptProperties().getProperty("pwd_" + u);
  return (override !== null && override !== "") ? override : USERS[u].password;
}

function login(username, password) {
  var u = (username || "").toLowerCase().trim();
  var current = getCurrentPassword_(u);
  if (current !== null && current === password) {
    var token = Utilities.getUuid();
    var payload = JSON.stringify({
      user: u, role: USERS[u].role, name: USERS[u].name, ts: Date.now()
    });
    CacheService.getScriptCache().put("sess_" + token, payload, SESSION_TTL_SEC);
    return { success: true, token: token, role: USERS[u].role, name: USERS[u].name };
  }
  return { success: false, message: "Invalid username or password." };
}

// Keeps a "Login Credentials" sheet in sync with the actual active
// password for every role, so Admin has a visible backup record —
// not just whatever's hidden in Script Properties. Called after every
// successful password change. Rewrites all 3 rows (cheap, only 3 roles)
// so the sheet is always accurate even if it drifted or was hand-edited.
function syncLoginCredentialsSheet_(changedRole, method, changedBy) {
  var ss = _getSS_();
  var sh = ss.getSheetByName(LOGIN_CRED_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOGIN_CRED_SHEET);
    sh.appendRow(["Role", "Current Password", "Last Changed Via", "Last Changed At", "Last Changed By"]);
    sh.setFrozenRows(1);
  }

  var data = sh.getDataRange().getValues();
  var rowByRole = {};
  for (var i = 1; i < data.length; i++) rowByRole[String(data[i][0]).toLowerCase()] = i + 1;

  var now = new Date();
  VALID_RESET_ROLES.forEach(function(r) {
    var pwd = getCurrentPassword_(r);
    var rowIdx = rowByRole[r];
    var isChanged = (r === changedRole);
    if (!rowIdx) {
      sh.appendRow([r, pwd, isChanged ? method : "", isChanged ? now : "", isChanged ? changedBy : ""]);
    } else {
      sh.getRange(rowIdx, 2, 1, 1).setValue(pwd); // keep password column accurate for every role
      if (isChanged) sh.getRange(rowIdx, 3, 1, 3).setValues([[method, now, changedBy]]);
    }
  });
}

// ── PASSWORD RESET (any role, admin approves) ────────────────
// Any account (Admin / Principal / Clerk) can request a reset. The
// Admin is always the approver — there are two ways a request gets
// resolved, whichever happens first:
//   (a) INSTANT — an OTP is emailed to the Admin's recovery inbox.
//       Whoever has that OTP (normally the Admin) enters it on the
//       same reset page along with a new password, right away.
//   (b) IN-APP — the request just sits pending. Next time the Admin
//       logs into the app, a notification lists it with Accept /
//       Reject buttons. Accept opens a "set new password" popup.
// Both paths write to the same pending-request record, so whichever
// happens first resolves it and the other becomes stale automatically.
var VALID_RESET_ROLES = ["admin", "principal", "clerk"];

function requireAdminToken_(token) {
  var sess = validateSession(token);
  if (!sess || sess.role !== "admin") return null;
  return sess;
}

// Step 1: requester picks a role on the (logged-out) reset page.
function requestPasswordReset(role) {
  var r = (role || "").toLowerCase().trim();
  if (VALID_RESET_ROLES.indexOf(r) === -1) {
    return { success: false, message: "Please select a valid role." };
  }

  var adminEmail = USER_RECOVERY_EMAIL["admin"];
  if (!adminEmail || adminEmail.indexOf("REPLACE_WITH_") === 0) {
    return { success: false, message: "No recovery email is configured for Admin yet. Contact your system administrator to set one up." };
  }

  var code = String(Math.floor(100000 + Math.random() * 900000));
  PropertiesService.getScriptProperties().setProperty("pwdreq_" + r, JSON.stringify({
    otp: code, ts: Date.now(), role: r
  }));

  MailApp.sendEmail({
    to: adminEmail,
    subject: SCHOOL_NAME + " — Password Reset Request (" + r + ")",
    body: "A password reset was requested for the \"" + r + "\" login.\n\n" +
          "To approve INSTANTLY: give this OTP to whoever is on the reset page now: " + code + "\n" +
          "(It will ask for this OTP plus the new password.)\n\n" +
          "Prefer to review it yourself? Just log into the Admin dashboard — you'll see this request " +
          "waiting there with Accept / Reject buttons, no OTP needed.\n\n" +
          "If you didn't expect this, you can ignore it or reject it from the dashboard."
  });

  return { success: true, message: "Request sent to Admin. Enter the OTP once you have it to reset instantly, or wait for Admin to approve it from their dashboard." };
}

// Step 2 (instant path): whoever holds the OTP submits it + a new
// password right here — this IS the approval, no separate Accept click needed.
function resetPasswordWithOtp(role, otp, newPassword) {
  var r = (role || "").toLowerCase().trim();
  if (VALID_RESET_ROLES.indexOf(r) === -1) {
    return { success: false, message: "Please select a valid role." };
  }
  var raw = PropertiesService.getScriptProperties().getProperty("pwdreq_" + r);
  if (!raw) return { success: false, message: "No pending request for this role. Submit a new request first." };

  var req;
  try { req = JSON.parse(raw); } catch (e) { return { success: false, message: "Request data is corrupted. Submit a new request." }; }

  if (req.otp !== String(otp || "").trim()) {
    return { success: false, message: "Incorrect OTP." };
  }
  if (!newPassword || newPassword.length < 6) {
    return { success: false, message: "New password must be at least 6 characters." };
  }

  PropertiesService.getScriptProperties().setProperty("pwd_" + r, newPassword);
  PropertiesService.getScriptProperties().deleteProperty("pwdreq_" + r);
  syncLoginCredentialsSheet_(r, "OTP (self-service)", r + " (via OTP)");
  return { success: true, message: "Password reset for \"" + r + "\" — it can be used to log in now." };
}

// In-app path (Admin only, must be logged in): list pending requests
// so the dashboard can show a notification with Accept/Reject.
function getPendingPasswordResetRequests(token) {
  if (!requireAdminToken_(token)) return { success: false, message: "Admin login required." };
  syncLoginCredentialsSheet_(null, "", ""); // refresh sheet so Admin sees current passwords even if nothing changed yet
  var props = PropertiesService.getScriptProperties();
  var out = [];
  VALID_RESET_ROLES.forEach(function(r) {
    var raw = props.getProperty("pwdreq_" + r);
    if (!raw) return;
    try {
      var req = JSON.parse(raw);
      out.push({ role: r, ts: req.ts });
    } catch (e) {}
  });
  return { success: true, requests: out };
}

// In-app path: Admin clicks Accept, sets a new password directly —
// no OTP involved since the Admin is already authenticated.
function approvePasswordReset(token, role, newPassword) {
  var sess = requireAdminToken_(token);
  if (!sess) return { success: false, message: "Admin login required." };
  var r = (role || "").toLowerCase().trim();
  if (VALID_RESET_ROLES.indexOf(r) === -1) return { success: false, message: "Invalid role." };
  if (!newPassword || newPassword.length < 6) {
    return { success: false, message: "New password must be at least 6 characters." };
  }
  PropertiesService.getScriptProperties().setProperty("pwd_" + r, newPassword);
  PropertiesService.getScriptProperties().deleteProperty("pwdreq_" + r);
  syncLoginCredentialsSheet_(r, "Admin Approved", sess.name || "admin");
  return { success: true, message: "Password reset for \"" + r + "\"." };
}

// In-app path: Admin clicks Reject — discards the request, no password change.
function rejectPasswordReset(token, role) {
  if (!requireAdminToken_(token)) return { success: false, message: "Admin login required." };
  var r = (role || "").toLowerCase().trim();
  if (VALID_RESET_ROLES.indexOf(r) === -1) return { success: false, message: "Invalid role." };
  PropertiesService.getScriptProperties().deleteProperty("pwdreq_" + r);
  return { success: true, message: "Request for \"" + r + "\" rejected." };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove("sess_" + token);
  return { success: true };
}

// Called by the client right after page load to silently restore a session
// from a token still held in sessionStorage, instead of forcing a fresh login.
function resumeSession(token) {
  var sess = validateSession(token);
  if (!sess) return { success: false };
  return { success: true, role: sess.role, name: sess.name };
}

function validateSession(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get("sess_" + token);
  if (!raw) return null;
  try {
    var data = JSON.parse(raw);
    // sliding expiry: refresh the TTL on every validated call
    CacheService.getScriptCache().put("sess_" + token, raw, SESSION_TTL_SEC);
    return data;
  } catch (e) {
    return null;
  }
}

// ── SHEET HELPER — request-scoped cache ─────────────────────
// SpreadsheetApp.openById() makes a network call to the Sheets API
// every time it is called. With 100+ getSheet() calls per request
// that was adding 5-8 seconds of pure API round-trip latency.
// Solution: open the spreadsheet ONCE per GAS execution and cache
// every sheet object in a plain JS object — subsequent getSheet()
// calls return the in-memory reference instantly (no network call).
var _ssCache_ = null;
var _sheetCache_ = {};

function _getSS_() {
  if (!_ssCache_) _ssCache_ = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache_;
}

function getSheet(name) {
  if (_sheetCache_[name]) return _sheetCache_[name];
  var ss = _getSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  _sheetCache_[name] = sh;
  return sh;
}

// Call after any write that changes profiles, fees, or settings so the
// next read in the same request (e.g. a dashboard reload) sees fresh data.
function invalidateCaches_() {
  _cachedProfiles_ = null;
  _cachedFeeMap_   = null;
  _cachedSettings_ = null;
  _settingsSeeded_ = false;
  // Sheet objects themselves remain valid — only the *data* is stale.
}

function safeNum(v) { var n = Number(v); return isNaN(n) ? 0 : n; }
function safeStr(v) { return (v === null || v === undefined) ? "" : String(v); }
function isVoided_(v) { return v === true || safeStr(v).toUpperCase() === "TRUE"; }
// BUGFIX: Settings checkbox cells (e.g. the "isTransfer" column for Expense
// Categories) are stored as real Boolean TRUE/FALSE by Sheets, not the text
// "TRUE". safeStr(true) returns the lowercase string "true", so the old
// strict `=== "TRUE"` checks scattered around the codebase silently failed
// for every boolean cell and only ever matched the literal text "TRUE".
// This was the root cause of "Bank to Bank Transfer" (and any other
// checkbox-driven transfer category) being wrongly counted as a real
// expense in the P&L / Balance Sheet even though its isTransfer flag was
// correctly set to TRUE in Settings. Use this helper everywhere a
// Settings TRUE/FALSE flag is read.
function isTrueVal_(v) { return v === true || safeStr(v).toUpperCase() === "TRUE"; }

// Parses a "YYYY-MM-DD" string (as produced by an <input type="date">) into a
// Date object at local midnight, avoiding the timezone-shift bug you get from
// `new Date("YYYY-MM-DD")` directly. Falls back to "now" if blank/invalid —
// this is what lets Clerk pick a past date when recording a late payment.
// (Notes item #8 — confirmed already fixed pre-v3; logic re-verified, unchanged.)
function parsePaymentDate_(dateStr) {
  if (!dateStr) return new Date();
  var m = String(dateStr).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return new Date();
  var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? new Date() : d;
}

// School academic year runs Jun–May, matching getMonthlyCollectionTrend().
// e.g. 12-Jul-2026 -> "26-27", 03-Mar-2027 -> "26-27".
function getAcademicYear_(date) {
  date = date || new Date();
  var y = date.getFullYear();
  var startYear = (date.getMonth() >= 5) ? y : y - 1; // month 5 = June
  return String(startYear).slice(-2) + "-" + String(startYear + 1).slice(-2);
}

// ── ACTIVE ACADEMIC YEAR ────────────────────────────────────
// getCurrentAcademicYear_() below is purely calendar-derived. This is
// deliberately DIFFERENT: it's the year every report defaults to, and it
// only moves forward when runClassPromotion() actually runs — not the
// moment the calendar ticks past May 31st. That matters because the
// school doesn't switch over in practice until an admin has actually
// promoted students; until then, "today" can already be in June while
// every live record on the Profile/Tuition_Fee sheets still belongs to
// the old year. Persisted in Script Properties, seeded from the
// calendar-derived year the very first time it's ever read.
function getActiveAcademicYear_() {
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('ACTIVE_ACADEMIC_YEAR');
  if (stored) return stored;
  var seeded = getCurrentAcademicYear_();
  props.setProperty('ACTIVE_ACADEMIC_YEAR', seeded);
  return seeded;
}

// "25-26" -> "26-27". Used to advance the active year by exactly one
// step when promotion runs, regardless of what the calendar says.
function _nextAcademicYear_(ay) {
  var parts = String(ay).split('-');
  function pad(n) { n = ((n % 100) + 100) % 100; return (n < 10 ? '0' : '') + n; }
  return pad(parseInt(parts[0], 10) + 1) + '-' + pad(parseInt(parts[1], 10) + 1);
}

// Every academic year we actually have data for — the active year, plus
// any past years already archived via Promote Students — for year-picker
// dropdowns in Fee Management and elsewhere.
function getAcademicYearOptions(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var active = getActiveAcademicYear_();
  var years = {};
  years[active] = true;
  _getSS_().getSheets().forEach(function(sh) {
    var m = /^Students_(\d{2}-\d{2})$/.exec(sh.getName());
    if (m) years[m[1]] = true;
  });
  return { success: true, active: active, years: Object.keys(years).sort() };
}

// Fee Receivable for the Balance Sheet, routed by which academic year the
// "As On" date actually falls in — see the comment at its one call site
// in getBalanceSheetData_ for why this split exists.
//
// Notes request: management wants to see outstanding receivables broken
// down HEAD-WISE (e.g. Tuition Fee ₹X, Transport Fee ₹Y) instead of one
// consolidated "Fee Receivable" figure, so they can tell which head is
// actually driving the dues. Returns { total, byHead, headWiseAvailable }:
//   - byHead is an ordered [{head, amount}] list, biggest first.
//   - headWiseAvailable is false for an ARCHIVED year, because the
//     Students_<year> snapshot only stores each student's Billed/Paid/Due
//     TOTALS, not a per-fee-head split (see getStudentListForYear()'s
//     comment) — so only the consolidated total is possible there.
function getFeeReceivableAsOf_(token, asOnDate) {
  var ay = getAcademicYear_(asOnDate);
  var sh = (ay === getActiveAcademicYear_()) ? null : _getSS_().getSheetByName("Students_" + ay);
  if (!sh) {
    // Current year (or a year that was never archived) — best available
    // figure is today's live Outstanding list, which already carries a
    // per-student per-head breakdown (see getOutstandingList_ / getFeeMap_).
    var list = getOutstandingList_(token);
    var byHeadMap = {}, headOrder = [];
    var total = 0;
    list.forEach(function(o) {
      total += o.due;
      var heads = o.heads || {};
      Object.keys(heads).forEach(function(h) {
        var due = heads[h] && heads[h].due;
        if (!due || due <= 0) return;
        if (!byHeadMap.hasOwnProperty(h)) { byHeadMap[h] = 0; headOrder.push(h); }
        byHeadMap[h] += due;
      });
    });
    var byHead = headOrder.map(function(h) { return { head: h, amount: byHeadMap[h] }; })
      .sort(function(a, b) { return b.amount - a.amount; });
    return { total: total, byHead: byHead, headWiseAvailable: true };
  }
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { total: 0, byHead: [], headWiseAvailable: false };
  var dueIdx = values[0].indexOf('Due');
  if (dueIdx === -1) return { total: 0, byHead: [], headWiseAvailable: false };
  var total2 = 0;
  values.slice(1).forEach(function(row) {
    var d = safeNum(row[dueIdx]);
    if (d > 0) total2 += d; // positive-only, same rule as everywhere else — see sumFeeRows_ / getOutstandingList_
  });
  // Archived year — no per-head split was ever saved for it.
  return { total: total2, byHead: [], headWiseAvailable: false };
}

// Same shape as getStudentList(), but for a past academic year that's
// already been archived (Settings → Promote Students) it reads that
// year's frozen Students_<year> snapshot instead of live data — so Fee
// Management / Outstanding can show correct figures for a closed year
// instead of always showing today's live numbers. Per-fee-head
// breakdown isn't available for archived years (the archive only stores
// each student's totals), so `heads` comes back empty and the UI treats
// the result as read-only.
function getStudentListForYear(token, year) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  if (!year || year === getActiveAcademicYear_()) {
    var live = getStudentList(token);
    if (live.error) return live;
    live.year = getActiveAcademicYear_();
    live.readOnly = false;
    return live;
  }

  var sh = _getSS_().getSheetByName("Students_" + year);
  if (!sh) return { error: "No archived data found for " + year + " — that year was never promoted/archived." };
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { success: true, data: [], readOnly: true, year: year };
  var col = {};
  values[0].forEach(function(h, i) { col[h] = i; });
  var data = values.slice(1).filter(function(row) { return row[col['Student ID']]; }).map(function(row) {
    return {
      studentId: row[col['Student ID']], name: row[col['Name']], class: row[col['Class']],
      section: row[col['Section']], transport: row[col['Transport']], gender: row[col['Gender']],
      status: row[col['Status']], photo: '',
      billed: safeNum(row[col['Billed']]), paid: safeNum(row[col['Paid']]), due: safeNum(row[col['Due']]),
      heads: {}
    };
  });
  return { success: true, data: data, readOnly: true, year: year };
}

function getReceiptPrefix_() {
  if (RECEIPT_PREFIX) return RECEIPT_PREFIX;
  var initials = SCHOOL_NAME.split(/\s+/).filter(Boolean).map(function(w) {
    return w.charAt(0);
  }).join("").toUpperCase();
  return initials || "REC";
}

// Maps fee head names to short 3-letter codes for receipt number prefixes.
// Falls back to the first 3 letters of the head name if not mapped.
function getFeeHeadCode_(headName) {
  var MAP = {
    "Admission Fee": "ADM", "Admission": "ADM",
    "Tuition Fee": "TUI", "Tution Fee": "TUI", "Tuition": "TUI",
    "Bus Fee": "BUS", "Transport Fee": "TRP",
    "Hostel Fee": "HOS", "Hostel": "HOS",
    "Books Fee": "BOK", "Books": "BOK",
    "Uniform Fee": "UNI", "Uniform": "UNI",
    "Badge Fee": "BDG", "Badge": "BDG",
    "Exam Fee": "EXM", "Exam": "EXM",
    "Other": "OTH", "Others": "OTH",
    "Old Fee": "OLD", "Old Due": "OLD"
  };
  if (MAP[headName]) return MAP[headName];
  // Auto-derive 3-letter code from first 3 letters of the head name
  return (headName || "REC").replace(/[^A-Za-z]/g, "").substring(0, 3).toUpperCase() || "REC";
}

// Scans the Transactions sheet for the highest sequence number already used
// for a specific fee head within the given academic year, and returns the
// next one — continuing the numbering per head rather than resetting.
// Format: SCHOOLINITIALS/HEADCODE/AY/SEQ e.g. AHS/ADM/26-27/0001
function getNextReceiptNo_(forDate, feeHead) {
  var ay      = getAcademicYear_(forDate);
  var school  = getReceiptPrefix_();
  var headCode = getFeeHeadCode_(feeHead || "");
  var prefix   = school + "/" + headCode + "/" + ay + "/";

  var txSh  = getSheet(SHEETS.TRANSACTIONS);
  var data  = txSh.getDataRange().getValues();
  var maxSeq = 0;
  for (var i = 1; i < data.length; i++) {
    var rec = safeStr(data[i][TX_COL.RECEIPT]);
    if (rec.indexOf(prefix) === 0) {
      var seq = parseInt(rec.substring(prefix.length), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  var next = maxSeq + 1;
  return prefix + ("0000" + next).slice(-4);
}

// Lets the client preview the receipt number before saving (e.g. as soon as
// the payment modal opens, or whenever the chosen date changes). The actual
// save still recomputes this under a lock, so it's always accurate even if
// two Clerk members are recording payments at the same moment.
function getNextReceiptNo(token, dateStr, feeHead) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  return { success: true, receiptNo: getNextReceiptNo_(parsePaymentDate_(dateStr), feeHead || "") };
}

// Scans Student Profile for the highest STU#### sequence already used and
// returns the next one. New in v3 — see Notes item #7.
function getNextStudentId_() {
  var sh = getSheet(SHEETS.PROFILE);
  var data = sh.getDataRange().getValues();
  var maxSeq = 0;
  for (var i = 2; i < data.length; i++) {
    var id = safeStr(data[i][SP_COL.STUDENT_ID]);
    if (id.indexOf(STUDENT_ID_PREFIX) === 0) {
      var seq = parseInt(id.substring(STUDENT_ID_PREFIX.length), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  return STUDENT_ID_PREFIX + ("0000" + (maxSeq + 1)).slice(-4);
}

// ============================================================
//  SETTINGS — single configurable source for Classes+Sections, Fee
//  Heads, Expense Categories, and Bank Accounts. Everything the school
//  might need to add, rename, or retire (a new class, a new section, a
//  new fee head, a new bank account) is edited from inside the app and
//  takes effect immediately — no code change required.
// ============================================================

// Ensures the Settings sheet exists and has its seed rows. Runs once —
// after the first row is written, this never overwrites existing rows,
// so anything the school has already added/edited/deactivated is safe.
// Creates a human-readable "User Roles & Permissions" sheet the first
// time the app runs, so anyone opening the spreadsheet can see exactly
// what each login (Admin / Principal / Clerk) is allowed to do. It is
// created ONCE and never overwritten afterwards, so any notes an admin
// adds to it are preserved. This sheet documents the same rules the code
// enforces in requireAdmin_() / requireManagement_() and the individual
// function checks.
function ensureRolesSheet_() {
  var ss = _getSS_();
  var sh = ss.getSheetByName(SHEETS.ROLES);
  if (sh) return; // already created — never overwrite
  sh = ss.insertSheet(SHEETS.ROLES);

  var Y = "Yes", N = "No";
  var rows = [
    ["USER ROLES & PERMISSIONS", "", "", ""],
    ["School DBMS \u00b7 Manha E-Solutions \u2014 who can do what", "", "", ""],
    ["", "", "", ""],
    ["Capability", "Admin", "Principal", "Clerk"],
    ["View dashboard & all reports", Y, Y, Y],
    ["Record fee receipts / payments", Y, Y, Y],
    ["Record expenses (incl. Cash Deposit / Withdrawal / Bank-to-Bank transfers)", Y, Y, Y],
    ["Enter marks", Y, Y, Y],
    ["Mark staff & student attendance", Y, Y, Y],
    ["Print receipts, statements & reports", Y, Y, Y],
    ["Add a student (single)", Y, Y, Y],
    ["Bulk add students (import)", Y, Y, N],
    ["Add a teacher", Y, Y, N],
    ["Edit / modify student details", Y, N, N],
    ["Edit student fees", Y, N, N],
    ["Edit / modify teacher details", Y, N, N],
    ["Remove / delete a student", Y, N, N],
    ["Cancel a fee receipt / payment", Y, N, N],
    ["Cancel an expense voucher", Y, N, N],
    ["Add / edit / hide fee heads (income heads)", Y, N, N],
    ["Add / edit / hide expense heads (incl. mark as Cash\u2194Bank transfer)", Y, N, N],
    ["Add / edit bank accounts", Y, N, N],
    ["Set / edit opening balances", Y, N, N],
    ["Manage classes, sections, exams, subjects", Y, N, N],
    ["Change tab visibility & payment-mode visibility", Y, N, N],
    ["Edit School Identity (letterhead)", Y, N, N],
    ["", "", "", ""],
    ["In short:", "", "", ""],
    ["\u2022 Clerk  \u2014 enter data (receipts, payments, expenses, marks, attendance) and view reports.", "", "", ""],
    ["\u2022 Principal  \u2014 everything Clerk can do, plus add / bulk-add students & teachers.", "", "", ""],
    ["\u2022 Admin  \u2014 full control: all of the above plus every edit, cancellation, deletion and Settings change.", "", "", ""],
    ["", "", "", ""],
    ["Note: This sheet is documentation only. The app enforces these rules in code.", "", "", ""]
  ];
  sh.getRange(1, 1, rows.length, 4).setValues(rows);

  // Formatting: title, header row, column widths.
  sh.getRange(1, 1, 1, 4).merge().setFontSize(15).setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1F3A5F").setHorizontalAlignment("center");
  sh.getRange(2, 1, 1, 4).merge().setFontStyle("italic").setFontColor("#555555").setHorizontalAlignment("center");
  sh.getRange(4, 1, 1, 4).setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#2C5F86");
  sh.setColumnWidth(1, 430);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 95);
  sh.setColumnWidth(4, 90);
  sh.getRange(5, 2, 22, 3).setHorizontalAlignment("center");
  // Colour the Yes/No cells for quick scanning.
  var permRange = sh.getRange(5, 2, 22, 3);
  var vals = permRange.getValues();
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < 3; c++) {
      var cell = sh.getRange(5 + r, 2 + c);
      if (vals[r][c] === "Yes") cell.setBackground("#E6F4EA").setFontColor("#137333");
      else if (vals[r][c] === "No") cell.setBackground("#FCE8E6").setFontColor("#C5221F");
    }
  }
  sh.setFrozenRows(4);
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setVerticalAlignment("middle"); } catch (e) {}
}

// Removes duplicate Settings rows of the same Type+Key, keeping only the
// first occurrence (lowest sheet row). Rows are deleted from the bottom up
// so indices don't shift during deletion. This is safe to call repeatedly —
// it is a no-op when no duplicates exist.
function deduplicateSettingsRows_(sh, type) {
  if (!sh || sh.getLastRow() < 2) return;
  var data = sh.getDataRange().getValues();
  var seen = {};
  var toDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) !== type) continue;
    var key = safeStr(data[i][SET_COL.KEY]);
    if (seen[key]) {
      toDelete.push(i + 1); // 1-based sheet row
    } else {
      seen[key] = true;
    }
  }
  toDelete.sort(function(a, b) { return b - a; }); // delete bottom-up
  toDelete.forEach(function(r) { sh.deleteRow(r); });
  if (toDelete.length) SpreadsheetApp.flush();
}


function ensureSettingsSeeded_() {
  // Skip on subsequent calls within the same GAS execution — the migration
  // and seed checks each do a getDataRange().getValues() which is expensive.
  if (_settingsSeeded_) return;
  _settingsSeeded_ = true;
  ensureRolesSheet_();
  var sh = getSheet(SHEETS.SETTINGS);
  if (sh.getLastRow() > 0) {
    migrateSettingsAddValue5Column_(sh);
    seedMissingSettingsTypes_(sh); // adds Exam/ClassSubjects rows to an already-seeded sheet
    // One-time self-heal for the column-drift/duplicate-row corruption
    // described above repairSettingsSheet_(). Gated by a Script Property
    // so it only ever runs once per install, not on every request.
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty("SETTINGS_SCHEMA_REPAIRED_V2") !== "1") {
      repairSettingsSheet_();
      props.setProperty("SETTINGS_SCHEMA_REPAIRED_V2", "1");
    }
    return; // already seeded (or has real data)
  }

  sh.appendRow(["Type", "Key", "Value1", "Value2", "Value3", "Value4", "Value5", "Active"]);
  var rows = [];
  SEED_CLASSES.forEach(function(c) {
    rows.push([SETTINGS_TYPE.CLASS, c.name, c.sections.join(","), "", "", "", "", true]);
  });
  SEED_FEE_HEADS.forEach(function(h) {
    rows.push([SETTINGS_TYPE.FEE_HEAD, h, "", "", "", "", "", true]);
  });
  SEED_EXPENSE_CATEGORIES.forEach(function(c) {
    rows.push([SETTINGS_TYPE.EXPENSE_CATEGORY, c.name, c.isTransfer ? "TRUE" : "FALSE", "", "", "", "", true]);
  });
  SEED_BANK_ACCOUNTS.forEach(function(a) {
    rows.push([SETTINGS_TYPE.BANK_ACCOUNT, a.label, a.accNo, a.bankName, a.ifsc, a.branch, "", true]);
  });
  SEED_EXAMS.forEach(function(e) {
    rows.push([SETTINGS_TYPE.EXAM, e.name, e.maxMarks, e.passMarks, e.group, "", "", true]);
  });
  SEED_CLASS_SUBJECTS.forEach(function(cs) {
    rows.push([SETTINGS_TYPE.CLASS_SUBJECTS, cs.className, cs.subjects.join(","), "", "", "", "", true]);
  });
  SEED_SUBJECTS.forEach(function(subj) {
    rows.push([SETTINGS_TYPE.SUBJECT, subj, "", "", "", "", "", true]);
  });
  if (rows.length) sh.getRange(2, 1, rows.length, 8).setValues(rows);
}

// Adds Exam and ClassSubjects seed rows to a Settings sheet that was
// already seeded BEFORE these types existed (i.e. every install created
// before this feature was added). Runs on every call but is a no-op
// once these types exist, so it's safe to call unconditionally.
function seedMissingSettingsTypes_(sh) {
  var data = sh.getDataRange().getValues();
  var hasExam = false, hasClassSubjects = false, hasSubject = false, hasBankToBankTransfer = false, hasBankCharges = false;
  for (var i = 1; i < data.length; i++) {
    var t = safeStr(data[i][SET_COL.TYPE]);
    if (t === SETTINGS_TYPE.EXAM) hasExam = true;
    if (t === SETTINGS_TYPE.CLASS_SUBJECTS) hasClassSubjects = true;
    if (t === SETTINGS_TYPE.SUBJECT) hasSubject = true;
    if (t === SETTINGS_TYPE.EXPENSE_CATEGORY && safeStr(data[i][SET_COL.KEY]) === "Bank to Bank Transfer") hasBankToBankTransfer = true;
    if (t === SETTINGS_TYPE.EXPENSE_CATEGORY && safeStr(data[i][SET_COL.KEY]) === "Bank Charges") hasBankCharges = true;
  }
  var newRows = [];
  if (!hasExam) {
    SEED_EXAMS.forEach(function(e) {
      newRows.push([SETTINGS_TYPE.EXAM, e.name, e.maxMarks, e.passMarks, e.group, "", "", true]);
    });
  }
  if (!hasClassSubjects) {
    SEED_CLASS_SUBJECTS.forEach(function(cs) {
      newRows.push([SETTINGS_TYPE.CLASS_SUBJECTS, cs.className, cs.subjects.join(","), "", "", "", "", true]);
    });
  }
  if (!hasSubject) {
    SEED_SUBJECTS.forEach(function(subj) {
      newRows.push([SETTINGS_TYPE.SUBJECT, subj, "", "", "", "", "", true]);
    });
  }
  // Notes request: "Add Bank to Bank transfer option" — schools that
  // already had Expense Categories seeded before this feature existed
  // need this added here explicitly, since the normal one-time category
  // seeding only ever runs once on a brand-new Settings sheet.
  if (!hasBankToBankTransfer) {
    newRows.push([SETTINGS_TYPE.EXPENSE_CATEGORY, "Bank to Bank Transfer", "TRUE", "", "", "", "", true]);
  }
  // Seed "Bank Charges" for existing installs — this is a real expense
  // (isTransfer: false) so it flows into P&L Expenditure and appears as
  // its own line, separate from Bank-to-Bank transfers.
  if (!hasBankCharges) {
    newRows.push([SETTINGS_TYPE.EXPENSE_CATEGORY, "Bank Charges", "FALSE", "", "", "", "", true]);
  }
  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
  }
  // One-time dedup: if earlier runs accidentally created multiple rows for the
  // same expense category (e.g. "Bank to Bank Transfer"), remove the duplicates
  // now, keeping only the first occurrence of each key.
  deduplicateSettingsRows_(sh, SETTINGS_TYPE.EXPENSE_CATEGORY);
}

// One-time, idempotent upgrade for a Settings sheet created before the
// Value5 (photo/logo) column existed. Detects the old 7-column header
// ("Type","Key","Value1..4","Active" with nothing in between) and
// inserts a blank "Value5" column before "Active" on every row, so
// existing Classes/FeeHeads/ExpenseCategories/BankAccounts data shifts
// into the new layout instead of being silently misread.
function migrateSettingsAddValue5Column_(sh) {
  var headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var hasValue5 = headerRow.indexOf("Value5") !== -1;
  if (hasValue5) return; // already migrated
  if (sh.getLastColumn() < 7) return; // not the shape we expect — leave alone rather than guess

  // Insert a new column before the current "Active" column (old layout:
  // column 7 / index 6), then label it, so every row's Active value
  // shifts right by one into its new position automatically.
  sh.insertColumnBefore(7);
  sh.getRange(1, 7).setValue("Value5");
}

// ── ONE-TIME SETTINGS SHEET REPAIR ──────────────────────────
// Over time, several appendRow() calls in this file (addClass, addFeeHead,
// addExpenseCategory, addBankAccount — now fixed above) wrote one column
// too few. Each time that happened, the trailing "Active" boolean landed
// one column early. Combined with the later migrateSettingsAddValue5Column_
// insert (which shifts everything AFTER it one column to the right), real
// installs ended up with extra stray columns (duplicate "Value5" headers,
// blank ghost columns) and the true Active flag scattered across different
// columns on different rows depending on when each row was added. This
// also produced duplicate rows for the same Type+Key (e.g. "Bank to Bank
// Transfer", "Subject" entries appearing twice).
//
// This rebuilds the ENTIRE Settings sheet from scratch into the canonical
// 8-column layout (Type, Key, Value1, Value2, Value3, Value4, Value5,
// Active), recovering the intended Active flag with a heuristic (the
// right-most TRUE/FALSE-looking cell in each row — every appendRow call
// always wrote the Active flag as its last argument, so it's always the
// right-most boolean-shaped value in the row no matter which column it
// drifted into), and removes duplicate rows (same Type+Key), preferring
// to keep an ACTIVE copy over an inactive one when both exist.
//
// Safe to call repeatedly — running it twice on an already-clean sheet
// is a no-op. Call once from the Apps Script editor (run repairSettingsSheetNow
// with no token check needed there), or via the exposed admin-only wrapper.
function repairSettingsSheet_() {
  var sh = getSheet(SHEETS.SETTINGS);
  if (sh.getLastRow() < 2) return { changed: false, message: "Settings sheet is empty — nothing to repair." };

  var data = sh.getDataRange().getValues();
  var header = data[0];
  var canonicalHeader = ["Type", "Key", "Value1", "Value2", "Value3", "Value4", "Value5", "Active"];

  function looksBoolish(v) {
    return v === true || v === false || safeStr(v).toUpperCase() === "TRUE" || safeStr(v).toUpperCase() === "FALSE";
  }

  var cleanRows = [];     // [type, key, v1, v2, v3, v4, v5, active]
  var bestByKey = {};     // "Type||Key" -> index into cleanRows of the best copy kept so far

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var type = safeStr(row[0]);
    var key  = safeStr(row[1]);
    if (!type || !key) continue; // skip genuinely blank rows

    var v1 = row[2] !== undefined ? row[2] : "";
    var v2 = row[3] !== undefined ? row[3] : "";
    var v3 = row[4] !== undefined ? row[4] : "";
    var v4 = row[5] !== undefined ? row[5] : "";
    // Value5 only ever meant anything for SchoolInfo (the logo URL/Drive
    // link). For every other type it's unused, so drop whatever garbage
    // drifted in there for those rows; for SchoolInfo, the logo lives at
    // the original index 6 — keep it if that cell isn't itself a stray
    // boolean from the drift.
    var v5 = "";
    if (type === SETTINGS_TYPE.SCHOOL_INFO && row[6] !== undefined && !looksBoolish(row[6])) {
      v5 = row[6];
    }

    // Recover the intended Active flag: scan the row right-to-left for the
    // right-most boolean-shaped cell (TRUE/FALSE or real Boolean) — that is
    // always where the originally-appended Active value ended up, however
    // far it drifted to the right due to column-count bugs/migrations.
    var active = true; // default — matches the rest of the codebase's "blank == active" behavior
    for (var c = row.length - 1; c >= 6; c--) {
      if (looksBoolish(row[c])) { active = isTrueVal_(row[c]); break; }
    }

    var dedupeKey = type + "||" + key;
    if (bestByKey.hasOwnProperty(dedupeKey)) {
      // Duplicate Type+Key — prefer keeping the ACTIVE copy if there's a
      // conflict, otherwise keep the first copy already recorded.
      var existingIdx = bestByKey[dedupeKey];
      if (!cleanRows[existingIdx][7] && active) {
        cleanRows[existingIdx] = [type, key, v1, v2, v3, v4, v5, active];
      }
      continue;
    }
    bestByKey[dedupeKey] = cleanRows.length;
    cleanRows.push([type, key, v1, v2, v3, v4, v5, active]);
  }

  var headerChanged = JSON.stringify(header.slice(0, 8)) !== JSON.stringify(canonicalHeader) || sh.getLastColumn() !== 8;
  var rowsRemoved = (data.length - 1) - cleanRows.length;
  if (!headerChanged && rowsRemoved === 0) {
    return { changed: false, message: "Settings sheet is already clean — no repair needed." };
  }

  // Rebuild: clear the whole sheet and write back the canonical 8 columns.
  sh.clear();
  sh.getRange(1, 1, 1, 8).setValues([canonicalHeader]);
  if (cleanRows.length) {
    sh.getRange(2, 1, cleanRows.length, 8).setValues(cleanRows);
  }
  SpreadsheetApp.flush();

  return {
    changed: true,
    message: "Settings sheet repaired: rebuilt to 8 clean columns, removed " + rowsRemoved + " duplicate row(s)."
  };
}

// Admin-only wrapper so this can be triggered from the web app (e.g. a
// "Repair Settings" button) instead of needing the Apps Script editor.
function repairSettingsSheetNow(token) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var result = repairSettingsSheet_();
  return { success: true, message: result.message };
}


function getSettingsRows_(type) {
  ensureSettingsSeeded_();
  // Cache the raw Settings data for the lifetime of this GAS execution.
  // Every call after the first returns the in-memory array instantly.
  if (!_cachedSettings_) {
    _cachedSettings_ = getSheet(SHEETS.SETTINGS).getDataRange().getValues();
  }
  var data = _cachedSettings_;
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (safeStr(r[SET_COL.TYPE]) !== type) continue;
    if (r[SET_COL.ACTIVE] === false || safeStr(r[SET_COL.ACTIVE]).toUpperCase() === "FALSE") continue;
    out.push({ rowIndex: i, key: safeStr(r[SET_COL.KEY]), v1: safeStr(r[SET_COL.VALUE1]),
               v2: safeStr(r[SET_COL.VALUE2]), v3: safeStr(r[SET_COL.VALUE3]), v4: safeStr(r[SET_COL.VALUE4]),
               v5: safeStr(r[SET_COL.VALUE5]) });
  }
  return out;
}

// ── FEE HEADS (for dropdown) ────────────────────────────────
function getFeeHeads(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.FEE_HEAD);
  return { success: true, data: rows.map(function(r){ return r.key; }) };
}

// Independent class list for filter dropdowns. New in v3 — Notes item #9.
// Fee Management's Class filter previously only got populated as a side
// effect of loading Student Records first; this call needs no other
// page to have loaded first, so it works regardless of navigation order.
// v4: classes (and their section lists) come from Settings, not a fixed array.
function getClassList(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.CLASS);
  return { success: true, data: rows.map(function(r){ return r.key; }) };
}

// Returns the section list configured for one specific class (e.g.
// "5" -> ["A","B","C"]), used to populate the Section dropdown once a
// class is picked in the Add/Edit Student form.
function getSectionsForClass(token, className) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.CLASS);
  var row = rows.filter(function(r){ return r.key === className; })[0];
  var sections = row && row.v1 ? row.v1.split(",").map(function(s){ return s.trim(); }).filter(Boolean) : [];
  return { success: true, data: sections };
}

// Full Classes+Sections list in one call, for the Settings screen itself.
function getClassesWithSections(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.CLASS);
  return { success: true, data: rows.map(function(r) {
    return { rowIndex: r.rowIndex, name: r.key, sections: r.v1 ? r.v1.split(",").map(function(s){return s.trim();}).filter(Boolean) : [] };
  }) };
}

// Adds a new class with its initial section list (comma-separated string
// or array both accepted from the client).
function addClass(token, className, sections) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!className) return { error: "Class name is required." };

  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.CLASS && safeStr(data[i][SET_COL.KEY]) === className) {
      return { error: "A class named \"" + className + "\" already exists." };
    }
  }
  var sectionStr = Array.isArray(sections) ? sections.join(",") : safeStr(sections);
  sh.appendRow([SETTINGS_TYPE.CLASS, className, sectionStr, "", "", "", "", true]);
  return { success: true, message: "Class \"" + className + "\" added." };
}

// Renames a class and/or replaces its full section list.
function updateClass(token, rowIndex, className, sections) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };

  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Class not found." };
  var sectionStr = Array.isArray(sections) ? sections.join(",") : safeStr(sections);
  sh.getRange(sheetRow, SET_COL.KEY + 1).setValue(className);
  sh.getRange(sheetRow, SET_COL.VALUE1 + 1).setValue(sectionStr);
  return { success: true, message: "Class updated." };
}

// Retires a class (soft-delete via Active=false) rather than removing the
// row outright, so historical student/fee records that reference it are
// never orphaned.
function deactivateClass(token, rowIndex) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Class not found." };
  sh.getRange(sheetRow, SET_COL.ACTIVE + 1).setValue(false);
  return { success: true, message: "Class removed from active lists." };
}

// ── FEE HEAD SETTINGS CRUD ──────────────────────────────────
function addFeeHead(token, name) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!name) return { error: "Fee head name is required." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.FEE_HEAD && safeStr(data[i][SET_COL.KEY]) === name) {
      return { error: "Fee head \"" + name + "\" already exists." };
    }
  }
  sh.appendRow([SETTINGS_TYPE.FEE_HEAD, name, "", "", "", "", "", true]);
  return { success: true, message: "Fee head \"" + name + "\" added." };
}

// Generic lookup used by the Settings UI when a row's index isn't
// already in hand (e.g. the simple Fee Heads list only shows names).
// Returns -1 if not found.
function findSettingsRowIndex(token, type, key) {
  if (!validateSession(token)) return -1;
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === type && safeStr(data[i][SET_COL.KEY]) === key) return i;
  }
  return -1;
}

// Full fee head list including inactive ones, with rowIndex + active
// state — used by the Settings screen's toggle UI. (getFeeHeads() above
// stays active-only, since that's what every form/chart should use.)
function getFeeHeadsFull(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) !== SETTINGS_TYPE.FEE_HEAD) continue;
    var activeVal = data[i][SET_COL.ACTIVE];
    var isActive = !(activeVal === false || safeStr(activeVal).toUpperCase() === "FALSE");
    out.push({ rowIndex: i, name: safeStr(data[i][SET_COL.KEY]), active: isActive });
  }
  return { success: true, data: out };
}

// Toggles a fee head on/off — Notes request: "whenever we select other
// heads also in Settings, they should become visible" while entering
// fees. Replaces the old one-way deactivateFeeHead() with a real toggle.
function setFeeHeadActive(token, rowIndex, active) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Fee head not found." };
  sh.getRange(sheetRow, SET_COL.ACTIVE + 1).setValue(!!active);
  return { success: true, message: active ? "Fee head is now visible." : "Fee head hidden from forms and charts." };
}

// Generic cascade-rename helper: rewrites every cell equal to oldVal in
// one column of a sheet to newVal, returning how many rows changed. Used
// when an admin renames a fee head or expense category so existing
// billing / receipt / voucher rows keep matching the renamed setting
// instead of being orphaned under the old name.
function renameInColumn_(sheetName, colIndex0, oldVal, newVal) {
  var sh = getSheet(sheetName);
  if (sh.getLastRow() < 2) return 0;
  var rng = sh.getRange(2, colIndex0 + 1, sh.getLastRow() - 1, 1);
  var vals = rng.getValues();
  var count = 0;
  for (var i = 0; i < vals.length; i++) {
    if (safeStr(vals[i][0]) === oldVal) { vals[i][0] = newVal; count++; }
  }
  if (count) rng.setValues(vals);
  return count;
}

// Renames (edits) a fee head — Notes request: an Edit option for income
// heads in Settings. Cascades the new name into Tution_Fee and
// Transactions so historical billing and receipts stay linked.
function updateFeeHead(token, rowIndex, newName) {
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  newName = safeStr(newName).trim();
  if (!newName) return { error: "Fee head name is required." };

  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Fee head not found." };

  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (i === Number(rowIndex)) continue;
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.FEE_HEAD && safeStr(data[i][SET_COL.KEY]) === newName) {
      return { error: "A fee head named \"" + newName + "\" already exists." };
    }
  }
  var oldName = safeStr(sh.getRange(sheetRow, SET_COL.KEY + 1).getValue());
  if (oldName === newName) return { success: true, message: "No change \u2014 the name is the same." };

  sh.getRange(sheetRow, SET_COL.KEY + 1).setValue(newName);
  var renamed = renameInColumn_(SHEETS.TUITION, TF_COL.FEE_HEAD, oldName, newName);
  renamed += renameInColumn_(SHEETS.TRANSACTIONS, TX_COL.HEAD, oldName, newName);
  SpreadsheetApp.flush();
  return { success: true, message: "Fee head renamed to \"" + newName + "\"" + (renamed ? " (" + renamed + " existing record(s) updated)." : ".") };
}

// ── EXPENSE CATEGORY SETTINGS CRUD ──────────────────────────
function getExpenseCategories(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.EXPENSE_CATEGORY);
  var transferCats = rows.filter(function(r){ return isTrueVal_(r.v1); }).map(function(r){ return r.key; });
  return { success: true, data: rows.map(function(r){ return r.key; }), transferCategories: transferCats };
}

// Full category list with rowIndex + isTransfer flag, for the Settings screen.
function getExpenseCategoriesFull(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  // Include INACTIVE rows too so the Settings screen can show them with a
  // Reactivate option — this prevents the confusing "already exists" error
  // when a user tries to add a category that was previously deactivated.
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) !== SETTINGS_TYPE.EXPENSE_CATEGORY) continue;
    var activeVal = data[i][SET_COL.ACTIVE];
    var isActive = !(activeVal === false || safeStr(activeVal).toUpperCase() === "FALSE");
    out.push({ rowIndex: i, name: safeStr(data[i][SET_COL.KEY]), isTransfer: isTrueVal_(data[i][SET_COL.VALUE1]), active: isActive });
  }
  return { success: true, data: out };
}

function addExpenseCategory(token, name, isTransfer) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!name) return { error: "Category name is required." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.EXPENSE_CATEGORY && safeStr(data[i][SET_COL.KEY]) === name) {
      return { error: "Category \"" + name + "\" already exists." };
    }
  }
  sh.appendRow([SETTINGS_TYPE.EXPENSE_CATEGORY, name, isTransfer ? "TRUE" : "FALSE", "", "", "", "", true]);
  return { success: true, message: "Category \"" + name + "\" added." };
}

function deactivateExpenseCategory(token, rowIndex) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Category not found." };
  sh.getRange(sheetRow, SET_COL.ACTIVE + 1).setValue(false);
  return { success: true, message: "Category removed from active lists." };
}

// Reactivates a previously deactivated expense category — the Settings UI
// now shows inactive rows with a "Reactivate" button so users can restore
// them instead of getting a confusing "already exists" error when trying to add
// a category that was removed earlier (e.g. "Bank to Bank Transfer").
function setExpenseCategoryActive(token, rowIndex, active) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Category not found." };
  if (safeStr(sh.getRange(sheetRow, SET_COL.TYPE + 1).getValue()) !== SETTINGS_TYPE.EXPENSE_CATEGORY) {
    return { error: "Row is not an expense category." };
  }
  sh.getRange(sheetRow, SET_COL.ACTIVE + 1).setValue(!!active);
  var name = safeStr(sh.getRange(sheetRow, SET_COL.KEY + 1).getValue());
  SpreadsheetApp.flush();
  return { success: true, message: "\"" + name + "\" " + (active ? "is now visible in expense forms." : "removed from active lists.") };
}

// The three built-in transfer directions the Cash/Bank reconciliation
// recognises by their EXACT names. Their isTransfer flag can be toggled,
// but their names must not change or the reconciliation would no longer
// know which direction the money moved.
var PROTECTED_TRANSFER_CATEGORIES = ["Cash Deposit", "Cash Withdrawal", "Bank to Bank Transfer"];

// Renames (edits) an expense head and/or toggles its "Cash\u2194Bank
// transfer" (contra) flag — Notes request: an Edit option for expense
// heads, and a checkbox that marks an entry as an internal cash/bank
// exchange (deposit, withdrawal, bank-to-bank) that is NOT counted as an
// expense in the Profit & Loss / Income & Expenditure statement.
function updateExpenseCategory(token, rowIndex, newName, isTransfer) {
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  newName = safeStr(newName).trim();
  if (!newName) return { error: "Category name is required." };

  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Category not found." };

  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (i === Number(rowIndex)) continue;
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.EXPENSE_CATEGORY && safeStr(data[i][SET_COL.KEY]) === newName) {
      return { error: "A category named \"" + newName + "\" already exists." };
    }
  }
  var oldName = safeStr(sh.getRange(sheetRow, SET_COL.KEY + 1).getValue());

  if (PROTECTED_TRANSFER_CATEGORIES.indexOf(oldName) !== -1 && newName !== oldName) {
    return { error: "\"" + oldName + "\" is a built-in cash/bank transfer type used by the reconciliation and cannot be renamed. You can rename your own categories, or add new ones." };
  }
  // The three built-in directions are always treated as transfers by the
  // reconciliation (matched by name), so their flag is locked ON to keep
  // the Settings display honest about how they behave.
  if (PROTECTED_TRANSFER_CATEGORIES.indexOf(oldName) !== -1) isTransfer = true;

  sh.getRange(sheetRow, SET_COL.KEY + 1).setValue(newName);
  sh.getRange(sheetRow, SET_COL.VALUE1 + 1).setValue(isTransfer ? "TRUE" : "FALSE");
  var renamed = 0;
  if (oldName !== newName) renamed = renameInColumn_(SHEETS.EXPENSES, EXP_COL.CATEGORY, oldName, newName);
  SpreadsheetApp.flush();

  var typeWord = isTransfer ? "Cash\u2194Bank transfer (not counted as an expense)" : "Expenditure";
  return { success: true, message: "Saved \"" + newName + "\" as " + typeWord + (renamed ? " (" + renamed + " existing voucher(s) updated)." : ".") };
}

// ── BANK ACCOUNT SETTINGS CRUD ──────────────────────────────
function getBankAccounts(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.BANK_ACCOUNT);
  return { success: true, data: rows.map(function(r) {
    return { rowIndex: r.rowIndex, label: r.key, accNo: r.v1, bankName: r.v2, ifsc: r.v3, branch: r.v4 };
  }) };
}

function addBankAccount(token, obj) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!obj || !obj.label) return { error: "Account label is required." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  sh.appendRow([SETTINGS_TYPE.BANK_ACCOUNT, obj.label, obj.accNo || "", obj.bankName || "", obj.ifsc || "", obj.branch || "", "", true]);
  return { success: true, message: "Bank account \"" + obj.label + "\" added." };
}

function updateBankAccount(token, rowIndex, obj) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Bank account not found." };
  sh.getRange(sheetRow, SET_COL.KEY + 1, 1, 5).setValues([[obj.label || "", obj.accNo || "", obj.bankName || "", obj.ifsc || "", obj.branch || ""]]);
  return { success: true, message: "Bank account updated." };
}

function deactivateBankAccount(token, rowIndex) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Bank account not found." };
  sh.getRange(sheetRow, SET_COL.ACTIVE + 1).setValue(false);
  return { success: true, message: "Bank account removed from active lists." };
}

// ── OPENING BALANCES (Cash + each Bank Account) ──────────────
// Notes request: a way to enter the REAL starting balance (with a date)
// for cash-in-hand and for each bank account, so the app's running
// Cash-in-Hand / Bank Balance figures can actually be made to match the
// real bank passbook / cash box — not just track movement from zero.
var OPENING_BALANCE_CASH_KEY = "Cash";

function getOpeningBalances(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) !== SETTINGS_TYPE.OPENING_BALANCE) continue;
    var key = safeStr(data[i][SET_COL.KEY]);
    out[key] = {
      amount: safeNum(data[i][SET_COL.VALUE1]),
      date: formatDateSafe_(data[i][SET_COL.VALUE2]) || "",
      rowIndex: i
    };
  }
  return { success: true, data: out };
}

// `accountKey` is "Cash" or a bank account's label (must match a Bank
// Account already configured in Settings, or "Cash" exactly).
function setOpeningBalance(token, accountKey, amount, dateStr) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!accountKey) return { error: "Account is required." };
  if (!dateStr) return { error: "Opening date is required." };

  var openingDate = parsePaymentDate_(dateStr);
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.OPENING_BALANCE && safeStr(data[i][SET_COL.KEY]) === accountKey) {
      sh.getRange(i + 1, SET_COL.VALUE1 + 1, 1, 2).setValues([[safeNum(amount), openingDate]]);
      SpreadsheetApp.flush();
      return { success: true, message: "Opening balance updated for " + accountKey + "." };
    }
  }
  sh.appendRow([SETTINGS_TYPE.OPENING_BALANCE, accountKey, safeNum(amount), openingDate, "", "", "", true]);
  SpreadsheetApp.flush();
  return { success: true, message: "Opening balance set for " + accountKey + "." };
}

// ── CASH / BANK RECONCILIATION ───────────────────────────────
// Notes request: a real passbook-style reconciliation so the app's
// figures can be made to match the actual bank account / physical cash
// box. For Cash:
//   Opening + Receipts(Cash) + Cash Withdrawal(bank->cash) - Cash
//   Deposit(cash->bank) - Expenses(Cash mode) = Cash in Hand
// For EACH bank account:
//   Opening + Receipts(that account) + Cash Deposit(into that account)
//   - Cash Withdrawal(from that account) - Expenses(that account) = Closing
// "Receipts" only counts Transactions where mode/bank account actually
// match; cancelled/voided rows are excluded everywhere, matching how
// every other report in this app already treats voided entries.
// Accumulates every Transaction/Expense movement for each account
// within one date window [from, to] (either bound may be null = open-
// ended). Used twice by getCashBankReconciliation(): once for "from the
// stored opening date up to just before the report's start" (to roll
// the opening balance forward) and once for "within the report's own
// date range" (the actual receipts/expenses shown).
function accumulateCashBankMovement_(accounts, from, to) {
  var txSh = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();
  for (var i = 1; i < txData.length; i++) {
    var t = txData[i];
    if (isVoided_(t[TX_COL.VOID])) continue;
    var d = t[TX_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (from && d.getTime() < from.getTime()) continue;
    if (to && d.getTime() > to.getTime()) continue;
    var amt = safeNum(t[TX_COL.AMOUNT]);
    var mode = safeStr(t[TX_COL.MODE]);
    var bankAcc = safeStr(t[TX_COL.BANK_ACCOUNT]);
    if (mode === "Cash") {
      accounts[OPENING_BALANCE_CASH_KEY].receipts += amt;
    } else if (bankAcc && accounts[bankAcc]) {
      accounts[bankAcc].receipts += amt;
    }
  }

  // Build a set of all isTransfer=TRUE category names from Settings so
  // custom "Cash↔Bank transfer" categories are excluded from real expenses
  // in the reconciliation — same as the three built-in transfer names below.
  var _catRowsForRecon = getSettingsRows_(SETTINGS_TYPE.EXPENSE_CATEGORY);
  var _transferCatSetForRecon = {};
  _catRowsForRecon.forEach(function(rc) { if (isTrueVal_(rc.v1)) _transferCatSetForRecon[rc.key] = true; });

  var expSh = getSheet(SHEETS.EXPENSES);
  var expData = expSh.getDataRange().getValues();
  for (var k = 1; k < expData.length; k++) {
    var r = expData[k];
    if (!r[EXP_COL.VOUCHER] || isVoided_(r[EXP_COL.VOID])) continue;
    var ed = r[EXP_COL.DATE];
    if (Object.prototype.toString.call(ed) !== "[object Date]") continue;
    if (from && ed.getTime() < from.getTime()) continue;
    if (to && ed.getTime() > to.getTime()) continue;
    var eAmt = safeNum(r[EXP_COL.AMOUNT]);
    var cat = safeStr(r[EXP_COL.CATEGORY]);
    var eMode = safeStr(r[EXP_COL.MODE]);
    var eBankAcc = safeStr(r[EXP_COL.BANK_ACCOUNT]);
    var eToBankAcc = safeStr(r[EXP_COL.TO_BANK_ACCOUNT]);

    if (cat === "Cash Deposit") {
      // Cash -> that bank account.
      accounts[OPENING_BALANCE_CASH_KEY].cashDepositOut += eAmt;
      if (eBankAcc && accounts[eBankAcc]) accounts[eBankAcc].cashDepositIn += eAmt;
      continue;
    }
    if (cat === "Cash Withdrawal") {
      // That bank account -> Cash.
      accounts[OPENING_BALANCE_CASH_KEY].cashWithdrawalIn += eAmt;
      if (eBankAcc && accounts[eBankAcc]) accounts[eBankAcc].cashWithdrawalOut += eAmt;
      continue;
    }
    if (cat === "Bank to Bank Transfer") {
      // One bank account -> another — Notes request: pick a From account
      // and a To account; the From account's balance decreases, the To
      // account's balance increases, Cash is untouched entirely.
      if (eBankAcc && accounts[eBankAcc]) accounts[eBankAcc].transferOut += eAmt;
      if (eToBankAcc && accounts[eToBankAcc]) accounts[eToBankAcc].transferIn += eAmt;
      continue;
    }
    // A real expense — debits whichever side it was actually paid from.
    // Custom categories flagged isTransfer=TRUE in Settings are contra entries
    // (internal money movements); exclude them from expenses entirely here,
    // same as the three built-in transfer category names above. They don't
    // add a second leg to cash/bank because their movement direction is
    // undefined (unlike Cash Deposit/Withdrawal/Bank-to-Bank which have
    // explicit From/To semantics). They simply pass through without effect.
    if (_transferCatSetForRecon[cat]) continue;
    if (eMode === "Cash") {
      accounts[OPENING_BALANCE_CASH_KEY].expenses += eAmt;
    } else if (eBankAcc && accounts[eBankAcc]) {
      accounts[eBankAcc].expenses += eAmt;
    }
  }

  // ── Loans ledger ── Loan received increases the account it was credited
  // to; a repayment decreases it by principal + interest. (Interest is the
  // only part treated as a P&L expense — see getIncomeExpenditureSummary.)
  var loanSh = _getSS_().getSheetByName(SHEETS.LOANS);
  if (loanSh) {
    var lData = loanSh.getDataRange().getValues();
    for (var L = 1; L < lData.length; L++) {
      var lr = lData[L];
      if (!lr[LOAN_COL.VOUCHER] || isVoided_(lr[LOAN_COL.VOID])) continue;
      var ld = lr[LOAN_COL.DATE];
      if (Object.prototype.toString.call(ld) !== "[object Date]") continue;
      if (from && ld.getTime() < from.getTime()) continue;
      if (to && ld.getTime() > to.getTime()) continue;
      var lp = safeNum(lr[LOAN_COL.PRINCIPAL]);
      var li = safeNum(lr[LOAN_COL.INTEREST]);
      var lAcct = (safeStr(lr[LOAN_COL.MODE]) === "Cash") ? OPENING_BALANCE_CASH_KEY : safeStr(lr[LOAN_COL.BANK_ACCOUNT]);
      if (!accounts[lAcct]) continue;
      if (safeStr(lr[LOAN_COL.TYPE]) === "Received") {
        accounts[lAcct].loanIn += lp;
      } else if (safeStr(lr[LOAN_COL.TYPE]) === "Repayment") {
        accounts[lAcct].loanOutP += lp;
        accounts[lAcct].loanOutI += li;
      }
    }
  }
}

function newCashBankAccumulator_(bankAccounts) {
  // Every account gets the SAME full set of fields, even though Cash
  // will only ever populate cashDepositOut/cashWithdrawalIn (never
  // cashDepositIn/cashWithdrawalOut, which are bank-only directions, and
  // vice versa for banks) — netMovement_() below is shared by both, so
  // every field must exist on every account or summing them produces NaN.
  function blank() {
    return { receipts: 0, cashWithdrawalIn: 0, cashWithdrawalOut: 0, cashDepositIn: 0, cashDepositOut: 0, expenses: 0, transferIn: 0, transferOut: 0, loanIn: 0, loanOutP: 0, loanOutI: 0 };
  }
  var accounts = {};
  accounts[OPENING_BALANCE_CASH_KEY] = blank();
  bankAccounts.forEach(function(label) {
    accounts[label] = blank();
  });
  return accounts;
}

function netMovement_(a) {
  return a.receipts + a.cashWithdrawalIn + a.cashDepositIn + a.transferIn + a.loanIn - a.cashDepositOut - a.cashWithdrawalOut - a.transferOut - a.expenses - a.loanOutP - a.loanOutI;
}

// Notes request: "Opening balance should work according to the period
// selected" — e.g. stored opening is ₹1000 as on 1-Apr; if ₹5000 came in
// on 1-May and ₹3000 went out on 10-May, a report run for 5-May to 15-
// May must show OPENING = ₹6000 (rolled forward through the 1-May
// receipt, which happened before the report window even though it's
// after the stored opening date) and CLOSING = ₹3000 (after the 10-May
// expense, which falls inside the window). The stored Settings value is
// only ever the TRUE starting point on its own date — every other
// period's opening is derived by rolling forward through everything
// that happened between the stored date and the period's own start.
function getCashBankReconciliation(token, fromDateStr, toDateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var from = fromDateStr ? parsePaymentDate_(fromDateStr) : null;
  var to   = toDateStr ? endOfDay_(parsePaymentDate_(toDateStr)) : null;

  var openingResult = getOpeningBalances(token);
  var openings = openingResult.data || {};
  var bankAccounts = getSettingsRows_(SETTINGS_TYPE.BANK_ACCOUNT).map(function(r) { return r.key; });

  // Pass 1: roll the stored opening balance forward to the moment right
  // before the report's own start date — only relevant if the report
  // has a "from" date at all; an open-ended report (from the very start
  // of records) has nothing to roll forward through.
  var rolledOpenings = {};
  rolledOpenings[OPENING_BALANCE_CASH_KEY] = openings[OPENING_BALANCE_CASH_KEY] ? openings[OPENING_BALANCE_CASH_KEY].amount : 0;
  bankAccounts.forEach(function(label) {
    rolledOpenings[label] = openings[label] ? openings[label].amount : 0;
  });

  if (from) {
    // FIX: each account must roll forward from its OWN stored opening
    // date, not a single date shared across every account. The previous
    // version found the single earliest opening date across ALL accounts
    // (Cash + every bank) and rolled EVERY account forward through that
    // same window. That's wrong the moment two accounts have different
    // opening dates: e.g. Cash opened 1-Jan and a bank account opened
    // 1-Jun — the bank account's stored ₹ balance already reflects
    // everything up to 1-Jun, so rolling it forward using the 1-Jan
    // window as well double-counted Jan-May movement into that bank
    // account's opening balance for any report starting after 1-Jun.
    // Each account is now rolled forward independently, from its own
    // stored date (if any) to just before the report's start — an
    // account with no stored opening date yet is left at 0, exactly as
    // before.
    var rollForwardEnd = new Date(from.getTime() - 1); // up to, but not including, the report's start

    var rollForwardOneAccount_ = function(key) {
      var stored = openings[key];
      if (!stored || !stored.date) return; // never set — stays at 0
      var storedDate = parsePaymentDate_(stored.date);
      if (storedDate.getTime() >= from.getTime()) return; // opening is on/after the report start — nothing to roll
      var acc = newCashBankAccumulator_(bankAccounts);
      accumulateCashBankMovement_(acc, storedDate, rollForwardEnd);
      rolledOpenings[key] += netMovement_(acc[key]);
    };

    rollForwardOneAccount_(OPENING_BALANCE_CASH_KEY);
    bankAccounts.forEach(rollForwardOneAccount_);
  }

  // Pass 2: the report's own period — these figures are what actually
  // get displayed as Receipts/Payments/Transfers for the chosen range.
  var accounts = newCashBankAccumulator_(bankAccounts);
  accumulateCashBankMovement_(accounts, from, to);

  var cashOpening = rolledOpenings[OPENING_BALANCE_CASH_KEY];
  var cashAcc = accounts[OPENING_BALANCE_CASH_KEY];
  var cashClosing = cashOpening + netMovement_(cashAcc);

  var bankResults = bankAccounts.map(function(label) {
    var opening = rolledOpenings[label];
    var a = accounts[label];
    var closing = opening + netMovement_(a);
    return {
      label: label,
      opening: opening,
      openingDate: openings[label] ? openings[label].date : "",
      receipts: a.receipts,
      cashDepositIn: a.cashDepositIn,
      cashWithdrawalOut: a.cashWithdrawalOut,
      transferIn: a.transferIn,
      transferOut: a.transferOut,
      expenses: a.expenses,
      closing: closing
    };
  });

  return {
    success: true,
    cash: {
      opening: cashOpening,
      openingDate: openings[OPENING_BALANCE_CASH_KEY] ? openings[OPENING_BALANCE_CASH_KEY].date : "",
      receipts: cashAcc.receipts,
      cashWithdrawalIn: cashAcc.cashWithdrawalIn,
      cashDepositOut: cashAcc.cashDepositOut,
      expenses: cashAcc.expenses,
      closing: cashClosing
    },
    banks: bankResults,
    totalCashAndBank: cashClosing + bankResults.reduce(function(s, b) { return s + b.closing; }, 0),
    loanNetPrincipal: (function(){ var n = 0; [OPENING_BALANCE_CASH_KEY].concat(bankAccounts).forEach(function(key){ n += (accounts[key].loanIn || 0) - (accounts[key].loanOutP || 0); }); return n; })(),
    loanReceipts:        (function(){ var n = 0; [OPENING_BALANCE_CASH_KEY].concat(bankAccounts).forEach(function(key){ n += (accounts[key].loanIn   || 0); }); return n; })(),
    loanPrincipalRepaid: (function(){ var n = 0; [OPENING_BALANCE_CASH_KEY].concat(bankAccounts).forEach(function(key){ n += (accounts[key].loanOutP || 0); }); return n; })(),
    loanInterestPaid:    (function(){ var n = 0; [OPENING_BALANCE_CASH_KEY].concat(bankAccounts).forEach(function(key){ n += (accounts[key].loanOutI || 0); }); return n; })()
  };
}

// ── PROFIT & LOSS REPORT ─────────────────────────────────────
// Notes request: "create a Profit and loss tab to know the income &
// expenditure properly with showing opening balances of Cash, bank
// accounts, receipts, and payments, closing balances of accounts, cash
// to tally income and expenditure like tally report." Combines the
// existing Income & Expenditure breakdown with the Cash/Bank
// reconciliation above, computed over the SAME date range, so the two
// halves of the report can never drift out of sync with each other.
function getProfitAndLossReport(token, fromDateStr, toDateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var incomeExp = getIncomeExpenditureSummary(token, fromDateStr, toDateStr);
  if (incomeExp.error) return incomeExp;

  var reconciliation = getCashBankReconciliation(token, fromDateStr, toDateStr);
  if (reconciliation.error) return reconciliation;

  var totalOpening = reconciliation.cash.opening + reconciliation.banks.reduce(function(s, b) { return s + b.opening; }, 0);
  var totalClosing = reconciliation.totalCashAndBank;

  return {
    success: true,
    income: incomeExp.income,
    expenditure: incomeExp.expenditure,
    netSurplus: incomeExp.income.total - incomeExp.expenditure.total,
    cash: reconciliation.cash,
    banks: reconciliation.banks,
    totalOpening: totalOpening,
    totalClosing: totalClosing,
    // What the closing balance SHOULD be if income/expenditure perfectly
    // explained the movement (ignoring the Cash<->Bank transfers, which
    // don't change the combined total) — a tally-style cross-check: this
    // should equal totalClosing exactly, since transfers net to zero
    // across Cash+Bank combined. If they ever don't match, something
    // (e.g. a transaction missing its bank account tag) needs checking.
    expectedClosing: totalOpening + incomeExp.income.total - incomeExp.expenditure.total + (reconciliation.loanNetPrincipal || 0),
    loanReceipts: reconciliation.loanReceipts || 0,
    loanPrincipalRepaid: reconciliation.loanPrincipalRepaid || 0,
    loanInterestPaid: reconciliation.loanInterestPaid || 0
  };
}

// ── SCHOOL IDENTITY (letterhead for printed receipts) ───────
// Single-row settings entry (Type=SchoolInfo, Key="Main") holding the
// school's own name/address/phone/email, used on the receipt print
// template. Falls back to SCHOOL_NAME + blanks if never filled in.
function getSchoolInfo(token) {
  // Allow unauthenticated calls (empty token) for login-page branding:
  // only name and logo are returned — no sensitive data.
  if (token && !validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.SCHOOL_INFO) {
      return {
        success: true,
        name:    safeStr(data[i][SET_COL.VALUE1]) || SCHOOL_NAME,
        address: safeStr(data[i][SET_COL.VALUE2]),
        phone:   safeStr(data[i][SET_COL.VALUE3]),
        email:   safeStr(data[i][SET_COL.VALUE4]),
        logo:    safeStr(data[i][SET_COL.VALUE5])
      };
    }
  }
  // No row yet — return the hardcoded fallback without creating one, so
  // an admin who never opens Settings still gets a usable receipt.
  return { success: true, name: SCHOOL_NAME, address: "", phone: "", email: "", logo: "" };
}

function saveSchoolInfo(token, obj) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };

  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.SCHOOL_INFO) {
      sh.getRange(i + 1, SET_COL.VALUE1 + 1, 1, 5).setValues([[obj.name || "", obj.address || "", obj.phone || "", obj.email || "", obj.logo || ""]]);
      return { success: true, message: "School details updated." };
    }
  }
  // First time saving — append a new row.
  sh.appendRow([SETTINGS_TYPE.SCHOOL_INFO, "Main", obj.name || "", obj.address || "", obj.phone || "", obj.email || "", obj.logo || "", true]);
  return { success: true, message: "School details saved." };
}

// ── FEATURE VISIBILITY (show/hide whole tabs) ───────────────
// Notes request: "create show, hide option in settings for marks entry,
// staff attendance, student attendance... if choose hide, hide from
// tabs also" — this controls whether the nav items themselves render,
// not just whether their content is reachable if you already know the
// URL/section id. Read is open to any logged-in role (everyone needs to
// know whether to show the tab); only an admin can change it.
// Notes request: "Create visible and hide option like fee heads — these
// fee heads option working properly." Rebuilt on the exact same pattern
// as Fee Heads: ONE Settings row per toggle (Key = the feature name),
// each with its own ACTIVE boolean updated in place via setValue() —
// not bundled into one row's Value1/2/3 with a delete-then-append cycle
// on every save, which was the old design and the likely real cause of
// the toggle "resetting" after a refresh.
var FEATURE_VISIBILITY_KEYS = ["marksEntry", "staffAttendance", "studentAttendance"];

function getFeatureVisibility(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var byKey = {};
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) !== SETTINGS_TYPE.FEATURE_VISIBILITY) continue;
    var activeVal = data[i][SET_COL.ACTIVE];
    var isActive = !(activeVal === false || safeStr(activeVal).toUpperCase() === "FALSE");
    byKey[safeStr(data[i][SET_COL.KEY])] = isActive;
  }

  var out = { success: true };
  FEATURE_VISIBILITY_KEYS.forEach(function(key) {
    // No row yet for this feature -> defaults to visible, same as a Fee
    // Head with no explicit Active value would. Keeps any install
    // upgrading from before this setting existed from losing a tab they
    // were already using.
    out[key] = (key in byKey) ? byKey[key] : true;
  });
  return out;
}

function setFeatureVisible(token, key, visible) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (FEATURE_VISIBILITY_KEYS.indexOf(key) === -1) return { error: "Unknown feature: " + key };

  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.FEATURE_VISIBILITY && safeStr(data[i][SET_COL.KEY]) === key) {
      sh.getRange(i + 1, SET_COL.ACTIVE + 1).setValue(!!visible);
      SpreadsheetApp.flush();
      return { success: true, message: visible ? "Tab is now visible." : "Tab hidden from the menu." };
    }
  }
  // No row yet for this feature — create it now, defaulting Active to
  // whatever was just requested (exactly how a brand-new Fee Head row
  // is created the first time it's added).
  sh.appendRow([SETTINGS_TYPE.FEATURE_VISIBILITY, key, "", "", "", "", "", !!visible]);
  SpreadsheetApp.flush();
  return { success: true, message: visible ? "Tab is now visible." : "Tab hidden from the menu." };
}

// ── PAYMENT MODE VISIBILITY (Cash / Bank) ────────────────────
// Notes request: "create an edit option for which mode is to be
// visible" — same exact pattern as Fee Heads and the Tab Visibility
// above: one Settings row per mode, each with its own ACTIVE boolean
// updated in place. Unlike Fee Heads, at least one mode must always
// stay visible — a school can't record any payment at all with both
// Cash and Bank turned off, so that combination is rejected outright.
var PAYMENT_MODE_KEYS = ["Cash", "Bank"];

function getPaymentModeVisibility(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var byKey = {};
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) !== SETTINGS_TYPE.PAYMENT_MODE_VISIBILITY) continue;
    var activeVal = data[i][SET_COL.ACTIVE];
    var isActive = !(activeVal === false || safeStr(activeVal).toUpperCase() === "FALSE");
    byKey[safeStr(data[i][SET_COL.KEY])] = isActive;
  }
  var out = { success: true };
  PAYMENT_MODE_KEYS.forEach(function(key) {
    out[key] = (key in byKey) ? byKey[key] : true; // defaults to visible, same reasoning as Tab Visibility
  });
  return out;
}

function setPaymentModeVisible(token, mode, visible) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (PAYMENT_MODE_KEYS.indexOf(mode) === -1) return { error: "Unknown payment mode: " + mode };

  if (!visible) {
    // Refuse to hide the LAST visible mode — confirmed no school can
    // record a payment at all with neither Cash nor Bank available.
    var current = getPaymentModeVisibility(token);
    if (current.error) return current;
    var otherKey = mode === "Cash" ? "Bank" : "Cash";
    if (!current[otherKey]) {
      return { error: "At least one payment mode (Cash or Bank) must stay visible." };
    }
  }

  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.PAYMENT_MODE_VISIBILITY && safeStr(data[i][SET_COL.KEY]) === mode) {
      sh.getRange(i + 1, SET_COL.ACTIVE + 1).setValue(!!visible);
      SpreadsheetApp.flush();
      return { success: true, message: mode + " is now " + (visible ? "visible." : "hidden from payment forms.") };
    }
  }
  sh.appendRow([SETTINGS_TYPE.PAYMENT_MODE_VISIBILITY, mode, "", "", "", "", "", !!visible]);
  SpreadsheetApp.flush();
  return { success: true, message: mode + " is now " + (visible ? "visible." : "hidden from payment forms.") };
}

// ============================================================
//  PHOTO UPLOAD — School Logo, Student Photo, Teacher/Staff Photo
//  Notes request: upload from device OR paste a Google Drive link;
//  every photo is stored in one shared Drive folder (PHOTO_FOLDER_ID).
// ============================================================

function getPhotoFolder_() {
  return DriveApp.getFolderById(PHOTO_FOLDER_ID);
}

// Uploads a file picked on the user's device. `base64Data` is the file
// content without the "data:image/...;base64," prefix (stripped client-
// side); `mimeType` and `fileName` come from the browser's File object.
// `category` ("school" | "student" | "teacher") and `entityId` (e.g. a
// Student ID) are used only to build a readable filename in Drive —
// they don't affect where the database record points.
function uploadPhotoFromDevice(token, base64Data, mimeType, fileName, category, entityId) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (!base64Data) return { error: "No file data received." };

  var maxBytes = 5 * 1024 * 1024; // 5MB safety cap — base64 is ~33% larger than raw bytes
  if (base64Data.length > maxBytes * 1.4) {
    return { error: "Image is too large. Please use a photo under 5MB." };
  }

  try {
    var folder = getPhotoFolder_();
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", buildPhotoFileName_(category, entityId, fileName));
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: driveDirectImageUrl_(file.getId()), fileId: file.getId() };
  } catch (e) {
    return { error: "Upload failed: " + e.message };
  }
}

function buildPhotoFileName_(category, entityId, originalName) {
  var prefix = { school: "SchoolLogo", student: "Student_", teacher: "Teacher_" }[category] || "Photo_";
  var ext = (originalName && originalName.match(/\.[A-Za-z0-9]+$/)) ? originalName.match(/\.[A-Za-z0-9]+$/)[0] : ".jpg";
  return prefix + (entityId || Utilities.getUuid().slice(0, 8)) + "_" + new Date().getTime() + ext;
}

// Converts a Drive file ID into a URL that reliably renders inside an
// <img> tag. The older "uc?export=view" endpoint opens a preview page
// when clicked directly, but Google increasingly blocks it from being
// hotlinked inside <img src="..."> — it can fail silently (broken image
// icon) even though the file itself is shared correctly. The thumbnail
// endpoint is the format Google's own embed/picker tools use and is far
// more reliable for this purpose. "=s1000" requests up to a 1000px image;
// Drive returns a smaller one if the original is smaller, never upscales.
function driveDirectImageUrl_(fileId) {
  return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
}

// ── STUDENT PROFILE SHEET READ ──────────────────────────────
// ── REQUEST-SCOPED DATA CACHE ────────────────────────────────
// Each GAS execution is a fresh V8 context, so these variables are
// reset to null at the start of every HTTP request — they live only
// for the duration of one function call. This means calling
// getAllStudentProfiles_() or getFeeMap_() five times in one request
// (dashboard + fee list + outstanding + profile + summary) only reads
// the spreadsheet ONCE; every subsequent call returns the same object.
var _cachedProfiles_ = null;
var _cachedFeeMap_   = null;
var _cachedSettings_ = null; // raw Settings sheet data
var _settingsSeeded_ = false; // skip repeat seed/migration checks

function getAllStudentProfiles_() {
  if (_cachedProfiles_) return _cachedProfiles_;
  var sh = getSheet(SHEETS.PROFILE);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 2; i < data.length; i++) {
    var r = data[i];
    if (!r[SP_COL.NAME] || r[SP_COL.NAME] === "") continue;
    out.push({
      rowIndex:     i,
      studentId:    safeStr(r[SP_COL.STUDENT_ID]),
      name:         safeStr(r[SP_COL.NAME]),
      aadhar:       safeStr(r[SP_COL.AADHAR]),
      childId:      safeStr(r[SP_COL.CHILD_ID]),
      pen:          safeStr(r[SP_COL.PEN]),
      apaar:        safeStr(r[SP_COL.APAAR]),
      joinDate:     formatDateSafe_(r[SP_COL.JOIN_DATE]),
      admNo:        safeStr(r[SP_COL.ADM_NO]),
      dob:          formatDateSafe_(r[SP_COL.DOB]),
      age:          safeStr(r[SP_COL.AGE]),
      gender:       safeStr(r[SP_COL.GENDER]),
      class:        safeStr(r[SP_COL.CLASS]),
      section:      safeStr(r[SP_COL.SECTION]),
      transport:    safeStr(r[SP_COL.TRANSPORT]),
      address:      safeStr(r[SP_COL.ADDRESS]),
      caste:        safeStr(r[SP_COL.CASTE]),
      subCaste:     safeStr(r[SP_COL.SUBCASTE]),
      fatherName:   safeStr(r[SP_COL.FATHER_NAME]),
      fatherAadhar: safeStr(r[SP_COL.FATHER_AADHAR]),
      fatherOcc:    safeStr(r[SP_COL.FATHER_OCC]),
      fatherMobile: safeStr(r[SP_COL.FATHER_MOBILE]),
      motherName:   safeStr(r[SP_COL.MOTHER_NAME]),
      motherAadhar: safeStr(r[SP_COL.MOTHER_AADHAR]),
      motherOcc:    safeStr(r[SP_COL.MOTHER_OCC]),
      motherMobile: safeStr(r[SP_COL.MOTHER_MOBILE]),
      rationCard:   safeStr(r[SP_COL.RATION_CARD]),
      bankAcc:      safeStr(r[SP_COL.BANK_ACC]),
      bankName:     safeStr(r[SP_COL.BANK_NAME]),
      ifsc:         safeStr(r[SP_COL.IFSC]),
      branch:       safeStr(r[SP_COL.BRANCH]),
      photo:        safeStr(r[SP_COL.PHOTO]).indexOf("#VALUE") === 0 ? "" : safeStr(r[SP_COL.PHOTO]),
      status:       safeStr(r[SP_COL.STATUS]) || "Active",
      udise:        safeStr(r[SP_COL.UDISE]),
      remarks:      safeStr(r[SP_COL.REMARKS])
    });
  }
  _cachedProfiles_ = out;
  return out;
}

function formatDateSafe_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    // Guard against the Excel epoch bug rows we saw (e.g. year 1915/1916)
    if (v.getFullYear() < 1950) return "";
    return Utilities.formatDate(v, Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy");
  }
  return safeStr(v);
}

function getStudentList(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var profiles = getAllStudentProfiles_().filter(function(p) { return p.status !== "Removed"; });
  var feeMap   = getFeeMap_();   // keyed by studentId -> {billed, paid, due, heads}
  var out = profiles.map(function(p) {
    var f = feeMap[p.studentId] || { billed: 0, paid: 0, due: 0, heads: {} };
    return {
      rowIndex:  p.rowIndex,
      studentId: p.studentId,
      name:      p.name,
      class:     p.class,
      section:   p.section,
      transport: p.transport,
      gender:    p.gender,
      status:    p.status,
      photo:     p.photo,
      billed:    f.billed,
      paid:      f.paid,
      due:       f.due,
      heads:     f.heads || {}   // head-wise breakdown for head filter in Fee Management
    };
  });
  return { success: true, data: out };
}

// ── STUDENT REGISTER EXCEL EXPORT ───────────────────────────
// Full profile + fee summary for every active student. The client
// builds the actual .xlsx via SheetJS (already loaded for Bulk Import) —
// this just hands back the rows.
function getStudentRegisterData(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var profiles = getAllStudentProfiles_().filter(function(p) { return p.status !== "Removed"; });
  var feeMap = getFeeMap_();
  var data = profiles.map(function(p) {
    var f = feeMap[p.studentId] || { billed: 0, paid: 0, due: 0 };
    return {
      studentId: p.studentId, name: p.name, admNo: p.admNo, class: p.class, section: p.section,
      gender: p.gender, dob: p.dob, age: p.age, joinDate: p.joinDate, transport: p.transport,
      caste: p.caste, subCaste: p.subCaste, aadhar: p.aadhar, childId: p.childId, pen: p.pen, apaar: p.apaar,
      udise: p.udise, address: p.address,
      fatherName: p.fatherName, fatherAadhar: p.fatherAadhar, fatherOcc: p.fatherOcc, fatherMobile: p.fatherMobile,
      motherName: p.motherName, motherAadhar: p.motherAadhar, motherOcc: p.motherOcc, motherMobile: p.motherMobile,
      rationCard: p.rationCard, bankAcc: p.bankAcc, bankName: p.bankName, ifsc: p.ifsc, branch: p.branch,
      status: p.status, remarks: p.remarks,
      billed: f.billed, paid: f.paid, due: f.due
    };
  });
  return { success: true, data: data, academicYear: getActiveAcademicYear_() };
}

// ── CLASS PROMOTION (ACADEMIC-YEAR ROLLOVER) ────────────────
// Uses the Class list already configured under Settings → Classes — same
// list the Add-Student "Class" dropdown is built from — as the
// promotion order. e.g. if Settings has Nursery, LKG, UKG, 1, 2 … 10 in
// that order, a student in "3" moves to "4" when promoted. The LAST
// class in that list is treated as the graduating class and is
// deliberately NOT auto-promoted or auto-removed — those students are
// flagged here for the admin to review and move to Dropouts manually,
// since "this student has left school" is a decision, not something to
// infer automatically from a class list.
function getPromotionPreview(token) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  var classList = getSettingsRows_(SETTINGS_TYPE.CLASS).map(function(r) { return r.key; });
  if (classList.length < 2) {
    return { error: "Add at least two classes, in promotion order (e.g. 1, 2, 3 …), under Settings → Classes before running promotion." };
  }

  var profiles = getAllStudentProfiles_().filter(function(p) { return p.status !== "Removed"; });
  var byClass = {};
  profiles.forEach(function(p) { byClass[p.class] = (byClass[p.class] || 0) + 1; });

  var rows = classList.map(function(cls, i) {
    var next = (i === classList.length - 1) ? null : classList[i + 1];
    return { class: cls, count: byClass[cls] || 0, next: next, graduating: next === null };
  });

  var unmapped = [];
  Object.keys(byClass).forEach(function(cls) {
    if (classList.indexOf(cls) === -1) unmapped.push({ class: cls || "(blank)", count: byClass[cls] });
  });

  var archiveName = "Students_" + getActiveAcademicYear_();
  var alreadyArchived = !!_getSS_().getSheetByName(archiveName);

  return {
    success: true,
    academicYear: getActiveAcademicYear_(),
    archiveName: archiveName,
    alreadyArchived: alreadyArchived,
    rows: rows,
    unmapped: unmapped,
    totalActive: profiles.length
  };
}

function runClassPromotion(token) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };

  var classList = getSettingsRows_(SETTINGS_TYPE.CLASS).map(function(r) { return r.key; });
  if (classList.length < 2) {
    return { error: "Add at least two classes, in promotion order, under Settings → Classes before running promotion." };
  }

  var ay = getActiveAcademicYear_();
  var archiveName = "Students_" + ay;
  var ss = _getSS_();
  if (ss.getSheetByName(archiveName)) {
    return { error: "A \"" + archiveName + "\" archive already exists — promotion for this year has likely already been run. Rename or delete that sheet first if you really need to run it again." };
  }

  var profiles = getAllStudentProfiles_().filter(function(p) { return p.status !== "Removed"; });
  var feeMap = getFeeMap_();

  // 1) Archive the full current roster, values-only, BEFORE touching
  // anything — this is the undo/safety-net for the whole operation.
  var archiveHeader = ['Student ID', 'Name', 'Adm No', 'Class', 'Section', 'Gender', 'DOB', 'Age', 'Join Date',
    'Transport', 'Father Name', 'Father Mobile', 'Mother Name', 'Mother Mobile', 'Address', 'Status',
    'Billed', 'Paid', 'Due', 'Archived Academic Year'];
  var archiveRows = profiles.map(function(p) {
    var f = feeMap[p.studentId] || { billed: 0, paid: 0, due: 0 };
    return [p.studentId, p.name, p.admNo, p.class, p.section, p.gender, p.dob, p.age, p.joinDate,
      p.transport, p.fatherName, p.fatherMobile, p.motherName, p.motherMobile, p.address, p.status,
      f.billed, f.paid, f.due, ay];
  });
  var archiveSheet = ss.insertSheet(archiveName);
  archiveSheet.getRange(1, 1, 1, archiveHeader.length).setValues([archiveHeader]);
  if (archiveRows.length) archiveSheet.getRange(2, 1, archiveRows.length, archiveHeader.length).setValues(archiveRows);
  archiveSheet.setFrozenRows(1);

  // 2) Promote everyone one step up the configured Class order.
  var sh = getSheet(SHEETS.PROFILE);
  var promoted = 0;
  var graduating = [];
  var unmapped = [];
  profiles.forEach(function(p) {
    var idx = classList.indexOf(p.class);
    if (idx === -1) {
      unmapped.push({ studentId: p.studentId, name: p.name, class: p.class });
      return;
    }
    if (idx === classList.length - 1) {
      graduating.push({ studentId: p.studentId, name: p.name, class: p.class });
      return;
    }
    sh.getRange(p.rowIndex + 1, SP_COL.CLASS + 1).setValue(classList[idx + 1]);
    promoted++;
  });

  invalidateCaches_();
  PropertiesService.getScriptProperties().setProperty('ACTIVE_ACADEMIC_YEAR', _nextAcademicYear_(ay));
  return {
    success: true,
    archiveSheet: archiveName,
    archivedCount: profiles.length,
    promotedCount: promoted,
    graduating: graduating,
    unmapped: unmapped,
    newActiveYear: _nextAcademicYear_(ay),
    message: "Archived " + profiles.length + " students to \"" + archiveName + "\" and promoted " + promoted + " to their next class."
  };
}

// ── TUITION FEE → combined billed amount per student ────────
// v5: Tution_Fee is now LONG FORMAT (one row per student per fee head),
// so this works for any number of custom fee heads created in Settings
// with no code change. Still keyed by Student ID, not by name — this is
// what fixes the duplicate-name billing collision (Notes item #7) and
// the wrong-column Dashboard bug (Notes item #1). The old separate
// "Bus_Fee" sheet is never read (Notes item #10).
function getFeeMap_() {
  if (_cachedFeeMap_) return _cachedFeeMap_;
  var map = {}; // studentId -> { billed, paid, due, heads: { headName: {billed, paid} } }

  var tSh = getSheet(SHEETS.TUITION);
  var tData = tSh.getDataRange().getValues();
  for (var i = 1; i < tData.length; i++) {
    var r = tData[i];
    var id = safeStr(r[TF_COL.STUDENT_ID]);
    var head = safeStr(r[TF_COL.FEE_HEAD]);
    if (!id || !head) continue;
    if (!map[id]) map[id] = { billed: 0, paid: 0, due: 0, heads: {} };

    var billedAmt = safeNum(r[TF_COL.BILLED]);
    var discount  = safeNum(r[TF_COL.DISCOUNT]);
    var netAmt    = safeNum(r[TF_COL.NET]) || (billedAmt - discount);

    if (!map[id].heads[head]) map[id].heads[head] = { billed: 0, paid: 0 };
    map[id].heads[head].billed += netAmt;
    map[id].billed += netAmt;
  }

  // Overlay actual payments from Transactions, broken down by fee head.
  // Matched by Student ID, and voided/cancelled rows are skipped entirely
  // so a cancelled payment never counts toward "paid" (Notes item #6).
  var txSh = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();
  for (var k = 1; k < txData.length; k++) {
    var t = txData[k];
    var sid = safeStr(t[TX_COL.STUDENT_ID]);
    if (!sid || isVoided_(t[TX_COL.VOID])) continue;
    var payHead = safeStr(t[TX_COL.HEAD]) || "Other";
    var amt  = safeNum(t[TX_COL.AMOUNT]);
    if (!map[sid]) map[sid] = { billed: 0, paid: 0, due: 0, heads: {} };
    if (!map[sid].heads[payHead]) map[sid].heads[payHead] = { billed: 0, paid: 0 };
    map[sid].heads[payHead].paid += amt;
    map[sid].paid += amt;
  }

  Object.keys(map).forEach(function(id) {
    map[id].due = map[id].billed - map[id].paid;
  });

  _cachedFeeMap_ = map;
  return map;
}

// ── INDIVIDUAL STUDENT PROFILE (full tracking) ──────────────
function getStudentProfileDetail(token, studentId) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var profiles = getAllStudentProfiles_();
  var profile  = profiles.filter(function(p) { return p.studentId === studentId; })[0];
  if (!profile) return { error: "Student not found." };

  var feeMap = getFeeMap_();
  var fee = feeMap[studentId] || { billed: 0, paid: 0, due: 0, heads: {} };

  // Build fee-head breakdown including heads with no billed amount but with payments
  var headBreakdown = [];
  var seen = {};
  var feeHeadRows = getSettingsRows_(SETTINGS_TYPE.FEE_HEAD);
  feeHeadRows.forEach(function(row) {
    var h = row.key;
    var d = fee.heads[h] || { billed: 0, paid: 0 };
    headBreakdown.push({ head: h, billed: d.billed, paid: d.paid, due: d.billed - d.paid });
    seen[h] = true;
  });
  Object.keys(fee.heads).forEach(function(h) {
    if (!seen[h]) {
      var d = fee.heads[h];
      headBreakdown.push({ head: h, billed: d.billed, paid: d.paid, due: d.billed - d.paid });
    }
  });

  // Transaction history for this student (by Student ID), most recent first.
  // Includes voided entries so the profile screen can show a struck-through
  // "Cancelled" line rather than silently deleting history (Notes item #6).
  var txSh   = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();
  var history = [];
  for (var i = 1; i < txData.length; i++) {
    var t = txData[i];
    if (safeStr(t[TX_COL.STUDENT_ID]) !== studentId) continue;
    history.push({
      rowIndex: i,
      date:     formatDateSafe_(t[TX_COL.DATE]) || safeStr(t[TX_COL.DATE]),
      class:    safeStr(t[TX_COL.CLASS]) || profile.class,
      receipt:  safeStr(t[TX_COL.RECEIPT]),
      head:     safeStr(t[TX_COL.HEAD]),
      mode:     safeStr(t[TX_COL.MODE]),
      amount:   safeNum(t[TX_COL.AMOUNT]),
      by:       safeStr(t[TX_COL.BY]),
      voided:   isVoided_(t[TX_COL.VOID])
    });
  }
  history.sort(function(a, b) { return (b.date < a.date) ? -1 : 1; });

  return {
    success: true,
    profile: profile,
    fee: {
      billed: fee.billed,
      paid:   fee.paid,
      due:    fee.due,
      heads:  headBreakdown
    },
    transactions: history
  };
}

// ── ADD NEW STUDENT ─────────────────────────────────────────
// v3: captures the full Student Profile field set (not just 4 fields) and
// writes a complete row into Tution_Fee covering all 5 fee heads, not just
// Tuition Fee — see Notes items #3 and #4. A new Student ID is generated
// here and used as the link between Student Profile, Tution_Fee and every
// future Transactions row (Notes item #7).
function addStudent(token, obj) {
  invalidateCaches_();
  var sess = validateSession(token);
  // Clerk can add individual students; only admin/principal can bulk-import
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin" && sess.role !== "principal" && sess.role !== "clerk") {
    return { error: PERM_MSG.MGMT_ONLY };
  }
  if (!obj || !obj.name) return { error: "Student name is required." };

  var lock = LockService.getScriptLock();
  var studentId;
  try {
    lock.waitLock(10000);
    studentId = getNextStudentId_();

    var sh  = getSheet(SHEETS.PROFILE);
    var row = sh.getLastRow() + 1;
    var arr = new Array(SP_NUM_COLS).fill("");

    arr[SP_COL.STUDENT_ID]    = studentId;
    arr[SP_COL.NAME]          = obj.name;
    arr[SP_COL.AADHAR]        = obj.aadhar || "";
    arr[SP_COL.CHILD_ID]      = obj.childId || "";
    arr[SP_COL.PEN]           = obj.pen || "";
    arr[SP_COL.APAAR]         = obj.apaar || "";
    arr[SP_COL.JOIN_DATE]     = obj.joinDate ? parsePaymentDate_(obj.joinDate) : new Date();
    arr[SP_COL.ADM_NO]        = obj.admNo || "";
    arr[SP_COL.DOB]           = obj.dob ? parsePaymentDate_(obj.dob) : "";
    arr[SP_COL.AGE]           = obj.age || "";
    arr[SP_COL.GENDER]        = obj.gender || "";
    arr[SP_COL.CLASS]         = obj.class || "";
    arr[SP_COL.SECTION]       = obj.section || "";
    arr[SP_COL.TRANSPORT]     = obj.transport || "";
    arr[SP_COL.ADDRESS]       = obj.address || "";
    arr[SP_COL.CASTE]         = obj.caste || "";
    arr[SP_COL.SUBCASTE]      = obj.subCaste || "";
    arr[SP_COL.FATHER_NAME]   = obj.fatherName || "";
    arr[SP_COL.FATHER_AADHAR] = obj.fatherAadhar || "";
    arr[SP_COL.FATHER_OCC]    = obj.fatherOcc || "";
    arr[SP_COL.FATHER_MOBILE] = obj.fatherMobile || "";
    arr[SP_COL.MOTHER_NAME]   = obj.motherName || "";
    arr[SP_COL.MOTHER_AADHAR] = obj.motherAadhar || "";
    arr[SP_COL.MOTHER_OCC]    = obj.motherOcc || "";
    arr[SP_COL.MOTHER_MOBILE] = obj.motherMobile || "";
    arr[SP_COL.RATION_CARD]   = obj.rationCard || "";
    arr[SP_COL.BANK_ACC]      = obj.bankAcc || "";
    arr[SP_COL.BANK_NAME]     = obj.bankName || "";
    arr[SP_COL.IFSC]          = obj.ifsc || "";
    arr[SP_COL.BRANCH]        = obj.branch || "";
    arr[SP_COL.PHOTO]         = obj.photo || "";
    arr[SP_COL.STATUS]        = "Active";
    arr[SP_COL.UDISE]         = obj.udise || "";
    arr[SP_COL.REMARKS]       = obj.remarks || "";

    sh.getRange(row, 1, 1, arr.length).setValues([arr]);

    // Write one Tution_Fee row per active fee head — v5 long format, so
    // this works for any number of custom heads created in Settings with
    // no code change (Notes request). obj.feeAmounts is { headName:
    // { billed, discount } }, built client-side from whichever fee-head
    // inputs were shown on the Add Student form (Settings-driven).
    var tSh  = getSheet(SHEETS.TUITION);
    var billingYear = obj.financialYear || getAcademicYear_();
    var feeAmounts = obj.feeAmounts || {};
    var feeRows = [];
    Object.keys(feeAmounts).forEach(function(headName) {
      var billedAmt = safeNum(feeAmounts[headName].billed);
      var discountAmt = safeNum(feeAmounts[headName].discount);
      if (billedAmt <= 0 && discountAmt <= 0) return; // skip heads left blank
      var rowArr = new Array(TF_NUM_COLS).fill("");
      rowArr[TF_COL.STUDENT_ID] = studentId;
      rowArr[TF_COL.NAME]       = obj.name;
      rowArr[TF_COL.CLASS]      = obj.class || "";
      rowArr[TF_COL.SECTION]    = obj.section || "";
      rowArr[TF_COL.FEE_HEAD]   = headName;
      rowArr[TF_COL.BILLED]     = billedAmt;
      rowArr[TF_COL.DISCOUNT]   = discountAmt;
      rowArr[TF_COL.NET]        = billedAmt - discountAmt;
      rowArr[TF_COL.ACADEMIC_YEAR] = billingYear;
      feeRows.push(rowArr);
    });
    if (feeRows.length) {
      var tRow = tSh.getLastRow() + 1;
      tSh.getRange(tRow, 1, feeRows.length, TF_NUM_COLS).setValues(feeRows);
    }

  } finally {
    lock.releaseLock();
  }

  return { success: true, studentId: studentId, message: "Student '" + obj.name + "' added successfully with ID " + studentId + "." };
}

// ── BULK STUDENT IMPORT (Excel/CSV template) ────────────────
// Notes request: download a template, fill it in outside the app, upload
// it back to create many students at once with auto-generated Student
// IDs. `rows` is an array of plain objects already parsed client-side
// from the uploaded CSV (header names map 1:1 to the keys below). Bad
// rows are skipped and reported (confirmed); good rows are written in
// one batch. Each row also carries an optional `feeAmounts` object —
// { headName: { billed, discount } }, same shape addStudent() already
// uses — built client-side from the dynamic per-fee-head Billed/Discount
// column pairs the template now includes (one pair per active fee head
// in Settings), so fee details can be filled in at import time instead
// of one student at a time afterward. A row with all fee columns left
// blank simply gets no Tution_Fee rows, same as before.
function bulkAddStudents(token, rows) {
  var sess = validateSession(token);
  var permErr = requireManagement_(sess);
  if (permErr) return permErr;
  if (!rows || !rows.length) return { error: "No rows to import." };
  if (rows.length > 500) return { error: "Please import in batches of 500 or fewer rows." };

  var validClasses = getSettingsRows_(SETTINGS_TYPE.CLASS).map(function(r) { return r.key; });
  var validGenders = ["Male", "Female", "Boy", "Girl", "Other"];
  var billingYear = getAcademicYear_();

  var lock = LockService.getScriptLock();
  var created = [], errors = [];
  try {
    lock.waitLock(15000);
    var sh = getSheet(SHEETS.PROFILE);

    // Compute the starting sequence ONCE, then increment in-memory for
    // the whole batch — calling getNextStudentId_() once per row would
    // re-scan the full sheet every time AND wouldn't see IDs generated
    // earlier in this same batch (since they aren't written yet).
    var data = sh.getDataRange().getValues();
    var maxSeq = 0;
    for (var i = 2; i < data.length; i++) {
      var existingId = safeStr(data[i][SP_COL.STUDENT_ID]);
      if (existingId.indexOf(STUDENT_ID_PREFIX) === 0) {
        var seq = parseInt(existingId.substring(STUDENT_ID_PREFIX.length), 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    }

    var newRows = [];
    var feeRows = [];
    rows.forEach(function(r, idx) {
      var rowNum = idx + 2; // +2: 1 for header, 1 for 1-based display
      var name = safeStr(r.name).trim();
      if (!name) { errors.push({ row: rowNum, reason: "Missing student name — row skipped." }); return; }

      var cls = safeStr(r.class).trim();
      if (!cls) { errors.push({ row: rowNum, reason: "Missing class — row skipped." }); return; }
      if (validClasses.length && validClasses.indexOf(cls) === -1) {
        errors.push({ row: rowNum, reason: "\"" + cls + "\" is not a configured class (check Settings → Classes) — row skipped." });
        return;
      }

      var gender = safeStr(r.gender).trim();
      if (gender && validGenders.indexOf(gender) === -1) {
        errors.push({ row: rowNum, reason: "\"" + gender + "\" is not a recognized gender (use Male/Female/Other) — row skipped." });
        return;
      }

      maxSeq++;
      var studentId = STUDENT_ID_PREFIX + ("0000" + maxSeq).slice(-4);

      var arr = new Array(SP_NUM_COLS).fill("");
      arr[SP_COL.STUDENT_ID]    = studentId;
      arr[SP_COL.NAME]          = name;
      arr[SP_COL.AADHAR]        = safeStr(r.aadhar);
      arr[SP_COL.CHILD_ID]      = safeStr(r.childId);
      arr[SP_COL.PEN]           = safeStr(r.pen);
      arr[SP_COL.APAAR]         = safeStr(r.apaar);
      arr[SP_COL.JOIN_DATE]     = r.joinDate ? (parseAnyDate_(r.joinDate) || new Date()) : new Date();
      arr[SP_COL.ADM_NO]        = safeStr(r.admNo);
      arr[SP_COL.DOB]           = r.dob ? (parseAnyDate_(r.dob) || "") : "";
      arr[SP_COL.AGE]           = safeStr(r.age);
      arr[SP_COL.GENDER]        = gender;
      arr[SP_COL.CLASS]         = cls;
      arr[SP_COL.SECTION]       = safeStr(r.section).trim();
      arr[SP_COL.TRANSPORT]     = safeStr(r.transport);
      arr[SP_COL.ADDRESS]       = safeStr(r.address);
      arr[SP_COL.CASTE]         = safeStr(r.caste);
      arr[SP_COL.SUBCASTE]      = safeStr(r.subCaste);
      arr[SP_COL.FATHER_NAME]   = safeStr(r.fatherName);
      arr[SP_COL.FATHER_AADHAR] = safeStr(r.fatherAadhar);
      arr[SP_COL.FATHER_OCC]    = safeStr(r.fatherOcc);
      arr[SP_COL.FATHER_MOBILE] = safeStr(r.fatherMobile);
      arr[SP_COL.MOTHER_NAME]   = safeStr(r.motherName);
      arr[SP_COL.MOTHER_AADHAR] = safeStr(r.motherAadhar);
      arr[SP_COL.MOTHER_OCC]    = safeStr(r.motherOcc);
      arr[SP_COL.MOTHER_MOBILE] = safeStr(r.motherMobile);
      arr[SP_COL.RATION_CARD]   = safeStr(r.rationCard);
      arr[SP_COL.BANK_ACC]      = safeStr(r.bankAcc);
      arr[SP_COL.BANK_NAME]     = safeStr(r.bankName);
      arr[SP_COL.IFSC]          = safeStr(r.ifsc);
      arr[SP_COL.BRANCH]        = safeStr(r.branch);
      arr[SP_COL.STATUS]        = "Active";
      arr[SP_COL.UDISE]         = safeStr(r.udise);
      arr[SP_COL.REMARKS]       = safeStr(r.remarks);

      newRows.push(arr);
      created.push({ row: rowNum, studentId: studentId, name: name, class: cls, section: arr[SP_COL.SECTION] });

      // Optional per-row fee amounts — { headName: { billed, discount } },
      // built client-side from the template's dynamic fee-head columns.
      var feeAmounts = r.feeAmounts || {};
      Object.keys(feeAmounts).forEach(function(headName) {
        var billedAmt = safeNum(feeAmounts[headName].billed);
        var discountAmt = safeNum(feeAmounts[headName].discount);
        if (billedAmt <= 0 && discountAmt <= 0) return; // skip heads left blank for this student
        var fRow = new Array(TF_NUM_COLS).fill("");
        fRow[TF_COL.STUDENT_ID] = studentId;
        fRow[TF_COL.NAME]       = name;
        fRow[TF_COL.CLASS]      = cls;
        fRow[TF_COL.SECTION]    = arr[SP_COL.SECTION];
        fRow[TF_COL.FEE_HEAD]   = headName;
        fRow[TF_COL.BILLED]     = billedAmt;
        fRow[TF_COL.DISCOUNT]   = discountAmt;
        fRow[TF_COL.NET]        = billedAmt - discountAmt;
        fRow[TF_COL.ACADEMIC_YEAR] = billingYear;
        feeRows.push(fRow);
      });
    });

    if (newRows.length) {
      var startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, newRows.length, SP_NUM_COLS).setValues(newRows);
    }
    if (feeRows.length) {
      var tSh = getSheet(SHEETS.TUITION);
      var tStartRow = tSh.getLastRow() + 1;
      tSh.getRange(tStartRow, 1, feeRows.length, TF_NUM_COLS).setValues(feeRows);
    }
    if (newRows.length || feeRows.length) SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return {
    success: true,
    createdCount: created.length,
    errorCount: errors.length,
    created: created,
    errors: errors,
    message: created.length
      ? "Imported " + created.length + " student(s)" + (feeRows.length ? " with fee details" : "") + (errors.length ? ", " + errors.length + " row(s) skipped." : ".")
      : "No students were imported — every row had a problem."
  };
}

// Parses a date that might arrive as "DD-MM-YYYY", "DD/MM/YYYY",
// "YYYY-MM-DD", or an Excel/Sheets serial date number (since browsers
// reading the uploaded CSV may hand back numbers for date-formatted
// cells depending on how the sheet was originally saved). Returns a
// real Date or null if it can't make sense of the input — callers fall
// back to leaving the field blank rather than guessing.
function parseAnyDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") return value;
  var str = safeStr(value).trim();
  if (!str) return null;

  // Excel serial date (days since 1899-12-30) — typically a bare number.
  if (/^\d+(\.\d+)?$/.test(str)) {
    var serial = parseFloat(str);
    if (serial > 0 && serial < 100000) {
      return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    }
  }

  var m = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));

  m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));

  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

// ── UPDATE EXISTING STUDENT (profile fields) ────────────────
// New in v3 — Notes item #6. Edits the Student Profile row in place by
// Student ID. Only overwrites fields actually supplied, so a partial edit
// (e.g. just a phone number) can't blank out the rest of the record. Fee
// amounts are NOT touched here — fee corrections happen via Fee
// Management / cancelPayment(), keeping financial and biodata edits separate.
function updateStudent(token, studentId, obj) {
  invalidateCaches_();
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  if (!studentId) return { error: "Student ID is required." };

  var sh = getSheet(SHEETS.PROFILE);
  var data = sh.getDataRange().getValues();
  var targetRow = -1;
  for (var i = 2; i < data.length; i++) {
    if (safeStr(data[i][SP_COL.STUDENT_ID]) === studentId) { targetRow = i + 1; break; }
  }
  if (targetRow === -1) return { error: "Student not found." };

  var fieldMap = {
    name: SP_COL.NAME, aadhar: SP_COL.AADHAR, childId: SP_COL.CHILD_ID, pen: SP_COL.PEN,
    apaar: SP_COL.APAAR, admNo: SP_COL.ADM_NO, age: SP_COL.AGE, gender: SP_COL.GENDER,
    class: SP_COL.CLASS, section: SP_COL.SECTION, transport: SP_COL.TRANSPORT, address: SP_COL.ADDRESS,
    caste: SP_COL.CASTE, subCaste: SP_COL.SUBCASTE, fatherName: SP_COL.FATHER_NAME,
    fatherAadhar: SP_COL.FATHER_AADHAR, fatherOcc: SP_COL.FATHER_OCC, fatherMobile: SP_COL.FATHER_MOBILE,
    motherName: SP_COL.MOTHER_NAME, motherAadhar: SP_COL.MOTHER_AADHAR, motherOcc: SP_COL.MOTHER_OCC,
    motherMobile: SP_COL.MOTHER_MOBILE, rationCard: SP_COL.RATION_CARD, bankAcc: SP_COL.BANK_ACC,
    bankName: SP_COL.BANK_NAME, ifsc: SP_COL.IFSC, branch: SP_COL.BRANCH, status: SP_COL.STATUS,
    udise: SP_COL.UDISE, remarks: SP_COL.REMARKS
  };
  Object.keys(fieldMap).forEach(function(key) {
    if (obj[key] !== undefined && obj[key] !== null) {
      sh.getRange(targetRow, fieldMap[key] + 1).setValue(obj[key]);
    }
  });
  if (obj.joinDate) sh.getRange(targetRow, SP_COL.JOIN_DATE + 1).setValue(parsePaymentDate_(obj.joinDate));
  if (obj.dob)      sh.getRange(targetRow, SP_COL.DOB + 1).setValue(parsePaymentDate_(obj.dob));

  return { success: true, message: "Student details updated successfully." };
}

// ── UPDATE STUDENT FEES (billing, any time after creation) ──
// Notes request: "now enable fee updation details also... after any
// time fee heads could be updated in student profile option." Fees can
// now also be set at Bulk Import time (see bulkAddStudents above); this
// function is for editing them afterward — correcting an amount, adding
// a fee head that wasn't billed initially, etc. `feeAmounts` is
// { headName: { billed, discount } }, same shape Add Student and Bulk
// Import use. Existing rows for a head are updated in place; a head
// with no existing row yet gets a new one — so this works both for
// billing a student for the first time and for correcting amounts later.
function updateStudentFees(token, studentId, feeAmounts) {
  invalidateCaches_();
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  if (!studentId) return { error: "Student ID is required." };
  if (!feeAmounts) return { error: "No fee amounts supplied." };

  var profiles = getAllStudentProfiles_();
  var profile = profiles.filter(function(p) { return p.studentId === studentId; })[0];
  if (!profile) return { error: "Student not found." };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = getSheet(SHEETS.TUITION);
    var data = sh.getDataRange().getValues();
    var existingRowByHead = {};
    for (var i = 1; i < data.length; i++) {
      if (safeStr(data[i][TF_COL.STUDENT_ID]) === studentId) {
        existingRowByHead[safeStr(data[i][TF_COL.FEE_HEAD])] = i + 1; // 1-based sheet row
      }
    }

    var newRows = [];
    Object.keys(feeAmounts).forEach(function(headName) {
      var billedAmt = safeNum(feeAmounts[headName].billed);
      var discountAmt = safeNum(feeAmounts[headName].discount);
      var rowArr = new Array(TF_NUM_COLS).fill("");
      rowArr[TF_COL.STUDENT_ID] = studentId;
      rowArr[TF_COL.NAME]       = profile.name;
      rowArr[TF_COL.CLASS]      = profile.class;
      rowArr[TF_COL.SECTION]    = profile.section || "";
      rowArr[TF_COL.FEE_HEAD]   = headName;
      rowArr[TF_COL.BILLED]     = billedAmt;
      rowArr[TF_COL.DISCOUNT]   = discountAmt;
      rowArr[TF_COL.NET]        = billedAmt - discountAmt;
      rowArr[TF_COL.ACADEMIC_YEAR] = getAcademicYear_();

      var existingRow = existingRowByHead[headName];
      if (existingRow) {
        sh.getRange(existingRow, 1, 1, TF_NUM_COLS).setValues([rowArr]);
      } else {
        newRows.push(rowArr);
      }
    });
    if (newRows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, TF_NUM_COLS).setValues(newRows);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return { success: true, message: "Fee details updated for " + profile.name + "." };
}

// ── DELETE STUDENT ───────────────────────────────────────────
// Notes request: a way to remove duplicate entries or students who've
// taken a Transfer Certificate, from Edit Student. This is a SOFT
// delete — the Student Profile row's status changes to "Removed" rather
// than the row (or anything in Tution_Fee/Transactions/Marks/Attendance
// that references this Student ID) being physically deleted. That keeps
// every past receipt, mark, and attendance record intact and still
// queryable by ID, while getStudentList() (Student Records, and every
// dropdown built from it) excludes "Removed" students so they actually
// disappear from daily use, matching what was asked for.
// ── DELETE STUDENT → ARCHIVE TO DROPOUTS ────────────────────
// Notes request: removing a student should genuinely shrink the active
// counts (Billed/Received/Outstanding, active student count) rather
// than just hiding them with a status flag — but their history must
// not be lost. This MOVES (not soft-flags) the student's data:
//   - Student Profile row  -> deleted from the live sheet, written into
//     "Dropouts" (same columns + Dropout Date + Reason)
//   - Every Tution_Fee row for them -> moved into "Dropout_Fee_History"
//   - Every Transactions row for them -> moved into "Dropout_Transactions"
// Because the rows are gone from the LIVE sheets, every dashboard/report
// that reads Student Profile/Tution_Fee/Transactions automatically
// reflects the smaller active totals with no separate "exclude removed"
// filtering needed anywhere else in the codebase.
function deleteStudent(token, studentId, reason) {
  invalidateCaches_();
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!studentId) return { error: "Student ID is required." };
  if (!reason || !reason.trim()) return { error: "A reason is required when removing a student." };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    // 1) Find and remove the Student Profile row, archiving it first.
    var profSh = getSheet(SHEETS.PROFILE);
    var profData = profSh.getDataRange().getValues();
    var profileRowIndex = -1, profileRow = null;
    for (var i = 2; i < profData.length; i++) {
      if (safeStr(profData[i][SP_COL.STUDENT_ID]) === studentId) {
        profileRowIndex = i;
        profileRow = profData[i];
        break;
      }
    }
    if (profileRowIndex === -1) return { error: "Student not found." };

    var dropoutSh = getSheet(SHEETS.DROPOUTS);
    if (dropoutSh.getLastRow() === 0) {
      // Mirrors Student Profile's own header row, plus two archive-only columns.
      var profHeaderRow = profData[1] || [];
      dropoutSh.appendRow(profHeaderRow.concat(["Dropout Date", "Reason"]));
    }
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy");
    dropoutSh.appendRow(profileRow.slice(0, SP_NUM_COLS).concat([stamp, reason.trim()]));
    profSh.deleteRow(profileRowIndex + 1); // +1: 0-based index -> 1-based sheet row

    // 2) Move every Tution_Fee row for this student into Dropout_Fee_History.
    var tSh = getSheet(SHEETS.TUITION);
    var tData = tSh.getDataRange().getValues();
    var tRowsToArchive = [], tRowsToDelete = [];
    for (var k = 1; k < tData.length; k++) {
      if (safeStr(tData[k][TF_COL.STUDENT_ID]) === studentId) {
        tRowsToArchive.push(tData[k]);
        tRowsToDelete.push(k + 1);
      }
    }
    if (tRowsToArchive.length) {
      var dfhSh = getSheet(SHEETS.DROPOUT_FEE_HISTORY);
      if (dfhSh.getLastRow() === 0) {
        dfhSh.appendRow(["Student ID", "Name", "Class", "Section", "Fee Head", "Billed", "Discount", "Net", "Academic Year", "Dropout Date"]);
      }
      var dfhRows = tRowsToArchive.map(function(r) { return r.slice(0, TF_NUM_COLS).concat([stamp]); });
      dfhSh.getRange(dfhSh.getLastRow() + 1, 1, dfhRows.length, TF_NUM_COLS + 1).setValues(dfhRows);
      tRowsToDelete.sort(function(a, b) { return b - a; }).forEach(function(r) { tSh.deleteRow(r); });
    }

    // 3) Mark Transactions rows for this student as "Dropout" (NOT moved),
    // so their collected payments remain in Receipts, P&L, Balance Sheet.
    // We move a COPY into Dropout_Transactions for audit reference, but keep
    // the originals in the live Transactions sheet so income totals are correct.
    var txSh = getSheet(SHEETS.TRANSACTIONS);
    var txData = txSh.getDataRange().getValues();
    var txRowsToArchive = [];
    for (var m = 1; m < txData.length; m++) {
      if (safeStr(txData[m][TX_COL.STUDENT_ID]) === studentId) {
        txRowsToArchive.push(txData[m]);
      }
    }
    if (txRowsToArchive.length) {
      // Write copies to Dropout_Transactions for audit reference
      var dtxSh = getSheet(SHEETS.DROPOUT_TRANSACTIONS);
      if (dtxSh.getLastRow() === 0) {
        dtxSh.appendRow(["Date", "Student ID", "Student Name", "Class", "Receipt", "Fee Head", "Mode", "Amount", "Recorded By", "Voided", "Bank Account", "Dropout Date"]);
      }
      var dtxRows = txRowsToArchive.map(function(r) {
        var padded = r.slice(0, 11);
        while (padded.length < 11) padded.push("");
        return padded.concat([stamp]);
      });
      dtxSh.getRange(dtxSh.getLastRow() + 1, 1, dtxRows.length, 12).setValues(dtxRows);
      // NOTE: We do NOT delete from the main Transactions sheet — this preserves
      // the collected amounts in Receipts Report, P&L, and Balance Sheet.
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return {
    success: true,
    message: "Student moved to Dropouts. They no longer appear in active students or outstanding dues. Their collected fee payments are retained in Receipts, P&L, and Balance Sheet — a 'Dropout Fee Received' line shows their total contribution. Their fee billing is excluded from active totals."
  };
}

// Returns the Dropouts list for a read-only review screen — Notes
// request: keep their details visible somewhere, just out of the active
// data. Pulls profile + a quick fee summary from the archived sheets.
function getDropoutsList(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var dropoutSh = getSheet(SHEETS.DROPOUTS);
  var data = dropoutSh.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };

  // Dropout Date / Reason are appended after the normal profile columns.
  var dropoutDateCol = SP_NUM_COLS, reasonCol = SP_NUM_COLS + 1;

  var dfhSh = getSheet(SHEETS.DROPOUT_FEE_HISTORY);
  var dfhData = dfhSh.getDataRange().getValues();
  var feeByStudent = {}; // studentId -> {billed, paid(from dropout tx), due}
  for (var i = 1; i < dfhData.length; i++) {
    var sid = safeStr(dfhData[i][TF_COL.STUDENT_ID]);
    if (!feeByStudent[sid]) feeByStudent[sid] = { billed: 0, paid: 0 };
    feeByStudent[sid].billed += safeNum(dfhData[i][TF_COL.NET]) || (safeNum(dfhData[i][TF_COL.BILLED]) - safeNum(dfhData[i][TF_COL.DISCOUNT]));
  }
  var dtxSh = getSheet(SHEETS.DROPOUT_TRANSACTIONS);
  var dtxData = dtxSh.getDataRange().getValues();
  for (var k = 1; k < dtxData.length; k++) {
    if (isVoided_(dtxData[k][TX_COL.VOID])) continue;
    var tsid = safeStr(dtxData[k][TX_COL.STUDENT_ID]);
    if (!feeByStudent[tsid]) feeByStudent[tsid] = { billed: 0, paid: 0 };
    feeByStudent[tsid].paid += safeNum(dtxData[k][TX_COL.AMOUNT]);
  }

  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var studentId = safeStr(row[SP_COL.STUDENT_ID]);
    if (!studentId) continue;
    var f = feeByStudent[studentId] || { billed: 0, paid: 0 };
    out.push({
      studentId: studentId,
      name: safeStr(row[SP_COL.NAME]),
      class: safeStr(row[SP_COL.CLASS]),
      section: safeStr(row[SP_COL.SECTION]),
      fatherName: safeStr(row[SP_COL.FATHER_NAME]),
      fatherMobile: safeStr(row[SP_COL.FATHER_MOBILE]),
      dropoutDate: safeStr(row[dropoutDateCol]),
      reason: safeStr(row[reasonCol]),
      billed: f.billed, paid: f.paid, due: f.billed - f.paid
    });
  }
  out.sort(function(a, b) { return (b.dropoutDate || "").localeCompare(a.dropoutDate || ""); });
  return { success: true, data: out };
}

// Dedicated, lightweight call for just a photo update, so the client can
// let Clerk update a photo (e.g. paste a new Drive image URL) without
// re-submitting the whole edit form. Notes item #6.
function updateStudentPhoto(token, studentId, photoUrl) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (!studentId) return { error: "Student ID is required." };

  var sh = getSheet(SHEETS.PROFILE);
  var data = sh.getDataRange().getValues();
  for (var i = 2; i < data.length; i++) {
    if (safeStr(data[i][SP_COL.STUDENT_ID]) === studentId) {
      sh.getRange(i + 1, SP_COL.PHOTO + 1).setValue(photoUrl || "");
      return { success: true, message: "Photo updated." };
    }
  }
  return { error: "Student not found." };
}

// ── RECORD A FEE PAYMENT (writes to Transactions sheet) ─────
// receiptNo is no longer taken from the client — it's generated here so the
// sequence always continues correctly even with concurrent submissions.
// paymentDateStr ("YYYY-MM-DD") lets clerk record a payment against a date
// in the past, e.g. cash collected last week but logged today.
// v3: now records against studentId (not just name/class) so duplicate
// student names never collide — Notes item #7.
function recordFeePayment(token, studentId, studentName, studentClass, feeHead, amount, mode, paymentDateStr, bankAccount) {
  invalidateCaches_();
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };

  amount = safeNum(amount);
  if (amount <= 0) return { error: "Enter a valid amount." };
  if (!feeHead) return { error: "Please select a fee head." };
  if (!studentId) return { error: "Student ID is required." };

  var payDate = parsePaymentDate_(paymentDateStr);
  if (payDate.getTime() > new Date().getTime() + 24 * 60 * 60 * 1000) {
    return { error: "Payment date cannot be in the future." };
  }

  var lock = LockService.getScriptLock();
  var receiptNo;
  var newRowIndex;
  try {
    lock.waitLock(10000); // up to 10s, so two clerk saving at once don't collide
    var txSh = getSheet(SHEETS.TRANSACTIONS);
    if (txSh.getLastRow() === 0) {
      txSh.appendRow(["Date", "Student ID", "Student", "Class", "Receipt", "Fee Head", "Mode", "Amount", "By", "Voided", "Bank Account"]);
    }
    receiptNo = getNextReceiptNo_(payDate, feeHead);
    txSh.appendRow([
      payDate, studentId, studentName, studentClass, receiptNo, feeHead, mode || "Cash", amount, sess.name, false, bankAccount || ""
    ]);
    newRowIndex = txSh.getLastRow() - 1; // 0-based index matching getDataRange().getValues()
  } finally {
    lock.releaseLock();
  }

  // Recompute this student's outstanding for immediate feedback
  var feeMap = getFeeMap_();
  var fee = feeMap[studentId] || { due: 0 };

  return {
    success: true,
    receiptNo: receiptNo,
    studentId: studentId,
    transactionRowIndex: newRowIndex,
    message: "Receipt " + receiptNo + " · ₹" + amount + " recorded against " + feeHead + " for " + studentName + ". Remaining due: ₹" + (fee.due > 0 ? fee.due : 0)
  };
}

// ── PRINTABLE RECEIPT (Parent Copy + Office Copy) ───────────
// Returns everything the client needs to render a printable receipt for
// one transaction: school letterhead, student details, the specific
// payment, and the student's running fee position. Used both right
// after recordFeePayment() succeeds, and later from the transaction
// history (re-print any past receipt on demand).
function getReceiptDetail(token, transactionRowIndex) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var txSh = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();
  var t = txData[Number(transactionRowIndex)];
  if (!t || !t[TX_COL.RECEIPT]) return { error: "Transaction not found." };

  var studentId = safeStr(t[TX_COL.STUDENT_ID]);
  var profiles = getAllStudentProfiles_();
  var p = profiles.filter(function(x) { return x.studentId === studentId; })[0];

  var feeMap = getFeeMap_();
  var fee = feeMap[studentId] || { billed: 0, paid: 0, due: 0 };

  var school = getSchoolInfo(token);

  return {
    success: true,
    school: { name: school.name, address: school.address, phone: school.phone, email: school.email, logo: school.logo },
    receipt: {
      receiptNo:   safeStr(t[TX_COL.RECEIPT]),
      date:        formatDateSafe_(t[TX_COL.DATE]) || safeStr(t[TX_COL.DATE]),
      studentId:   studentId,
      // Prefer what's stored on the transaction itself, but fall back to
      // the student's current profile if that's blank — this covers
      // payments recorded while the profile-screen "Record Payment"
      // button had the bug that sent an empty name/class (see
      // openPayFromProfile() fix), so old receipts print correctly too.
      studentName: safeStr(t[TX_COL.STUDENT_NAME]) || (p ? p.name : ""),
      class:       safeStr(t[TX_COL.CLASS]) || (p ? p.class : ""),
      section:     p ? p.section : "",
      fatherName:  p ? p.fatherName : "",
      feeHead:     safeStr(t[TX_COL.HEAD]),
      amount:      safeNum(t[TX_COL.AMOUNT]),
      mode:        safeStr(t[TX_COL.MODE]),
      bankAccount: safeStr(t[TX_COL.BANK_ACCOUNT]),
      recordedBy:  safeStr(t[TX_COL.BY]),
      voided:      isVoided_(t[TX_COL.VOID])
    },
    feePosition: { billed: fee.billed, paid: fee.paid, due: fee.due > 0 ? fee.due : 0 }
  };
}

// New in v3 — Notes item #6. Marks the Transactions row as voided rather
// than deleting it, so there is always an audit trail of what was
// recorded and who cancelled it. Voided rows are excluded from getFeeMap_().
function cancelPayment(token, rowIndex, reason) {
  invalidateCaches_();
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  if (rowIndex === undefined || rowIndex === null) return { error: "Transaction not specified." };

  var txSh = getSheet(SHEETS.TRANSACTIONS);
  var sheetRow = Number(rowIndex) + 1; // rowIndex is 0-based into getDataRange().getValues()
  if (sheetRow < 2 || sheetRow > txSh.getLastRow()) return { error: "Transaction not found." };

  var existingReceipt = txSh.getRange(sheetRow, TX_COL.RECEIPT + 1).getValue();
  if (!existingReceipt) return { error: "Transaction not found." };

  txSh.getRange(sheetRow, TX_COL.VOID + 1).setValue(true);
  var byCell = safeStr(txSh.getRange(sheetRow, TX_COL.BY + 1).getValue());
  txSh.getRange(sheetRow, TX_COL.BY + 1).setValue(byCell + " (cancelled by " + sess.name + (reason ? ": " + reason : "") + ")");

  return { success: true, message: "Receipt " + existingReceipt + " has been cancelled." };
}

// ── DUES / WHATSAPP REMINDERS ────────────────────────────────
// v3: rebuilt message template — shows a full fee-head-wise breakdown
// (billed / received / balance) instead of a single due figure, so
// parents get full clarity in one message — Notes item #11.
function getWhatsAppReminders(token, schoolName) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var profiles = getAllStudentProfiles_();
  var feeMap   = getFeeMap_();
  var school   = schoolName || SCHOOL_NAME;

  var result = [];
  profiles.forEach(function(p) {
    var f = feeMap[p.studentId];
    if (!f || f.due <= 0) return;
    var msg = buildDueReminderMessage_(school, p, f);
    var phone = p.fatherMobile ? p.fatherMobile.replace(/\D/g, "") : "";
    if (phone.length === 10) phone = "91" + phone;
    result.push({
      studentId: p.studentId, name: p.name, class: p.class, due: f.due,
      phone: phone,
      waURL: "https://wa.me/" + phone + "?text=" + encodeURIComponent(msg)
    });
  });
  return { success: true, data: result };
}

function buildDueReminderMessage_(school, p, f) {
  var lines = [];
  lines.push("Dear Parent,");
  lines.push("");
  lines.push("This is a fee reminder from *" + school + "*.");
  lines.push("");
  lines.push("Student: *" + p.name + "* (" + p.studentId + ")");
  lines.push("Class: " + p.class);
  lines.push("");
  lines.push("*Fee Head-wise Status:*");
  (f.heads ? Object.keys(f.heads) : []).forEach(function(h) {
    var d = f.heads[h];
    if (!d || (d.billed <= 0 && d.paid <= 0)) return;
    var bal = d.billed - d.paid;
    lines.push("• " + h + " — Billed: ₹" + d.billed + " | Received: ₹" + d.paid + " | Balance: " + (bal > 0 ? "₹" + bal : "Nil"));
  });
  lines.push("");
  lines.push("*Total Billed:* ₹" + f.billed);
  lines.push("*Total Received:* ₹" + f.paid);
  lines.push("*Total Outstanding:* ₹" + f.due);
  lines.push("");
  lines.push("Kindly clear the dues at your earliest convenience.");
  lines.push("");
  lines.push("_Powered by Manha E-Solutions_");
  return lines.join("\n");
}

// ── OUTSTANDING DETAILS — Excel/PDF export (Notes request) ──
// Reuses the same outstanding list getWhatsAppReminders() already
// computes (every student with due > 0), exported as a school-wide
// report instead of per-student WhatsApp messages.
function getOutstandingList_(token) {
  var profiles = getAllStudentProfiles_();
  var feeMap   = getFeeMap_();
  var feeHeadRows = getSettingsRows_(SETTINGS_TYPE.FEE_HEAD);
  var allHeads = feeHeadRows.map(function(r) { return r.key; });
  var result = [];
  profiles.forEach(function(p) {
    var f = feeMap[p.studentId];
    if (!f || f.due <= 0) return;
    // Build per-head breakdown for this student
    var headDetails = {};
    allHeads.forEach(function(h) {
      var hd = (f.heads && f.heads[h]) || { billed: 0, paid: 0 };
      headDetails[h] = { billed: hd.billed, paid: hd.paid, due: hd.billed - hd.paid };
    });
    // Also capture any heads that appear in transactions but not in fee-head settings
    if (f.heads) {
      Object.keys(f.heads).forEach(function(h) {
        if (!headDetails[h]) {
          var hd = f.heads[h];
          headDetails[h] = { billed: hd.billed, paid: hd.paid, due: hd.billed - hd.paid };
        }
      });
    }
    result.push({
      studentId: p.studentId, name: p.name, class: p.class,
      section: p.section || '', billed: f.billed, paid: f.paid, due: f.due,
      heads: headDetails
    });
  });
  result.sort(function(a, b) { return b.due - a.due; });
  return result;
}

function exportOutstandingCsv(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var school = getSchoolInfo(token);
  var list = getOutstandingList_(token);
  var asOfDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy");

  // Collect all fee heads that appear across students
  var headSet = {};
  list.forEach(function(s) { Object.keys(s.heads || {}).forEach(function(h) { headSet[h] = true; }); });
  var heads = Object.keys(headSet).sort();

  var lines = [];
  lines.push(csvEscapeRow_([school.name || SCHOOL_NAME]));
  lines.push(csvEscapeRow_(["Outstanding Fee Details — Head-wise — As on " + asOfDateStr]));
  lines.push("");

  // Header row: fixed cols + one col per head (Due only) + totals
  var headerRow = ["Student ID", "Name", "Class", "Section"];
  heads.forEach(function(h) { headerRow.push(h + " (Billed)", h + " (Paid)", h + " (Due)"); });
  headerRow.push("Total Billed", "Total Paid", "Total Outstanding");
  lines.push(csvEscapeRow_(headerRow));

  var colTotals = { billed: 0, paid: 0, due: 0 };
  var headColTotals = {};
  heads.forEach(function(h) { headColTotals[h] = { billed: 0, paid: 0, due: 0 }; });

  list.forEach(function(s) {
    var row = [s.studentId, s.name, s.class, s.section];
    heads.forEach(function(h) {
      var hd = (s.heads && s.heads[h]) || { billed: 0, paid: 0, due: 0 };
      row.push(hd.billed, hd.paid, hd.due);
      headColTotals[h].billed += hd.billed;
      headColTotals[h].paid   += hd.paid;
      headColTotals[h].due    += hd.due;
    });
    row.push(s.billed, s.paid, s.due);
    colTotals.billed += s.billed; colTotals.paid += s.paid; colTotals.due += s.due;
    lines.push(csvEscapeRow_(row));
  });

  // Totals row
  var totRow = ["", "TOTAL", "", ""];
  heads.forEach(function(h) {
    totRow.push(headColTotals[h].billed, headColTotals[h].paid, headColTotals[h].due);
  });
  totRow.push(colTotals.billed, colTotals.paid, colTotals.due);
  lines.push(csvEscapeRow_(totRow));

  return { success: true, csv: lines.join("\r\n"), count: list.length, totalDue: colTotals.due };
}

function getOutstandingListPdf(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var school = getSchoolInfo(token);
  var list   = getOutstandingList_(token);
  var asOfDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy");

  // Union of all heads across students (only heads with any billed or due amount)
  var headSet = {};
  list.forEach(function(s) {
    Object.keys(s.heads || {}).forEach(function(h) {
      if ((s.heads[h].billed || 0) > 0 || (s.heads[h].due || 0) > 0) headSet[h] = true;
    });
  });
  var heads = Object.keys(headSet).sort();
  var colW = heads.length > 0 ? Math.floor(54 / heads.length) + "%" : "18%";

  var totalDue = 0;
  var rowsHtml = "";

  list.forEach(function(s, idx) {
    var bg = idx % 2 === 0 ? "#fff" : "#f9fbfc";
    // head cells
    var headCells = heads.map(function(h) {
      var hd = s.heads[h] || { billed: 0, paid: 0, due: 0 };
      var color = hd.due > 0 ? "#C0392B" : "#1A7A4A";
      return "<td style='text-align:right;padding:5px 7px;font-size:10.5px;color:" + color + ";'>" +
        (hd.due > 0 ? "<b>₹" + fmt_(hd.due) + "</b>" : "<span style='color:#aaa;'>—</span>") + "</td>";
    }).join("");
    rowsHtml +=
      "<tr style='background:" + bg + ";'>" +
        "<td style='padding:5px 7px;font-size:10.5px;'>" + escHtml_(s.name) + "</td>" +
        "<td style='padding:5px 7px;font-size:10.5px;'>" + escHtml_(s.class) + (s.section ? "-" + escHtml_(s.section) : "") + "</td>" +
        headCells +
        "<td style='text-align:right;padding:5px 7px;font-size:10.5px;color:#C0392B;font-weight:bold;'>₹" + fmt_(s.due) + "</td>" +
      "</tr>";
    totalDue += s.due;
  });

  if (!list.length) {
    rowsHtml = "<tr><td colspan='" + (heads.length + 3) + "' style='text-align:center;color:#888;padding:16px;'>No outstanding dues — every student is clear.</td></tr>";
  }

  // Head column headers
  var headHeaders = heads.map(function(h) {
    return "<th style='padding:6px 7px;text-align:right;font-size:10px;'>" + escHtml_(h) + "<br><span style='font-weight:400;opacity:.8;font-size:9px;'>(Due)</span></th>";
  }).join("");

  var html =
    "<div style='font-family:Arial,sans-serif;color:#1A1A1A;'>" +
      "<div style='text-align:center;border-bottom:3px solid #0D2137;padding-bottom:10px;margin-bottom:14px;'>" +
        "<div style='font-size:20px;font-weight:bold;color:#0D2137;'>" + escHtml_(school.name) + "</div>" +
        (school.address ? "<div style='font-size:11px;color:#555;margin-top:3px;'>" + escHtml_(school.address) + "</div>" : "") +
      "</div>" +
      "<div style='text-align:center;font-size:15px;font-weight:bold;color:#0A7E8C;margin-bottom:4px;'>OUTSTANDING FEE DETAILS — HEAD-WISE</div>" +
      "<div style='text-align:center;font-size:11px;color:#666;margin-bottom:16px;'>As on " + escHtml_(asOfDateStr) + " &middot; " + list.length + " student(s) with dues</div>" +
      "<table style='width:100%;border-collapse:collapse;font-size:11px;'>" +
        "<thead><tr style='background:#0D2137;color:#fff;'>" +
          "<th style='padding:6px 7px;text-align:left;'>Name</th>" +
          "<th style='padding:6px 7px;text-align:left;'>Class</th>" +
          headHeaders +
          "<th style='padding:6px 7px;text-align:right;'>Total Due</th>" +
        "</tr></thead>" +
        "<tbody>" + rowsHtml + "</tbody>" +
      "</table>" +
      "<table style='width:100%;border-collapse:collapse;margin-top:16px;'>" +
        "<tr><td style='background:#C0392B;color:#fff;padding:12px;text-align:center;border-radius:4px;'>" +
          "<div style='font-size:11px;'>TOTAL OUTSTANDING ACROSS " + list.length + " STUDENT(S)</div>" +
          "<div style='font-size:18px;font-weight:bold;'>₹" + fmt_(totalDue) + "</div>" +
        "</td></tr>" +
      "</table>" +
      "<div style='margin-top:24px;font-size:10px;color:#888;text-align:center;border-top:1px solid #ddd;padding-top:8px;'>" +
        "Generated via School DBMS &middot; Manha E-Solutions &middot; " + escHtml_(asOfDateStr) +
      "</div>" +
    "</div>";

  var fileName = "Outstanding_Headwise_" + asOfDateStr.replace(/-/g, "");
  var pdfBytes;
  try { pdfBytes = htmlToPdfBlob_(html, fileName); }
  catch (e) { return { error: "Could not generate the PDF: " + e.message }; }

  return { success: true, fileName: fileName + ".pdf", pdfBase64: Utilities.base64Encode(pdfBytes), count: list.length, totalDue: totalDue };
}

// ============================================================
//  OUTSTANDING FEE STATEMENT — PDF (letterhead + fee card + full
//  transaction history "as on <date>"). Notes request: make the
//  WhatsApp outstanding reminder look professional by attaching a real
//  statement instead of a plain text message.
//
//  WhatsApp's free wa.me link can only pre-fill TEXT — it cannot attach
//  a file automatically; that requires paid WhatsApp Business API access
//  most schools don't have. So the flow here is: generate this PDF,
//  trigger a download for the Clerk, then open a short WhatsApp
//  message alongside it so they can attach the PDF with one tap before
//  sending. See getOutstandingStatementPdf() below.
// ============================================================

// Builds the statement's HTML. Deliberately plain, table-based markup —
// this gets inserted into a temporary Google Doc and exported to PDF, so
// only the small subset of CSS that Apps Script's HTML-to-Doc conversion
// actually honors is used (basic colors, borders, padding, alignment).
function buildOutstandingStatementHtml_(school, p, f, history, asOfDateStr) {
  var rowsHtml = history.map(function(t) {
    var statusTxt = t.voided ? "Cancelled" : "";
    var amtTxt = t.voided ? "<s>₹" + fmt_(t.amount) + "</s>" : "₹" + fmt_(t.amount);
    return "<tr>" +
      "<td>" + escHtml_(t.date) + "</td>" +
      "<td>" + escHtml_(t.receipt || "—") + "</td>" +
      "<td>" + escHtml_(t.head) + "</td>" +
      "<td>" + escHtml_(t.mode) + "</td>" +
      "<td style='text-align:right;'>" + amtTxt + "</td>" +
      "<td>" + escHtml_(statusTxt) + "</td>" +
      "</tr>";
  }).join("");
  if (!history.length) {
    rowsHtml = "<tr><td colspan='6' style='text-align:center;color:#888;'>No payments recorded yet.</td></tr>";
  }

  var headRowsHtml = f.heads.map(function(h) {
    if (h.billed <= 0 && h.paid <= 0) return "";
    return "<tr>" +
      "<td>" + escHtml_(h.head) + "</td>" +
      "<td style='text-align:right;'>₹" + fmt_(h.billed) + "</td>" +
      "<td style='text-align:right;color:#1A7A4A;'>₹" + fmt_(h.paid) + "</td>" +
      "<td style='text-align:right;color:#C0392B;'>" + (h.due > 0 ? "₹" + fmt_(h.due) : "Nil") + "</td>" +
      "</tr>";
  }).join("");

  return "" +
    "<div style='font-family:Arial,sans-serif;color:#1A1A1A;'>" +
      "<div style='text-align:center;border-bottom:3px solid #0D2137;padding-bottom:10px;margin-bottom:14px;'>" +
        "<div style='font-size:20px;font-weight:bold;color:#0D2137;'>" + escHtml_(school.name) + "</div>" +
        (school.address ? "<div style='font-size:11px;color:#555;margin-top:3px;'>" + escHtml_(school.address) + "</div>" : "") +
        (school.phone || school.email ? "<div style='font-size:11px;color:#555;margin-top:2px;'>" + [school.phone, school.email].filter(Boolean).map(escHtml_).join(" &nbsp;|&nbsp; ") + "</div>" : "") +
      "</div>" +
      "<div style='text-align:center;font-size:15px;font-weight:bold;color:#0A7E8C;margin-bottom:4px;'>OUTSTANDING FEE STATEMENT</div>" +
      "<div style='text-align:center;font-size:11px;color:#666;margin-bottom:16px;'>As on " + escHtml_(asOfDateStr) + "</div>" +

      "<table style='width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px;'>" +
        "<tr><td style='width:50%;padding:3px 0;color:#666;'>Student Name</td><td style='font-weight:bold;'>" + escHtml_(p.name) + " (" + escHtml_(p.studentId) + ")</td></tr>" +
        "<tr><td style='padding:3px 0;color:#666;'>Class / Section</td><td style='font-weight:bold;'>" + escHtml_(p.class) + (p.section ? " - " + escHtml_(p.section) : "") + "</td></tr>" +
        "<tr><td style='padding:3px 0;color:#666;'>Father's Name</td><td style='font-weight:bold;'>" + escHtml_(p.fatherName || "—") + "</td></tr>" +
        "<tr><td style='padding:3px 0;color:#666;'>Contact No.</td><td style='font-weight:bold;'>" + escHtml_(p.fatherMobile || "—") + "</td></tr>" +
      "</table>" +

      "<div style='font-size:13px;font-weight:bold;color:#0D2137;margin-bottom:6px;'>Fee Head-wise Summary (as on " + escHtml_(asOfDateStr) + ")</div>" +
      "<table style='width:100%;border-collapse:collapse;font-size:11.5px;margin-bottom:10px;'>" +
        "<tr style='background:#0D2137;color:#fff;'><th style='padding:5px 8px;text-align:left;'>Fee Head</th><th style='padding:5px 8px;text-align:right;'>Billed</th><th style='padding:5px 8px;text-align:right;'>Received</th><th style='padding:5px 8px;text-align:right;'>Balance</th></tr>" +
        headRowsHtml +
      "</table>" +

      "<table style='width:100%;border-collapse:collapse;margin-bottom:16px;'>" +
        "<tr>" +
          "<td style='width:33%;background:#1A7A4A;color:#fff;padding:10px;text-align:center;border-radius:4px;'><div style='font-size:10px;'>TOTAL BILLED</div><div style='font-size:16px;font-weight:bold;'>₹" + fmt_(f.billed) + "</div></td>" +
          "<td style='width:1%;'>&nbsp;</td>" +
          "<td style='width:33%;background:#0A7E8C;color:#fff;padding:10px;text-align:center;border-radius:4px;'><div style='font-size:10px;'>TOTAL RECEIVED</div><div style='font-size:16px;font-weight:bold;'>₹" + fmt_(f.paid) + "</div></td>" +
          "<td style='width:1%;'>&nbsp;</td>" +
          "<td style='width:33%;background:#C0392B;color:#fff;padding:10px;text-align:center;border-radius:4px;'><div style='font-size:10px;'>OUTSTANDING</div><div style='font-size:16px;font-weight:bold;'>₹" + fmt_(f.due) + "</div></td>" +
        "</tr>" +
      "</table>" +

      "<div style='font-size:13px;font-weight:bold;color:#0D2137;margin-bottom:6px;'>Transaction History</div>" +
      "<table style='width:100%;border-collapse:collapse;font-size:11px;'>" +
        "<tr style='background:#F3F7FA;'><th style='padding:5px 8px;text-align:left;border-bottom:1px solid #ccc;'>Date</th><th style='padding:5px 8px;text-align:left;border-bottom:1px solid #ccc;'>Receipt</th><th style='padding:5px 8px;text-align:left;border-bottom:1px solid #ccc;'>Fee Head</th><th style='padding:5px 8px;text-align:left;border-bottom:1px solid #ccc;'>Mode</th><th style='padding:5px 8px;text-align:right;border-bottom:1px solid #ccc;'>Amount</th><th style='padding:5px 8px;text-align:left;border-bottom:1px solid #ccc;'>Status</th></tr>" +
        rowsHtml +
      "</table>" +

      "<div style='margin-top:24px;font-size:10px;color:#888;text-align:center;border-top:1px solid #ddd;padding-top:8px;'>" +
        "This is a system-generated statement as on " + escHtml_(asOfDateStr) + ". For any discrepancy, please contact the school office.<br>Generated via School DBMS &middot; Manha E-Solutions" +
      "</div>" +
    "</div>";
}

function fmt_(n) { n = safeNum(n); return n.toLocaleString("en-IN"); }
function escHtml_(s) {
  return safeStr(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Converts an HTML string into real PDF bytes by writing it into a
// temporary Google Doc (Apps Script's HTML-to-Doc importer handles basic
// tables/colors/borders correctly) and exporting that Doc as PDF, then
// deleting the temporary file immediately afterward. This is the
// standard reliable way to produce a styled PDF from Apps Script without
// any external service.
function htmlToPdfBlob_(html, fileName) {
  if (typeof Drive === "undefined" || !Drive.Files) {
    throw new Error("The Drive Advanced Service isn't enabled for this script yet. In the Apps Script editor: Services (+) → find \"Drive API\" → Add. Then try again.");
  }
  var tempDoc = null;
  try {
    var htmlBlob = Utilities.newBlob(html, "text/html", "temp.html");
    var tempFile = Drive.Files.insert(
      { title: fileName, mimeType: "application/vnd.google-apps.document" },
      htmlBlob,
      { convert: true }
    );
    tempDoc = DriveApp.getFileById(tempFile.id);
    var pdfBlob = tempDoc.getAs("application/pdf").setName(fileName + ".pdf");
    var bytes = pdfBlob.getBytes();
    return bytes;
  } finally {
    if (tempDoc) {
      try { tempDoc.setTrashed(true); } catch (e) { /* best effort cleanup */ }
    }
  }
}

// Builds the outstanding statement PDF for one student "as on" today,
// returns it as base64 for the client to download, plus a short
// WhatsApp message (the full breakdown now lives in the PDF, not the
// chat text) so Clerk can attach the statement and send both together.
function getOutstandingStatementPdf(token, studentId) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var detail = getStudentProfileDetail(token, studentId);
  if (detail.error) return detail;
  var p = detail.profile, f = detail.fee, history = detail.transactions;

  var school = getSchoolInfo(token);
  var asOfDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy");

  var html = buildOutstandingStatementHtml_(school, p, f, history, asOfDateStr);
  var fileName = "Fee_Statement_" + p.studentId + "_" + asOfDateStr.replace(/-/g, "");

  var pdfBytes;
  try {
    pdfBytes = htmlToPdfBlob_(html, fileName);
  } catch (e) {
    return { error: "Could not generate the PDF: " + e.message };
  }

  var phone = p.fatherMobile ? p.fatherMobile.replace(/\D/g, "") : "";
  if (phone.length === 10) phone = "91" + phone;

  var msgLines = [];
  msgLines.push("Dear Parent,");
  msgLines.push("");
  msgLines.push("Please find attached the outstanding fee statement for *" + p.name + "* (" + p.studentId + ", Class " + p.class + ") as on " + asOfDateStr + ".");
  msgLines.push("");
  msgLines.push("*Total Outstanding: ₹" + fmt_(f.due) + "*");
  msgLines.push("");
  msgLines.push("Kindly clear the dues at your earliest convenience. For any queries, please contact the school office.");
  msgLines.push("");
  msgLines.push("_" + school.name + "_");
  var message = msgLines.join("\n");

  return {
    success: true,
    fileName: fileName + ".pdf",
    pdfBase64: Utilities.base64Encode(pdfBytes),
    waURL: "https://wa.me/" + phone + "?text=" + encodeURIComponent(message),
    waMessage: message,
    phone: phone
  };
}

// New in v3 — payment-received confirmation reminder (Notes item #11).
// Called right after recordFeePayment() succeeds, or on demand from the
// Student Profile transaction history, using the transactionRowIndex
// returned by recordFeePayment() / stored against each history row.
function getPaymentConfirmation(token, studentId, transactionRowIndex) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var profiles = getAllStudentProfiles_();
  var p = profiles.filter(function(x) { return x.studentId === studentId; })[0];
  if (!p) return { error: "Student not found." };

  var txSh = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();
  var t = txData[Number(transactionRowIndex)];
  if (!t) return { error: "Transaction not found." };

  var feeMap = getFeeMap_();
  var f = feeMap[studentId] || { billed: 0, paid: 0, due: 0 };

  var lines = [];
  lines.push("Dear Parent,");
  lines.push("");
  lines.push("We have received your payment at *" + SCHOOL_NAME + "*. Thank you!");
  lines.push("");
  lines.push("Student: *" + p.name + "* (" + p.studentId + ")");
  lines.push("Class: " + p.class);
  lines.push("Receipt No: *" + safeStr(t[TX_COL.RECEIPT]) + "*");
  lines.push("Fee Head: " + safeStr(t[TX_COL.HEAD]));
  lines.push("Date Received: " + formatDateSafe_(t[TX_COL.DATE]));
  lines.push("Amount Received: *₹" + safeNum(t[TX_COL.AMOUNT]) + "*");
  lines.push("");
  lines.push("*Total Received to Date:* ₹" + f.paid);
  lines.push("*Total Outstanding Balance:* " + (f.due > 0 ? "₹" + f.due : "Nil — fully paid"));
  lines.push("");
  lines.push("_Powered by Manha E-Solutions_");
  var msg = lines.join("\n");

  var phone = p.fatherMobile ? p.fatherMobile.replace(/\D/g, "") : "";
  if (phone.length === 10) phone = "91" + phone;

  return {
    success: true,
    waURL: "https://wa.me/" + phone + "?text=" + encodeURIComponent(msg),
    phone: phone
  };
}

// ── STAFF ────────────────────────────────────────────────────
// ============================================================
//  TEACHER PROFILE — rebuilt to mirror Student Profile: full info
//  capture, add/edit from inside the app, auto-generated Teacher ID.
//  Replaces the old read-only Staff Details sheet going forward; the
//  original Staff Details sheet is left untouched as historical record,
//  and its rows are copied into Teacher Profile once automatically the
//  first time this runs (see migrateStaffDetailsToTeacherProfile_()).
// ============================================================

function getNextTeacherId_() {
  var sh = getSheet(SHEETS.TEACHERS);
  var data = sh.getDataRange().getValues();
  var maxSeq = 0;
  for (var i = 1; i < data.length; i++) {
    var id = safeStr(data[i][TEACHER_COL.TEACHER_ID]);
    if (id.indexOf(TEACHER_ID_PREFIX) === 0) {
      var seq = parseInt(id.substring(TEACHER_ID_PREFIX.length), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  return TEACHER_ID_PREFIX + ("0000" + (maxSeq + 1)).slice(-4);
}

// One-time import: if Teacher Profile is empty but the old Staff Details
// sheet has rows, copy them across (assigning each a new Teacher ID) so
// existing staff records aren't lost when switching to the new system.
// Runs automatically inside getTeacherList(); safe to call repeatedly —
// it does nothing once Teacher Profile already has at least one row.
function migrateStaffDetailsToTeacherProfile_() {
  var teachSh = getSheet(SHEETS.TEACHERS);
  if (teachSh.getLastRow() > 0) return; // already migrated (or already in use)

  var staffSh = _getSS_().getSheetByName(SHEETS.STAFF);
  teachSh.appendRow([
    "Teacher ID", "Name of the Employee", "Aadhar No", "PAN No", "National Code", "Gender", "DOB",
    "Father's Name", "Mother's Name", "Caste", "Sub Caste", "Marital Status", "Contact No", "Email",
    "Address", "Academic Qualification", "Professional Qualification", "Experience", "Department",
    "School Type", "Subject", "Designation", "DOJ", "Account No", "Name of the Bank", "IFSC", "Branch",
    "U-Dise Status", "Photo", "UAN", "ESI", "ABHA", "Nominee - Father", "Nominee - Mother",
    "Nominee - Spouse", "Status", "Remarks"
  ]);
  if (!staffSh) return; // no old sheet to migrate from — fine, just starts empty

  var staffData = staffSh.getDataRange().getValues();
  var rowsToWrite = [];
  for (var i = 4; i < staffData.length; i++) { // old sheet: header at row 4 (index 3), data from row 5 (index 4)
    var r = staffData[i];
    if (!r[0]) continue;
    var teacherId = TEACHER_ID_PREFIX + ("0000" + (rowsToWrite.length + 1)).slice(-4);
    rowsToWrite.push([
      teacherId, safeStr(r[0]), safeStr(r[1]), safeStr(r[2]), safeStr(r[3]), safeStr(r[4]),
      r[5] || "", safeStr(r[6]), safeStr(r[7]), safeStr(r[8]), safeStr(r[9]), safeStr(r[10]),
      safeStr(r[11]), safeStr(r[12]), safeStr(r[13]), safeStr(r[14]), safeStr(r[15]), safeStr(r[16]),
      safeStr(r[17]), safeStr(r[18]), safeStr(r[19]), safeStr(r[20]), r[21] || "", safeStr(r[22]),
      safeStr(r[23]), safeStr(r[24]), safeStr(r[25]), safeStr(r[26]), safeStr(r[27]), safeStr(r[28]),
      safeStr(r[29]), safeStr(r[30]), safeStr(r[31]), safeStr(r[32]), safeStr(r[33]), "Active", ""
    ]);
  }
  if (rowsToWrite.length) {
    teachSh.getRange(2, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
  }
}

function getAllTeacherProfiles_() {
  migrateStaffDetailsToTeacherProfile_();
  var sh = getSheet(SHEETS.TEACHERS);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[TEACHER_COL.NAME] || r[TEACHER_COL.NAME] === "") continue;
    out.push({
      rowIndex:          i,
      teacherId:         safeStr(r[TEACHER_COL.TEACHER_ID]),
      name:              safeStr(r[TEACHER_COL.NAME]),
      aadhar:            safeStr(r[TEACHER_COL.AADHAR]),
      pan:               safeStr(r[TEACHER_COL.PAN]),
      nationalCode:      safeStr(r[TEACHER_COL.NATIONAL_CODE]),
      gender:            safeStr(r[TEACHER_COL.GENDER]),
      dob:               formatDateSafe_(r[TEACHER_COL.DOB]),
      fatherName:        safeStr(r[TEACHER_COL.FATHER_NAME]),
      motherName:        safeStr(r[TEACHER_COL.MOTHER_NAME]),
      caste:             safeStr(r[TEACHER_COL.CASTE]),
      subCaste:          safeStr(r[TEACHER_COL.SUBCASTE]),
      maritalStatus:     safeStr(r[TEACHER_COL.MARITAL_STATUS]),
      contactNo:         safeStr(r[TEACHER_COL.CONTACT_NO]),
      email:             safeStr(r[TEACHER_COL.EMAIL]),
      address:           safeStr(r[TEACHER_COL.ADDRESS]),
      academicQual:      safeStr(r[TEACHER_COL.ACADEMIC_QUAL]),
      professionalQual:  safeStr(r[TEACHER_COL.PROFESSIONAL_QUAL]),
      experience:        safeStr(r[TEACHER_COL.EXPERIENCE]),
      department:        safeStr(r[TEACHER_COL.DEPARTMENT]),
      schoolType:        safeStr(r[TEACHER_COL.SCHOOL_TYPE]),
      subject:           safeStr(r[TEACHER_COL.SUBJECT]),
      designation:       safeStr(r[TEACHER_COL.DESIGNATION]),
      doj:               formatDateSafe_(r[TEACHER_COL.DOJ]),
      bankAcc:           safeStr(r[TEACHER_COL.BANK_ACC]),
      bankName:          safeStr(r[TEACHER_COL.BANK_NAME]),
      ifsc:              safeStr(r[TEACHER_COL.IFSC]),
      branch:            safeStr(r[TEACHER_COL.BRANCH]),
      udise:             safeStr(r[TEACHER_COL.UDISE]),
      photo:             safeStr(r[TEACHER_COL.PHOTO]).indexOf("#VALUE") === 0 ? "" : safeStr(r[TEACHER_COL.PHOTO]),
      uan:               safeStr(r[TEACHER_COL.UAN]),
      esi:               safeStr(r[TEACHER_COL.ESI]),
      abha:              safeStr(r[TEACHER_COL.ABHA]),
      nomineeFather:     safeStr(r[TEACHER_COL.NOMINEE_FATHER]),
      nomineeMother:     safeStr(r[TEACHER_COL.NOMINEE_MOTHER]),
      nomineeSpouse:     safeStr(r[TEACHER_COL.NOMINEE_SPOUSE]),
      status:            safeStr(r[TEACHER_COL.STATUS]) || "Active",
      remarks:           safeStr(r[TEACHER_COL.REMARKS])
    });
  }
  return out;
}

function getTeacherList(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var teachers = getAllTeacherProfiles_();
  return { success: true, data: teachers.map(function(t) {
    return {
      teacherId: t.teacherId, name: t.name, designation: t.designation, department: t.department,
      subject: t.subject, contactNo: t.contactNo, email: t.email, status: t.status, photo: t.photo
    };
  }) };
}

function getTeacherProfileDetail(token, teacherId) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var teachers = getAllTeacherProfiles_();
  var t = teachers.filter(function(x) { return x.teacherId === teacherId; })[0];
  if (!t) return { error: "Teacher not found." };
  return { success: true, profile: t };
}

// Adds a new teacher. New teachers are written straight into the Teacher
// Profile sheet — the same sheet every other screen reads from — so a
// newly added teacher appears immediately everywhere, with no separate
// "sync" step.
function addTeacher(token, obj) {
  var sess = validateSession(token);
  var permErr = requireManagement_(sess);
  if (permErr) return permErr;
  if (!obj || !obj.name) return { error: "Teacher name is required." };

  var lock = LockService.getScriptLock();
  var teacherId;
  try {
    lock.waitLock(10000);
    migrateStaffDetailsToTeacherProfile_();
    teacherId = getNextTeacherId_();

    var sh  = getSheet(SHEETS.TEACHERS);
    var row = sh.getLastRow() + 1;
    var arr = new Array(TEACHER_NUM_COLS).fill("");

    arr[TEACHER_COL.TEACHER_ID]        = teacherId;
    arr[TEACHER_COL.NAME]              = obj.name;
    arr[TEACHER_COL.AADHAR]            = obj.aadhar || "";
    arr[TEACHER_COL.PAN]               = obj.pan || "";
    arr[TEACHER_COL.NATIONAL_CODE]     = obj.nationalCode || "";
    arr[TEACHER_COL.GENDER]            = obj.gender || "";
    arr[TEACHER_COL.DOB]               = obj.dob ? parsePaymentDate_(obj.dob) : "";
    arr[TEACHER_COL.FATHER_NAME]       = obj.fatherName || "";
    arr[TEACHER_COL.MOTHER_NAME]       = obj.motherName || "";
    arr[TEACHER_COL.CASTE]             = obj.caste || "";
    arr[TEACHER_COL.SUBCASTE]          = obj.subCaste || "";
    arr[TEACHER_COL.MARITAL_STATUS]    = obj.maritalStatus || "";
    arr[TEACHER_COL.CONTACT_NO]        = obj.contactNo || "";
    arr[TEACHER_COL.EMAIL]             = obj.email || "";
    arr[TEACHER_COL.ADDRESS]           = obj.address || "";
    arr[TEACHER_COL.ACADEMIC_QUAL]     = obj.academicQual || "";
    arr[TEACHER_COL.PROFESSIONAL_QUAL] = obj.professionalQual || "";
    arr[TEACHER_COL.EXPERIENCE]        = obj.experience || "";
    arr[TEACHER_COL.DEPARTMENT]        = obj.department || "";
    arr[TEACHER_COL.SCHOOL_TYPE]       = obj.schoolType || "";
    arr[TEACHER_COL.SUBJECT]           = obj.subject || "";
    arr[TEACHER_COL.DESIGNATION]       = obj.designation || "";
    arr[TEACHER_COL.DOJ]               = obj.doj ? parsePaymentDate_(obj.doj) : new Date();
    arr[TEACHER_COL.BANK_ACC]          = obj.bankAcc || "";
    arr[TEACHER_COL.BANK_NAME]         = obj.bankName || "";
    arr[TEACHER_COL.IFSC]              = obj.ifsc || "";
    arr[TEACHER_COL.BRANCH]            = obj.branch || "";
    arr[TEACHER_COL.UDISE]             = obj.udise || "";
    arr[TEACHER_COL.PHOTO]             = obj.photo || "";
    arr[TEACHER_COL.UAN]               = obj.uan || "";
    arr[TEACHER_COL.ESI]               = obj.esi || "";
    arr[TEACHER_COL.ABHA]              = obj.abha || "";
    arr[TEACHER_COL.NOMINEE_FATHER]    = obj.nomineeFather || "";
    arr[TEACHER_COL.NOMINEE_MOTHER]    = obj.nomineeMother || "";
    arr[TEACHER_COL.NOMINEE_SPOUSE]    = obj.nomineeSpouse || "";
    arr[TEACHER_COL.STATUS]            = "Active";
    arr[TEACHER_COL.REMARKS]           = obj.remarks || "";

    sh.getRange(row, 1, 1, arr.length).setValues([arr]);
  } finally {
    lock.releaseLock();
  }

  return { success: true, teacherId: teacherId, message: "Teacher '" + obj.name + "' added successfully with ID " + teacherId + "." };
}

// Edits an existing teacher's profile in place by Teacher ID. Same
// partial-update pattern as updateStudent() — only fields actually
// supplied are overwritten.
function updateTeacher(token, teacherId, obj) {
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  if (!teacherId) return { error: "Teacher ID is required." };

  var sh = getSheet(SHEETS.TEACHERS);
  var data = sh.getDataRange().getValues();
  var targetRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][TEACHER_COL.TEACHER_ID]) === teacherId) { targetRow = i + 1; break; }
  }
  if (targetRow === -1) return { error: "Teacher not found." };

  var fieldMap = {
    name: TEACHER_COL.NAME, aadhar: TEACHER_COL.AADHAR, pan: TEACHER_COL.PAN,
    nationalCode: TEACHER_COL.NATIONAL_CODE, gender: TEACHER_COL.GENDER,
    fatherName: TEACHER_COL.FATHER_NAME, motherName: TEACHER_COL.MOTHER_NAME,
    caste: TEACHER_COL.CASTE, subCaste: TEACHER_COL.SUBCASTE, maritalStatus: TEACHER_COL.MARITAL_STATUS,
    contactNo: TEACHER_COL.CONTACT_NO, email: TEACHER_COL.EMAIL, address: TEACHER_COL.ADDRESS,
    academicQual: TEACHER_COL.ACADEMIC_QUAL, professionalQual: TEACHER_COL.PROFESSIONAL_QUAL,
    experience: TEACHER_COL.EXPERIENCE, department: TEACHER_COL.DEPARTMENT, schoolType: TEACHER_COL.SCHOOL_TYPE,
    subject: TEACHER_COL.SUBJECT, designation: TEACHER_COL.DESIGNATION, bankAcc: TEACHER_COL.BANK_ACC,
    bankName: TEACHER_COL.BANK_NAME, ifsc: TEACHER_COL.IFSC, branch: TEACHER_COL.BRANCH,
    udise: TEACHER_COL.UDISE, uan: TEACHER_COL.UAN, esi: TEACHER_COL.ESI, abha: TEACHER_COL.ABHA,
    nomineeFather: TEACHER_COL.NOMINEE_FATHER, nomineeMother: TEACHER_COL.NOMINEE_MOTHER,
    nomineeSpouse: TEACHER_COL.NOMINEE_SPOUSE, status: TEACHER_COL.STATUS, remarks: TEACHER_COL.REMARKS
  };
  Object.keys(fieldMap).forEach(function(key) {
    if (obj[key] !== undefined && obj[key] !== null) {
      sh.getRange(targetRow, fieldMap[key] + 1).setValue(obj[key]);
    }
  });
  if (obj.dob) sh.getRange(targetRow, TEACHER_COL.DOB + 1).setValue(parsePaymentDate_(obj.dob));
  if (obj.doj) sh.getRange(targetRow, TEACHER_COL.DOJ + 1).setValue(parsePaymentDate_(obj.doj));

  return { success: true, message: "Teacher details updated successfully." };
}

function updateTeacherPhoto(token, teacherId, photoUrl) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (!teacherId) return { error: "Teacher ID is required." };

  var sh = getSheet(SHEETS.TEACHERS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][TEACHER_COL.TEACHER_ID]) === teacherId) {
      sh.getRange(i + 1, TEACHER_COL.PHOTO + 1).setValue(photoUrl || "");
      return { success: true, message: "Photo updated." };
    }
  }
  return { error: "Teacher not found." };
}

// ============================================================
//  STAFF ATTENDANCE — daily Present/Absent/Leave marking, with a
//  monthly summary (days present, leaves taken) for payroll.
// ============================================================

// Returns every teacher with their attendance status for one specific
// date, defaulting to "Present" for anyone not yet marked that day —
// this is what the daily marking screen displays and lets Clerk tap to
// change.
function getStaffAttendanceForDate(token, dateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var teachers = getAllTeacherProfiles_().filter(function(t) { return t.status === "Active"; });
  var date = parsePaymentDate_(dateStr);
  var dayInfo = getDayHolidayInfo_(date);

  var sh = getSheet(SHEETS.STAFF_ATTENDANCE);
  var data = sh.getDataRange().getValues();
  var marked = {}; // teacherId -> {status, remarks}
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var d = r[SATT_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (!isSameDay_(d, date)) continue;
    marked[safeStr(r[SATT_COL.TEACHER_ID])] = { status: safeStr(r[SATT_COL.STATUS]), remarks: safeStr(r[SATT_COL.REMARKS]) };
  }

  return {
    success: true,
    date: dateStr,
    isHoliday: dayInfo.isHoliday,
    holidayLabel: dayInfo.label,
    dayName: dayInfo.dayName,
    data: teachers.map(function(t) {
      var m = marked[t.teacherId];
      return {
        teacherId: t.teacherId, name: t.name, designation: t.designation,
        status: m ? m.status : "Present", remarks: m ? m.remarks : "", alreadyMarked: !!m
      };
    })
  };
}

function isSameDay_(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
}

// Saves one day's attendance for every teacher in one call. `entries` is
// an array of { teacherId, name, status, remarks }. Re-marking the same
// date replaces that day's existing rows rather than appending
// duplicates, so this can be safely re-saved if corrected later.
function saveStaffAttendance(token, dateStr, entries) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (!dateStr || !entries || !entries.length) return { error: "Nothing to save." };

  var date = parsePaymentDate_(dateStr);
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = getSheet(SHEETS.STAFF_ATTENDANCE);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["Date", "Teacher ID", "Name", "Status", "Remarks", "Marked By"]);
    }

    // Remove any existing rows for this exact date first (re-mark = replace).
    var data = sh.getDataRange().getValues();
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      var d = data[i][SATT_COL.DATE];
      if (Object.prototype.toString.call(d) === "[object Date]" && isSameDay_(d, date)) {
        rowsToDelete.push(i + 1); // 1-based sheet row
      }
    }
    // Delete from the bottom up so row numbers above aren't shifted mid-loop.
    rowsToDelete.sort(function(a, b) { return b - a; }).forEach(function(r) { sh.deleteRow(r); });

    var newRows = entries.map(function(e) {
      return [date, e.teacherId, e.name, e.status, e.remarks || "", sess.name];
    });
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, SATT_NUM_COLS).setValues(newRows);
  } finally {
    lock.releaseLock();
  }

  return { success: true, message: "Attendance saved for " + entries.length + " staff member(s)." };
}

// Monthly summary per teacher for payroll: days present, days absent,
// leaves taken, and a working-days denominator (Sundays excluded by
// default — see countWorkingDays_).
function getStaffAttendanceSummary(token, month, year) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var teachers = getAllTeacherProfiles_().filter(function(t) { return t.status === "Active"; });

  var sh = getSheet(SHEETS.STAFF_ATTENDANCE);
  var data = sh.getDataRange().getValues();
  var counts = {}; // teacherId -> {present, absent, leave}
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var d = r[SATT_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (d.getMonth() + 1 !== Number(month) || d.getFullYear() !== Number(year)) continue;
    var tid = safeStr(r[SATT_COL.TEACHER_ID]);
    if (!counts[tid]) counts[tid] = { present: 0, absent: 0, leave: 0 };
    var status = safeStr(r[SATT_COL.STATUS]);
    if (status === "Present") counts[tid].present++;
    else if (status === "Absent") counts[tid].absent++;
    else if (status === "Leave") counts[tid].leave++;
  }

  var workingDays = countWorkingDaysInMonth_(Number(month), Number(year));

  return {
    success: true,
    month: month, year: year, workingDays: workingDays,
    data: teachers.map(function(t) {
      var c = counts[t.teacherId] || { present: 0, absent: 0, leave: 0 };
      return {
        teacherId: t.teacherId, name: t.name, designation: t.designation,
        present: c.present, absent: c.absent, leave: c.leave,
        markedDays: c.present + c.absent + c.leave,
        workingDays: workingDays
      };
    })
  };
}

// Counts weekdays (Mon-Sat) in a given month/year — Sundays are treated
// as non-working by default. This is only used as a denominator for
// display; actual present/absent/leave counts always come from real
// marked attendance rows, never assumed.
function countWorkingDaysInMonth_(month, year) {
  return countWorkingDaysInMonthCal_(Number(month), Number(year));
}

// ============================================================
//  ACADEMIC CALENDAR — admin-defined Holidays, Exam days, Special
//  occasions, and explicit Working-day overrides. Stored in the
//  Settings sheet (TYPE = AcademicCalendar): KEY=start date (yyyy-MM-dd),
//  V1=category (Holiday|Exam|Special|Working), V2=title, V3=end date
//  (yyyy-MM-dd, blank = single day). Drives the "working days" used by
//  both staff and student attendance, so holidays are excluded
//  automatically and a working Sunday can be added back.
// ============================================================

var CAL_CATEGORIES = ["Holiday", "Exam", "Special", "Working"];

// Iterate each calendar date from startStr..endStr inclusive, calling fn(dateObj, ymdStr).
function eachDateInRange_(startStr, endStr, fn) {
  var s = parsePaymentDate_(startStr);
  var e = endStr ? parsePaymentDate_(endStr) : s;
  if (e.getTime() < s.getTime()) { var t = s; s = e; e = t; }
  var guard = 0;
  var d = new Date(s.getFullYear(), s.getMonth(), s.getDate());
  while (d.getTime() <= e.getTime() && guard < 1000) {
    fn(new Date(d.getTime()), _ymd_(d));
    d.setDate(d.getDate() + 1);
    guard++;
  }
}

// Normalize a cell value that Sheets may have auto-converted from a
// date string ("2025-10-02") into a Date object on storage.
// getValues() returns Date objects for date-formatted cells, which
// String() turns into "Thu Oct 02 2025 00:00:00 GMT+0530..." and
// breaks every regex-based date parser. Always run calendar date
// cells through this before passing to parsePaymentDate_ or _ymd_.
function normCalDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return _ymd_(v);
  var s = String(v);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? (m[1] + '-' + m[2] + '-' + m[3]) : '';
}

// Build { 'yyyy-MM-dd': [ {category, title}, ... ] } from all calendar rows.
// Reads the sheet directly (not via getSettingsRows_) so we can apply
// normCalDate_() to the KEY and VALUE3 columns — Sheets auto-converts
// date-string entries to Date objects, which safeStr() mangles.
function buildCalendarMap_() {
  ensureSettingsSeeded_();
  // Use the request-scoped Settings cache; avoids a second getDataRange() call
  // when buildCalendarMap_ is invoked alongside getWorkingDaysInfo or attendance.
  if (!_cachedSettings_) {
    _cachedSettings_ = getSheet(SHEETS.SETTINGS).getDataRange().getValues();
  }
  var data = _cachedSettings_;
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (safeStr(r[SET_COL.TYPE]) !== SETTINGS_TYPE.ACADEMIC_CALENDAR) continue;
    if (r[SET_COL.ACTIVE] === false || safeStr(r[SET_COL.ACTIVE]).toUpperCase() === 'FALSE') continue;
    var startStr = normCalDate_(r[SET_COL.KEY]);
    var endStr   = normCalDate_(r[SET_COL.VALUE3]);
    var category = safeStr(r[SET_COL.VALUE1]);
    var title    = safeStr(r[SET_COL.VALUE2]);
    if (!startStr) continue;
    eachDateInRange_(startStr, endStr, function(cat, ttl) {
      return function(d, ymd) {
        if (!map[ymd]) map[ymd] = [];
        map[ymd].push({ category: cat, title: ttl });
      };
    }(category, title));
  }
  return map;
}

// A day is a working day if: explicitly marked "Working" (override),
// OR it is Mon-Sat AND not marked "Holiday". Sundays are off by default.
function isWorkingDay_(dateObj, calMap) {
  var entries = calMap[_ymd_(dateObj)] || [];
  var hasWorking = entries.some(function(e) { return e.category === "Working"; });
  if (hasWorking) return true;
  var hasHoliday = entries.some(function(e) { return e.category === "Holiday"; });
  if (hasHoliday) return false;
  return dateObj.getDay() !== 0; // 0 = Sunday
}

var DAY_NAMES_ = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

// Used by the attendance screens to show a holiday/Sunday banner instead
// of (or above) the marking grid for a given date. Priority: an explicit
// "Working" override always wins (e.g. a working Sunday), then any
// "Holiday" calendar entry (Second Saturday, festival, etc — title is
// shown verbatim), then a plain Sunday.
function getDayHolidayInfo_(dateObj, calMap) {
  calMap = calMap || buildCalendarMap_();
  var dayName = DAY_NAMES_[dateObj.getDay()];
  var entries = calMap[_ymd_(dateObj)] || [];

  var working = entries.filter(function(e) { return e.category === "Working"; })[0];
  if (working) {
    return { isHoliday: false, label: "Working day override — " + working.title, dayName: dayName };
  }

  var holiday = entries.filter(function(e) { return e.category === "Holiday"; })[0];
  if (holiday) {
    return { isHoliday: true, label: holiday.title || "Holiday", dayName: dayName };
  }

  if (dateObj.getDay() === 0) {
    return { isHoliday: true, label: "Sunday", dayName: dayName };
  }

  return { isHoliday: false, label: "", dayName: dayName };
}

// Expands a recurring rule (e.g. "every 2nd Saturday") into individual
// Holiday rows across a date range, so the admin doesn't have to add
// each occurrence by hand. nthOccurrence: 1-4 for the Nth weekday of the
// month, or 0 for "last" weekday of the month. dayOfWeek: 0=Sun..6=Sat.
function addRecurringCalendarHoliday(token, fromDate, toDate, dayOfWeek, nthOccurrence, title) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  var permErr = requireManagement_(sess); if (permErr) return permErr;
  if (!fromDate || !toDate) return { error: "Pick a from and to date." };
  if (!title) return { error: "Please enter a title / description." };
  dayOfWeek = Number(dayOfWeek); nthOccurrence = Number(nthOccurrence);
  if (isNaN(dayOfWeek) || dayOfWeek < 0 || dayOfWeek > 6) return { error: "Pick a valid weekday." };

  var from = parsePaymentDate_(fromDate), to = parsePaymentDate_(toDate);
  if (to.getTime() < from.getTime()) return { error: "End date cannot be before the start date." };

  var dates = [];
  var cursor = new Date(from.getFullYear(), from.getMonth(), 1);
  var guard = 0;
  while (cursor.getTime() <= to.getTime() && guard < 600) {
    var matches = [];
    var lastDay = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0).getDate();
    for (var d = 1; d <= lastDay; d++) {
      var dt = new Date(cursor.getFullYear(), cursor.getMonth(), d);
      if (dt.getDay() === dayOfWeek) matches.push(dt);
    }
    var picked = nthOccurrence === 0 ? matches[matches.length - 1] : matches[nthOccurrence - 1];
    if (picked && picked.getTime() >= from.getTime() && picked.getTime() <= to.getTime()) {
      dates.push(picked);
    }
    cursor.setMonth(cursor.getMonth() + 1);
    guard++;
  }

  if (!dates.length) return { error: "No matching dates found in that range." };

  var sh = getSheet(SHEETS.SETTINGS);
  dates.forEach(function(dt) {
    sh.appendRow([SETTINGS_TYPE.ACADEMIC_CALENDAR, dt, 'Holiday', title, '', '', '', true]);
  });
  return { success: true, message: title + ' added as a holiday on ' + dates.length + ' date(s).' };
}

function countWorkingDaysInRangeCal_(fromObj, toObj, calMap) {
  calMap = calMap || buildCalendarMap_();
  var count = 0, guard = 0;
  var d = new Date(fromObj.getFullYear(), fromObj.getMonth(), fromObj.getDate());
  while (d.getTime() <= toObj.getTime() && guard < 1000) {
    if (isWorkingDay_(d, calMap)) count++;
    d.setDate(d.getDate() + 1);
    guard++;
  }
  return count;
}

function countWorkingDaysInMonthCal_(month, year, calMap) {
  var first = new Date(year, month - 1, 1);
  var last = new Date(year, month, 0);
  return countWorkingDaysInRangeCal_(first, last, calMap || buildCalendarMap_());
}

// Returns all calendar entries (sorted by date) for the Settings screen.
function getAcademicCalendar(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureSettingsSeeded_();
  if (!_cachedSettings_) {
    _cachedSettings_ = getSheet(SHEETS.SETTINGS).getDataRange().getValues();
  }
  var data = _cachedSettings_;
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (safeStr(r[SET_COL.TYPE]) !== SETTINGS_TYPE.ACADEMIC_CALENDAR) continue;
    if (r[SET_COL.ACTIVE] === false || safeStr(r[SET_COL.ACTIVE]).toUpperCase() === 'FALSE') continue;
    out.push({
      rowIndex: i,
      date:     normCalDate_(r[SET_COL.KEY]),
      endDate:  normCalDate_(r[SET_COL.VALUE3]),
      category: safeStr(r[SET_COL.VALUE1]),
      title:    safeStr(r[SET_COL.VALUE2])
    });
  }
  out.sort(function(a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  return { success: true, data: out };
}

function _validCalCategory_(c) { return CAL_CATEGORIES.indexOf(c) !== -1; }

function addAcademicCalendarEntry(token, date, endDate, category, title) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  var permErr = requireManagement_(sess); if (permErr) return permErr;
  if (!date) return { error: "Please pick a date." };
  if (!_validCalCategory_(category)) return { error: "Pick a valid category." };
  if (!title) return { error: "Please enter a title / description." };
  if (endDate && parsePaymentDate_(endDate).getTime() < parsePaymentDate_(date).getTime()) {
    return { error: "End date cannot be before the start date." };
  }
  var sh = getSheet(SHEETS.SETTINGS);
  // Write start/end as Date objects so Sheets stores them consistently as
  // dates (not as auto-converted strings), making normCalDate_() reliable
  // on the read path. endDate goes to VALUE3 (col 4) — the column
  // buildCalendarMap_ reads. Prior bug: it was written to VALUE4 (col 5).
  var startObj = parsePaymentDate_(date);
  var endObj   = endDate ? parsePaymentDate_(endDate) : '';
  sh.appendRow([SETTINGS_TYPE.ACADEMIC_CALENDAR, startObj, category, title, endObj, '', '', true]);
  return { success: true, message: title + ' added to the academic calendar.' };
}

function updateAcademicCalendarEntry(token, rowIndex, date, endDate, category, title) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  var permErr = requireManagement_(sess); if (permErr) return permErr;
  if (!date) return { error: "Please pick a date." };
  if (!_validCalCategory_(category)) return { error: "Pick a valid category." };
  if (!title) return { error: "Please enter a title / description." };
  if (endDate && parsePaymentDate_(endDate).getTime() < parsePaymentDate_(date).getTime()) {
    return { error: "End date cannot be before the start date." };
  }
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (safeStr(sh.getRange(sheetRow, SET_COL.TYPE + 1).getValue()) !== SETTINGS_TYPE.ACADEMIC_CALENDAR) {
    return { error: "Could not locate that calendar entry. Refresh and try again." };
  }
  sh.getRange(sheetRow, SET_COL.KEY    + 1).setValue(parsePaymentDate_(date));
  sh.getRange(sheetRow, SET_COL.VALUE1 + 1).setValue(category);
  sh.getRange(sheetRow, SET_COL.VALUE2 + 1).setValue(title);
  sh.getRange(sheetRow, SET_COL.VALUE3 + 1).setValue(endDate ? parsePaymentDate_(endDate) : '');
  return { success: true, message: "Calendar entry updated." };
}

function deleteAcademicCalendarEntry(token, rowIndex) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  var permErr = requireManagement_(sess); if (permErr) return permErr;
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (safeStr(sh.getRange(sheetRow, SET_COL.TYPE + 1).getValue()) !== SETTINGS_TYPE.ACADEMIC_CALENDAR) {
    return { error: "Could not locate that calendar entry. Refresh and try again." };
  }
  sh.deleteRow(sheetRow);
  return { success: true, message: "Calendar entry removed." };
}

// Working-days finder + this-month event lists, for the calendar screen.
function getWorkingDaysInfo(token, month, year) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  month = Number(month); year = Number(year);
  var calMap = buildCalendarMap_();
  var workingDays = countWorkingDaysInMonthCal_(month, year, calMap);
  var totalDays = new Date(year, month, 0).getDate();
  var sundays = 0;
  for (var d = 1; d <= totalDays; d++) { if (new Date(year, month - 1, d).getDay() === 0) sundays++; }

  // Use the cached Settings data — avoids a redundant getDataRange() call.
  ensureSettingsSeeded_();
  if (!_cachedSettings_) {
    _cachedSettings_ = getSheet(SHEETS.SETTINGS).getDataRange().getValues();
  }
  var data = _cachedSettings_;
  var events = [];
  var mStart = new Date(year, month - 1, 1), mEnd = new Date(year, month, 0);
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (safeStr(r[SET_COL.TYPE]) !== SETTINGS_TYPE.ACADEMIC_CALENDAR) continue;
    if (r[SET_COL.ACTIVE] === false || safeStr(r[SET_COL.ACTIVE]).toUpperCase() === 'FALSE') continue;
    var startStr = normCalDate_(r[SET_COL.KEY]);
    var endStr   = normCalDate_(r[SET_COL.VALUE3]);
    if (!startStr) continue;
    var s = parsePaymentDate_(startStr);
    var e = endStr ? parsePaymentDate_(endStr) : s;
    if (e.getTime() < mStart.getTime() || s.getTime() > mEnd.getTime()) continue;
    events.push({ date: startStr, endDate: endStr, category: safeStr(r[SET_COL.VALUE1]), title: safeStr(r[SET_COL.VALUE2]) });
  }
  events.sort(function(a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  return {
    success: true, month: month, year: year,
    totalDays: totalDays, sundays: sundays, workingDays: workingDays, events: events
  };
}

// ============================================================
//  STUDENT ATTENDANCE — daily Present/Absent marking per student
//  (like staff attendance). One row per student PER DAY. A monthly
//  summary and per-student history are computed on read, using the
//  Academic Calendar for the working-days denominator.
// ============================================================

// One-time migration: the earlier version stored a consolidated MONTHLY
// row (header began "Student ID","Name","Class","Section","Month",...).
// If that old layout is detected, archive the sheet so a fresh daily
// sheet is created — no monthly data is lost, it just moves aside.
function ensureDailyStudentAttendance_() {
  var ss = _getSS_();
  var sh = ss.getSheetByName(SHEETS.STUDENT_ATTENDANCE);
  if (!sh) return; // getSheet() will create a fresh daily sheet on demand
  if (sh.getLastRow() === 0) return; // empty — header written on first save
  var header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(function(h) { return safeStr(h); });
  var isOldMonthly = header[0] === "Student ID" && header.indexOf("Month") !== -1 && header.indexOf("Present Days") !== -1;
  if (isOldMonthly) {
    var archiveName = "Student Attendance (Monthly Archive)";
    if (ss.getSheetByName(archiveName)) archiveName = archiveName + " " + new Date().getTime();
    sh.setName(archiveName);
    // A fresh "Student Attendance" sheet is created lazily by getSheet().
  }
}

// Returns every active student in a class (+ optional section) with their
// Present/Absent status for one date — defaulting to "Present" for anyone
// not yet marked that day, so the teacher just flips the absentees.
function getStudentAttendanceForDate(token, className, section, dateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureDailyStudentAttendance_();
  var date = parsePaymentDate_(dateStr);
  var dayInfo = getDayHolidayInfo_(date);
  if (!className) return { success: true, date: dateStr, isHoliday: dayInfo.isHoliday, holidayLabel: dayInfo.label, dayName: dayInfo.dayName, data: [] };
  var profiles = getAllStudentProfiles_().filter(function(p) {
    if (!isActiveStudent_(p.status)) return false;
    if (className && safeStr(p.class) !== className) return false;
    if (section && safeStr(p.section) !== section) return false;
    return true;
  });

  var sh = getSheet(SHEETS.STUDENT_ATTENDANCE);
  var data = sh.getDataRange().getValues();
  var marked = {}; // studentId -> {status, remarks}
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var d = r[STATT_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (!isSameDay_(d, date)) continue;
    marked[safeStr(r[STATT_COL.STUDENT_ID])] = { status: safeStr(r[STATT_COL.STATUS]), remarks: safeStr(r[STATT_COL.REMARKS]) };
  }

  return {
    success: true, date: dateStr,
    isHoliday: dayInfo.isHoliday,
    holidayLabel: dayInfo.label,
    dayName: dayInfo.dayName,
    data: profiles.map(function(p) {
      var m = marked[safeStr(p.studentId)];
      return {
        studentId: safeStr(p.studentId), name: safeStr(p.name),
        class: safeStr(p.class), section: safeStr(p.section),
        status: m ? m.status : "Present", remarks: m ? m.remarks : "", alreadyMarked: !!m
      };
    })
  };
}

// Saves one day's attendance for a whole class in one call. `entries` is
// an array of { studentId, name, class, section, status, remarks }.
// Re-marking the same date + class replaces just those rows (so marking
// Class 5 never disturbs Class 6's attendance for the same day).
function saveStudentAttendanceDaily(token, dateStr, entries) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (!dateStr || !entries || !entries.length) return { error: "Nothing to save." };
  ensureDailyStudentAttendance_();

  var date = parsePaymentDate_(dateStr);
  var idsBeingSaved = {};
  entries.forEach(function(e) { idsBeingSaved[safeStr(e.studentId)] = true; });

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = getSheet(SHEETS.STUDENT_ATTENDANCE);
    if (sh.getLastRow() === 0) sh.appendRow(STATT_DAILY_HEADER);

    // Remove existing rows for this date for the students being re-marked.
    var data = sh.getDataRange().getValues();
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      var d = data[i][STATT_COL.DATE];
      if (Object.prototype.toString.call(d) !== "[object Date]") continue;
      if (isSameDay_(d, date) && idsBeingSaved[safeStr(data[i][STATT_COL.STUDENT_ID])]) {
        rowsToDelete.push(i + 1);
      }
    }
    rowsToDelete.sort(function(a, b) { return b - a; }).forEach(function(r) { sh.deleteRow(r); });

    var newRows = entries.map(function(e) {
      return [date, e.studentId, e.name, e.class, e.section || "", e.status || "Present", e.remarks || "", sess.name];
    });
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, STATT_NUM_COLS).setValues(newRows);
  } finally {
    lock.releaseLock();
  }

  return { success: true, message: "Attendance saved for " + entries.length + " student(s)." };
}

// Monthly summary per student for a class: days present / absent / marked,
// the working-days denominator (from the Academic Calendar), and %.
function getStudentAttendanceSummary(token, className, section, month, year) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureDailyStudentAttendance_();
  month = Number(month); year = Number(year);
  var profiles = getAllStudentProfiles_().filter(function(p) {
    if (!isActiveStudent_(p.status)) return false;
    if (className && safeStr(p.class) !== className) return false;
    if (section && safeStr(p.section) !== section) return false;
    return true;
  });

  var sh = getSheet(SHEETS.STUDENT_ATTENDANCE);
  var data = sh.getDataRange().getValues();
  var counts = {}; // studentId -> {present, absent}
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var d = r[STATT_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (d.getMonth() + 1 !== month || d.getFullYear() !== year) continue;
    var sid = safeStr(r[STATT_COL.STUDENT_ID]);
    if (!counts[sid]) counts[sid] = { present: 0, absent: 0 };
    var st = safeStr(r[STATT_COL.STATUS]);
    if (st === "Present") counts[sid].present++;
    else if (st === "Absent") counts[sid].absent++;
  }

  var workingDays = countWorkingDaysInMonthCal_(month, year);

  return {
    success: true, month: month, year: year, workingDays: workingDays,
    data: profiles.map(function(p) {
      var c = counts[safeStr(p.studentId)] || { present: 0, absent: 0 };
      var marked = c.present + c.absent;
      var pct = workingDays > 0 ? Math.round((c.present / workingDays) * 1000) / 10 : 0;
      return {
        studentId: safeStr(p.studentId), name: safeStr(p.name), class: safeStr(p.class), section: safeStr(p.section),
        present: c.present, absent: c.absent, markedDays: marked, workingDays: workingDays, percentage: pct
      };
    })
  };
}

// Per-student month-by-month history (computed from daily rows), used on
// the Student Profile screen and the report card. Same return shape as
// before so those screens keep working: history[] of
// {month, year, workingDays, presentDays, percentage, remarks}.
function getStudentAttendanceHistory(token, studentId) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureDailyStudentAttendance_();
  var sh = getSheet(SHEETS.STUDENT_ATTENDANCE);
  var data = sh.getDataRange().getValues();
  var byMonth = {}; // "year-month" -> {present, absent}
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (safeStr(r[STATT_COL.STUDENT_ID]) !== studentId) continue;
    var d = r[STATT_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    var key = d.getFullYear() + "-" + (d.getMonth() + 1);
    if (!byMonth[key]) byMonth[key] = { year: d.getFullYear(), month: d.getMonth() + 1, present: 0, absent: 0 };
    var st = safeStr(r[STATT_COL.STATUS]);
    if (st === "Present") byMonth[key].present++;
    else if (st === "Absent") byMonth[key].absent++;
  }

  var calMap = buildCalendarMap_();
  var history = Object.keys(byMonth).map(function(k) {
    var m = byMonth[k];
    var wd = countWorkingDaysInMonthCal_(m.month, m.year, calMap);
    var pct = wd > 0 ? Math.round((m.present / wd) * 1000) / 10 : 0;
    return { month: m.month, year: m.year, workingDays: wd, presentDays: m.present, percentage: pct, remarks: (m.absent ? (m.absent + " absent") : "") };
  });
  history.sort(function(a, b) { return (a.year - b.year) || (a.month - b.month); });

  var totalWorking = history.reduce(function(s, h) { return s + h.workingDays; }, 0);
  var totalPresent = history.reduce(function(s, h) { return s + h.presentDays; }, 0);
  var overallPct = totalWorking > 0 ? Math.round((totalPresent / totalWorking) * 1000) / 10 : 0;

  return { success: true, history: history, overall: { workingDays: totalWorking, presentDays: totalPresent, percentage: overallPct } };
}

// Active students marked Absent on a given date (default today), with the
// parent mobile — powers the WhatsApp "Absent" recipient filter.
function getAbsentStudentsForDate_(dateStr) {
  var date = parsePaymentDate_(dateStr);
  var sh = getSheet(SHEETS.STUDENT_ATTENDANCE);
  var data = sh.getDataRange().getValues();
  var absentIds = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var d = r[STATT_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (isSameDay_(d, date) && safeStr(r[STATT_COL.STATUS]) === "Absent") {
      absentIds[safeStr(r[STATT_COL.STUDENT_ID])] = true;
    }
  }
  return absentIds;
}

// ============================================================
//  EXAM PATTERN SETTINGS — which exams exist, their max/pass marks, and
//  which subjects each class is examined in. Fully editable from
//  Settings, same pattern as Classes/Fee Heads/Bank Accounts above.
// ============================================================

// Active exams only — used by the Marks Entry screen's exam dropdown.
function getExams(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.EXAM);
  return { success: true, data: rows.map(function(r) {
    return { rowIndex: r.rowIndex, name: r.key, maxMarks: Number(r.v1) || 0, passMarks: Number(r.v2) || 0, group: r.v3 };
  }) };
}

function addExam(token, obj) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!obj || !obj.name) return { error: "Exam name is required." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.EXAM && safeStr(data[i][SET_COL.KEY]) === obj.name) {
      return { error: "An exam named \"" + obj.name + "\" already exists." };
    }
  }
  sh.appendRow([SETTINGS_TYPE.EXAM, obj.name, safeNum(obj.maxMarks), safeNum(obj.passMarks), obj.group || "FA", "", "", true]);
  return { success: true, message: "Exam \"" + obj.name + "\" added." };
}

function updateExam(token, rowIndex, obj) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Exam not found." };
  sh.getRange(sheetRow, SET_COL.KEY + 1, 1, 4).setValues([[obj.name || "", safeNum(obj.maxMarks), safeNum(obj.passMarks), obj.group || "FA"]]);
  return { success: true, message: "Exam updated." };
}

function deactivateExam(token, rowIndex) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Exam not found." };
  sh.getRange(sheetRow, SET_COL.ACTIVE + 1).setValue(false);
  return { success: true, message: "Exam removed from active lists." };
}

// ── MASTER SUBJECT LIST ──────────────────────────────────────
// Notes request: "subject add option is also needed like computers,
// vocational, etc." — a school-wide list any Class→Subjects mapping is
// built from, instead of free-typed text.
function getSubjects(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.SUBJECT);
  return { success: true, data: rows.map(function(r) { return r.key; }) };
}

function addSubject(token, name) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!name) return { error: "Subject name is required." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.SUBJECT && safeStr(data[i][SET_COL.KEY]) === name) {
      return { error: "Subject \"" + name + "\" already exists." };
    }
  }
  sh.appendRow([SETTINGS_TYPE.SUBJECT, name, "", "", "", "", "", true]);
  return { success: true, message: "Subject \"" + name + "\" added." };
}

function deactivateSubject(token, rowIndex) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Subject not found." };
  sh.getRange(sheetRow, SET_COL.ACTIVE + 1).setValue(false);
  return { success: true, message: "Subject removed from active lists. Existing class mappings that already use it are not affected." };
}

// Full subject list with rowIndex, for the Settings screen's
// add/remove UI (mirrors getFeeHeadsFull()'s pattern).
function getSubjectsFull(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) !== SETTINGS_TYPE.SUBJECT) continue;
    var activeVal = data[i][SET_COL.ACTIVE];
    var isActive = !(activeVal === false || safeStr(activeVal).toUpperCase() === "FALSE");
    out.push({ rowIndex: i, name: safeStr(data[i][SET_COL.KEY]), active: isActive });
  }
  return { success: true, data: out };
}

// ── PER-SUBJECT MAX/PASS MARKS (Class + Exam + Subject) ──────
// Notes request: pass marks (and max marks) can vary per subject per
// exam per class — e.g. FA exams are normally 50/18, but Physical
// Science / Biological Science are graded out of half that (25/9) each
// within the same FA exam, then summed back into one "Science" result
// worth the full 50 for ranking/report-card purposes.
//
// A row in Settings (Type=ExamSubjectMark, Key="Class|Exam|Subject") is
// only created the first time a combination is actually edited away
// from its default — there's no need to pre-populate every class x
// exam x subject combination. getEffectiveSubjectMarks_() below is the
// single place that resolves "what are this subject's real max/pass
// marks", falling through: explicit override → combined-science half
// of the exam default → plain exam default.

function examSubjectMarkKey_(className, examName, subject) {
  return className + "|" + examName + "|" + subject;
}

// Returns { maxMarks, passMarks, isCombinedScience } for one specific
// class+exam+subject, resolving any explicit override and applying the
// combined-science halving rule. `exam` is the exam object (already
// looked up by the caller) so this never re-fetches Settings per subject.
function getEffectiveSubjectMarks_(className, exam, subject) {
  var isCombined = COMBINED_SCIENCE_SUBJECTS.indexOf(subject) !== -1;
  var baseMax = exam.maxMarks, basePass = exam.passMarks;
  if (isCombined) {
    // Halve the exam's normal max/pass for entry purposes — FA 50/18 -> 25/9 each.
    baseMax = Math.round((exam.maxMarks / 2) * 100) / 100;
    basePass = Math.round((exam.passMarks / 2) * 100) / 100;
  }

  var override = examSubjectMarkOverrides_[examSubjectMarkKey_(className, exam.name, subject)];
  return {
    maxMarks: override ? override.maxMarks : baseMax,
    passMarks: override ? override.passMarks : basePass,
    isCombinedScience: isCombined
  };
}

// Lazily-loaded cache of all ExamSubjectMark override rows, keyed by
// "Class|Exam|Subject" -> {maxMarks, passMarks, rowIndex}. Loaded once
// per request (Apps Script executions are short-lived, so this never
// goes stale within a single call) rather than re-scanning Settings for
// every subject of every student.
var examSubjectMarkOverrides_ = null;
function loadExamSubjectMarkOverrides_() {
  if (examSubjectMarkOverrides_) return;
  examSubjectMarkOverrides_ = {};
  var rows = getSettingsRows_(SETTINGS_TYPE.EXAM_SUBJECT_MARK);
  rows.forEach(function(r) {
    examSubjectMarkOverrides_[r.key] = { maxMarks: Number(r.v1) || 0, passMarks: Number(r.v2) || 0, rowIndex: r.rowIndex };
  });
}

// Returns the full per-subject max/pass grid for a class+exam — every
// subject configured for that class, with its effective (resolved)
// max/pass marks — used by both the Marks Entry screen (to size each
// input correctly) and the Settings → Subject Marks editor.
function getSubjectMarksForClassExam(token, className, examName) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  loadExamSubjectMarkOverrides_();

  var examsResult = getExams(token);
  var exam = (examsResult.data || []).filter(function(e) { return e.name === examName; })[0];
  if (!exam) return { error: "Exam \"" + examName + "\" not found or inactive." };

  var subjectsResult = getSubjectsForClass(token, className);
  var subjects = subjectsResult.data || [];

  var data = subjects.map(function(subj) {
    var eff = getEffectiveSubjectMarks_(className, exam, subj);
    var key = examSubjectMarkKey_(className, examName, subj);
    var override = examSubjectMarkOverrides_[key];
    return {
      subject: subj, maxMarks: eff.maxMarks, passMarks: eff.passMarks,
      isCombinedScience: eff.isCombinedScience, isOverridden: !!override,
      rowIndex: override ? override.rowIndex : null
    };
  });

  return { success: true, exam: exam, data: data };
}

// Saves (or updates) an explicit max/pass override for one class+exam+
// subject combination. Setting both values back to exactly the resolved
// default removes the override row instead of keeping a redundant one.
function setExamSubjectMarks(token, className, examName, subject, maxMarks, passMarks) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!className || !examName || !subject) return { error: "Class, exam, and subject are required." };

  maxMarks = safeNum(maxMarks);
  passMarks = safeNum(passMarks);
  if (maxMarks <= 0) return { error: "Enter valid max marks." };
  if (passMarks < 0 || passMarks > maxMarks) return { error: "Pass marks must be between 0 and max marks." };

  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var key = examSubjectMarkKey_(className, examName, subject);
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.EXAM_SUBJECT_MARK && safeStr(data[i][SET_COL.KEY]) === key) {
      sh.getRange(i + 1, SET_COL.VALUE1 + 1, 1, 2).setValues([[maxMarks, passMarks]]);
      examSubjectMarkOverrides_ = null; // invalidate cache
      return { success: true, message: "Marks updated for " + subject + " in " + examName + " (Class " + className + ")." };
    }
  }
  sh.appendRow([SETTINGS_TYPE.EXAM_SUBJECT_MARK, key, maxMarks, passMarks, "", "", "", true]);
  examSubjectMarkOverrides_ = null; // invalidate cache
  return { success: true, message: "Custom marks set for " + subject + " in " + examName + " (Class " + className + ")." };
}

// Removes an override, reverting that subject back to the exam's plain
// default (or combined-science half-default).
function clearExamSubjectMarksOverride(token, rowIndex) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  var sh = getSheet(SHEETS.SETTINGS);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > sh.getLastRow()) return { error: "Override not found." };
  sh.deleteRow(sheetRow);
  examSubjectMarkOverrides_ = null; // invalidate cache
  return { success: true, message: "Reverted to default marks." };
}

// Subject list for one class — used by the Marks Entry screen once a
// class is picked, to know which subjects to show rows for.
function getSubjectsForClass(token, className) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.CLASS_SUBJECTS);
  var row = rows.filter(function(r) { return r.key === className; })[0];
  var subjects = row && row.v1 ? row.v1.split(",").map(function(s) { return s.trim(); }).filter(Boolean) : [];
  return { success: true, data: subjects };
}

// Full class-subject mapping for the Settings screen itself.
function getClassSubjectsFull(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var rows = getSettingsRows_(SETTINGS_TYPE.CLASS_SUBJECTS);
  return { success: true, data: rows.map(function(r) {
    return { rowIndex: r.rowIndex, className: r.key, subjects: r.v1 ? r.v1.split(",").map(function(s) { return s.trim(); }).filter(Boolean) : [] };
  }) };
}

function setClassSubjects(token, className, subjects) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin") return { error: PERM_MSG.ADMIN_ONLY };
  if (!className) return { error: "Class is required." };

  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var subjectStr = Array.isArray(subjects) ? subjects.join(",") : safeStr(subjects);
  for (var i = 1; i < data.length; i++) {
    if (safeStr(data[i][SET_COL.TYPE]) === SETTINGS_TYPE.CLASS_SUBJECTS && safeStr(data[i][SET_COL.KEY]) === className) {
      sh.getRange(i + 1, SET_COL.VALUE1 + 1).setValue(subjectStr);
      return { success: true, message: "Subjects updated for Class " + className + "." };
    }
  }
  sh.appendRow([SETTINGS_TYPE.CLASS_SUBJECTS, className, subjectStr, "", "", "", "", true]);
  return { success: true, message: "Subjects set for Class " + className + "." };
}

// ============================================================
//  MARKS ENTRY & AUTO-RANKING
//  One row per student per subject per exam. Rank is computed on
//  read from each student's total marks across that class+exam's
//  subjects, never stored — so changing a mark or the exam's pass-mark
//  setting always re-evaluates rank/pass-fail correctly with no stale
//  data left behind.
// ============================================================

// Returns the current academic year string the same way getAcademicYear_
// does, for tagging new marks rows.
function getCurrentAcademicYear_() {
  return getAcademicYear_(new Date());
}

// Loads the mark sheet for one class + exam: every active student in
// that class, with their existing marks (if any) for every subject
// configured for that class. This is what the Marks Entry screen
// renders as an editable grid.
function getMarksSheet(token, className, section, examName) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  if (!className || !examName) return { error: "Class and exam are required." };

  var subjectsResult = getSubjectsForClass(token, className);
  var subjects = subjectsResult.data || [];
  if (!subjects.length) return { error: "No subjects configured for Class " + className + " yet — set them up in Settings → Exam Pattern first." };

  var examsResult = getExams(token);
  var exam = (examsResult.data || []).filter(function(e) { return e.name === examName; })[0];
  if (!exam) return { error: "Exam \"" + examName + "\" not found or inactive." };

  // Per-subject max/pass marks (Notes request: pass/max marks can vary
  // per subject per exam per class — e.g. Physical/Biological Science
  // graded out of half the normal max within the same exam).
  loadExamSubjectMarkOverrides_();
  var subjectMeta = subjects.map(function(subj) {
    var eff = getEffectiveSubjectMarks_(className, exam, subj);
    return { subject: subj, maxMarks: eff.maxMarks, passMarks: eff.passMarks, isCombinedScience: eff.isCombinedScience };
  });

  var profiles = getAllStudentProfiles_().filter(function(p) {
    if (p.status !== "Active") return false;
    if (p.class !== className) return false;
    if (section && p.section !== section) return false;
    return true;
  });

  var sh = getSheet(SHEETS.MARKS);
  var data = sh.getDataRange().getValues();
  var existing = {}; // studentId -> { subject: marksObtained }
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (safeStr(r[MARKS_COL.EXAM]) !== examName || safeStr(r[MARKS_COL.CLASS]) !== className) continue;
    var sid = safeStr(r[MARKS_COL.STUDENT_ID]);
    if (!existing[sid]) existing[sid] = {};
    existing[sid][safeStr(r[MARKS_COL.SUBJECT])] = safeNum(r[MARKS_COL.MARKS_OBTAINED]);
  }

  var students = profiles.map(function(p) {
    var marks = {};
    subjects.forEach(function(subj) {
      marks[subj] = (existing[p.studentId] && existing[p.studentId][subj] !== undefined) ? existing[p.studentId][subj] : "";
    });
    return { studentId: p.studentId, name: p.name, class: p.class, section: p.section, marks: marks };
  });

  return { success: true, subjects: subjects, subjectMeta: subjectMeta, exam: exam, data: students };
}

// Saves marks for a whole class+exam in one call. `entries` is an array
// of { studentId, name, class, section, marks: {subject: value} }.
// Replaces any existing rows for this exact class+exam+student rather
// than appending duplicates, same re-save pattern as attendance.
function saveMarksSheet(token, className, examName, subjects, entries) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (!className || !examName || !entries || !entries.length) return { error: "Nothing to save." };

  var academicYear = getCurrentAcademicYear_();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = getSheet(SHEETS.MARKS);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["Student ID", "Name", "Class", "Section", "Exam", "Subject", "Marks Obtained", "Max Marks", "Pass Marks", "Academic Year", "Entered By"]);
    }

    // Remove existing rows for this class+exam (any student, any subject
    // already configured) before re-inserting fresh — simplest way to
    // guarantee no duplicate/stale subject rows linger after a re-save.
    var data = sh.getDataRange().getValues();
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      if (safeStr(data[i][MARKS_COL.CLASS]) === className && safeStr(data[i][MARKS_COL.EXAM]) === examName) {
        rowsToDelete.push(i + 1);
      }
    }
    rowsToDelete.sort(function(a, b) { return b - a; }).forEach(function(r) { sh.deleteRow(r); });

    var examsResult = getExams(token);
    var exam = (examsResult.data || []).filter(function(e) { return e.name === examName; })[0];
    loadExamSubjectMarkOverrides_();

    var newRows = [];
    entries.forEach(function(e) {
      subjects.forEach(function(subj) {
        var val = e.marks && e.marks[subj] !== undefined && e.marks[subj] !== "" ? safeNum(e.marks[subj]) : null;
        if (val === null) return; // skip subjects left blank for this student
        var eff = exam ? getEffectiveSubjectMarks_(className, exam, subj) : { maxMarks: 0 };
        newRows.push([e.studentId, e.name, e.class, e.section || "", examName, subj, val, eff.maxMarks, eff.passMarks, academicYear, sess.name]);
      });
    });
    if (newRows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, MARKS_NUM_COLS).setValues(newRows);
    }
  } finally {
    lock.releaseLock();
  }

  return { success: true, message: "Marks saved for " + entries.length + " student(s)." };
}

// Computes total/percentage/pass-fail/RANK for every student in a
// class+exam, ranked within that class only (Notes confirmed: rank
// within own class, not school-wide). Ties share the same rank (e.g.
// two students tied for 2nd both show rank 2, next student is rank 4).
function getMarksWithRank(token, className, examName) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var sheet = getMarksSheet(token, className, "", examName);
  if (sheet.error) return sheet;

  var exam = sheet.exam;

  // Build the grading-component list: most subjects are 1:1, but the
  // two combined-science subjects collapse into a single "Science"
  // component whose marks are the SUM of both and whose max/pass marks
  // are the exam's plain (non-halved) defaults — Notes request: "while
  // writing exams biology & physical science subject marks cumulates
  // one subject... FA exam 25 marks for Physical Science, 25 marks for
  // Biology Science" i.e. entered separately out of half each, graded
  // together out of the full amount.
  var presentScience = sheet.subjectMeta.filter(function(m) { return m.isCombinedScience; });
  var components = sheet.subjectMeta.filter(function(m) { return !m.isCombinedScience; }).map(function(m) {
    return { label: m.subject, subjects: [m.subject], maxMarks: m.maxMarks, passMarks: m.passMarks };
  });
  if (presentScience.length) {
    components.push({
      label: "Science", subjects: presentScience.map(function(m) { return m.subject; }),
      maxMarks: exam.maxMarks, passMarks: exam.passMarks
    });
  }

  var results = sheet.data.map(function(s) {
    var subjectsAttempted = 0, total = 0, subjectFails = [];
    components.forEach(function(comp) {
      var anyEntered = comp.subjects.some(function(subj) { return s.marks[subj] !== "" && s.marks[subj] !== undefined && s.marks[subj] !== null; });
      if (!anyEntered) return;
      var compTotal = comp.subjects.reduce(function(sum, subj) {
        var m = s.marks[subj];
        return sum + (m === "" || m === undefined || m === null ? 0 : Number(m));
      }, 0);
      subjectsAttempted++;
      total += compTotal;
      if (compTotal < comp.passMarks) subjectFails.push(comp.label);
    });
    var maxPossible = components.reduce(function(sum, comp) {
      var anyEntered = comp.subjects.some(function(subj) { return s.marks[subj] !== "" && s.marks[subj] !== undefined && s.marks[subj] !== null; });
      return sum + (anyEntered ? comp.maxMarks : 0);
    }, 0);
    var percentage = maxPossible > 0 ? Math.round((total / maxPossible) * 1000) / 10 : 0;
    var passed = subjectsAttempted > 0 && subjectFails.length === 0;
    return {
      studentId: s.studentId, name: s.name, marks: s.marks,
      total: total, maxPossible: maxPossible, percentage: percentage,
      passed: passed, failedSubjects: subjectFails, subjectsAttempted: subjectsAttempted
    };
  });

  // Rank by total descending; students with zero subjects attempted
  // (absent for the whole exam) are excluded from ranking entirely.
  var ranked = results.filter(function(r) { return r.subjectsAttempted > 0; })
    .sort(function(a, b) { return b.total - a.total; });
  var rankMap = {};
  var currentRank = 0, lastTotal = null, seen = 0;
  ranked.forEach(function(r) {
    seen++;
    if (r.total !== lastTotal) { currentRank = seen; lastTotal = r.total; }
    rankMap[r.studentId] = currentRank;
  });

  results.forEach(function(r) { r.rank = rankMap[r.studentId] || null; });
  results.sort(function(a, b) { return (a.rank || 9999) - (b.rank || 9999); });

  return { success: true, subjects: sheet.subjects, components: components, exam: exam, data: results };
}

// Every exam result for one student across the current academic year —
// used on the Student Profile screen and the consolidated report card's
// marks progress section.
function getStudentMarksHistory(token, studentId) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var sh = getSheet(SHEETS.MARKS);
  var data = sh.getDataRange().getValues();
  // examName -> { subjectMarks: {subject: marksObtained}, subjectMax: {subject: maxMarks},
  //               subjectPass: {subject: passMarks}, class }
  var byExam = {};

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (safeStr(r[MARKS_COL.STUDENT_ID]) !== studentId) continue;
    var examName = safeStr(r[MARKS_COL.EXAM]);
    if (!byExam[examName]) byExam[examName] = { subjectMarks: {}, subjectMax: {}, subjectPass: {}, class: safeStr(r[MARKS_COL.CLASS]) };
    var subj = safeStr(r[MARKS_COL.SUBJECT]);
    byExam[examName].subjectMarks[subj] = safeNum(r[MARKS_COL.MARKS_OBTAINED]);
    byExam[examName].subjectMax[subj] = safeNum(r[MARKS_COL.MAX_MARKS]); // real per-subject max, stored at save time
    byExam[examName].subjectPass[subj] = safeNum(r[MARKS_COL.PASS_MARKS]); // real per-subject pass mark, stored at save time
  }

  var examsResult = getExams(token);
  var examMeta = {};
  (examsResult.data || []).forEach(function(e) { examMeta[e.name] = e; });

  var history = Object.keys(byExam).map(function(examName) {
    var ex = byExam[examName];
    var exam = examMeta[examName];
    var subjectNames = Object.keys(ex.subjectMarks);

    // Combine Physical Science + Biological Science into one "Science"
    // result (summed marks, full exam max/pass) — same rule as
    // getMarksWithRank(), applied here per-student instead of per-class.
    var presentScience = subjectNames.filter(function(s) { return COMBINED_SCIENCE_SUBJECTS.indexOf(s) !== -1; });
    var plainSubjects = subjectNames.filter(function(s) { return COMBINED_SCIENCE_SUBJECTS.indexOf(s) === -1; });

    var total = plainSubjects.reduce(function(s, subj) { return s + ex.subjectMarks[subj]; }, 0);
    var maxPossible = plainSubjects.reduce(function(s, subj) { return s + ex.subjectMax[subj]; }, 0);
    var failedSubjects = plainSubjects.filter(function(subj) { return ex.subjectMarks[subj] < ex.subjectPass[subj]; });

    if (presentScience.length) {
      var scienceTotal = presentScience.reduce(function(s, subj) { return s + ex.subjectMarks[subj]; }, 0);
      total += scienceTotal;
      var scienceMax = exam ? exam.maxMarks : presentScience.reduce(function(s, subj) { return s + ex.subjectMax[subj]; }, 0);
      maxPossible += scienceMax;
      var sciencePass = exam ? exam.passMarks : presentScience.reduce(function(s, subj) { return s + ex.subjectPass[subj]; }, 0);
      if (scienceTotal < sciencePass) failedSubjects.push("Science");
    }

    var percentage = maxPossible > 0 ? Math.round((total / maxPossible) * 1000) / 10 : 0;
    return {
      exam: examName, group: exam ? exam.group : "",
      class: ex.class, subjects: ex.subjectMarks, total: total, maxPossible: maxPossible,
      percentage: percentage, passed: failedSubjects.length === 0, failedSubjects: failedSubjects
    };
  });

  // Sort by the school's exam cycle order (FA1, FA2, SA1, FA3, FA4, SA2)
  // where possible, falling back to alphabetical for any custom exam names.
  var order = SEED_EXAMS.map(function(e) { return e.name; });
  history.sort(function(a, b) {
    var ia = order.indexOf(a.exam), ib = order.indexOf(b.exam);
    if (ia === -1 && ib === -1) return a.exam.localeCompare(b.exam);
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });

  return { success: true, history: history };
}

// ============================================================
//  CONSOLIDATED STUDENT REPORT CARD
//  Notes request: "student consolidated report card with student
//  profile complete, financial report with history, attendance report,
//  marks progress report in a professional reporting format with print
//  option". Reuses every module already built above rather than
//  re-querying the sheets directly, so this always reflects the same
//  numbers seen elsewhere in the app.
// ============================================================
function getConsolidatedReportCard(token, studentId) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var profileDetail = getStudentProfileDetail(token, studentId);
  if (profileDetail.error) return profileDetail;

  var attendance = getStudentAttendanceHistory(token, studentId);
  var marks = getStudentMarksHistory(token, studentId);
  var school = getSchoolInfo(token);

  return {
    success: true,
    school: school,
    profile: profileDetail.profile,
    finance: {
      billed: profileDetail.fee.billed,
      paid: profileDetail.fee.paid,
      due: profileDetail.fee.due,
      heads: profileDetail.fee.heads,
      transactions: profileDetail.transactions
    },
    attendance: attendance.error ? { history: [], overall: { workingDays: 0, presentDays: 0, percentage: 0 } } : attendance,
    marks: marks.error ? { history: [] } : marks,
    generatedOn: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy")
  };
}



// Builds a single-exam report-card PDF (profile + subject marks + result)
// and a pre-filled WhatsApp message — the marks counterpart of the
// outstanding fee statement attachment. Used by the WhatsApp module.
function getReportCardPdf(token, studentId, examName) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  if (!examName) return { error: "Please choose an exam first." };

  var detail = getStudentProfileDetail(token, studentId);
  if (detail.error) return detail;
  var p = detail.profile;

  var marksRes = getStudentMarksHistory(token, studentId);
  var examRec = null;
  (marksRes.history || []).forEach(function(h) { if (h.exam === examName) examRec = h; });
  if (!examRec || !Object.keys(examRec.subjects || {}).length) {
    return { error: "No marks recorded for " + p.name + " in " + examName + " yet." };
  }

  var attRes = getStudentAttendanceHistory(token, studentId);
  var attOverall = (attRes && attRes.overall) ? attRes.overall : { workingDays: 0, presentDays: 0, percentage: 0 };

  var school = getSchoolInfo(token);
  var asOfDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy");
  var html = buildExamReportCardHtml_(school, p, examName, examRec, attOverall, asOfDateStr);
  var fileName = "Report_Card_" + p.studentId + "_" + examName.replace(/[^A-Za-z0-9]/g, "");

  var pdfBytes;
  try { pdfBytes = htmlToPdfBlob_(html, fileName); }
  catch (e) { return { error: "Could not generate the PDF: " + e.message }; }

  var phone = p.fatherMobile ? p.fatherMobile.replace(/\D/g, "") : (p.motherMobile ? p.motherMobile.replace(/\D/g, "") : "");
  if (phone.length === 10) phone = "91" + phone;

  var msg = [
    "Dear Parent,", "",
    "Please find attached the *" + examName + "* report card for *" + p.name + "* (" + p.studentId + ", Class " + p.class + ").", "",
    "Marks: *" + examRec.total + " / " + examRec.maxPossible + "* (" + examRec.percentage + "%) \u2014 " + (examRec.passed ? "Passed" : "Needs Improvement"), "",
    "For any queries, please contact the school office.", "",
    "_" + school.name + "_"
  ].join("\n");

  return {
    success: true,
    fileName: fileName + ".pdf",
    pdfBase64: Utilities.base64Encode(pdfBytes),
    waURL: "https://wa.me/" + phone + "?text=" + encodeURIComponent(msg),
    phone: phone, studentName: p.name
  };
}

function buildExamReportCardHtml_(school, p, examName, examRec, attOverall, asOfDateStr) {
  var subjectRows = Object.keys(examRec.subjects).map(function(subj) {
    return "<tr><td style='padding:5px 8px;border-bottom:1px solid #e2e2e2;'>" + escHtml_(subj) + "</td>" +
      "<td style='padding:5px 8px;border-bottom:1px solid #e2e2e2;text-align:right;'>" + safeNum(examRec.subjects[subj]) + "</td></tr>";
  }).join("");
  var resultColor = examRec.passed ? "#1A7A4A" : "#C0392B";
  var resultText = examRec.passed ? "PASS" : "NEEDS IMPROVEMENT";
  return "" +
    "<div style='font-family:Arial,sans-serif;color:#1A1A1A;'>" +
      "<div style='text-align:center;border-bottom:3px solid #0D2137;padding-bottom:10px;margin-bottom:14px;'>" +
        "<div style='font-size:20px;font-weight:bold;color:#0D2137;'>" + escHtml_(school.name) + "</div>" +
        (school.address ? "<div style='font-size:11px;color:#555;margin-top:3px;'>" + escHtml_(school.address) + "</div>" : "") +
        (school.phone || school.email ? "<div style='font-size:11px;color:#555;margin-top:2px;'>" + [school.phone, school.email].filter(Boolean).map(escHtml_).join(" &nbsp;|&nbsp; ") + "</div>" : "") +
      "</div>" +
      "<div style='text-align:center;font-size:15px;font-weight:bold;color:#0A7E8C;margin-bottom:4px;'>REPORT CARD \u2014 " + escHtml_(examName) + "</div>" +
      "<div style='text-align:center;font-size:11px;color:#666;margin-bottom:16px;'>Generated on " + escHtml_(asOfDateStr) + "</div>" +
      "<table style='width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px;'>" +
        "<tr><td style='width:50%;padding:3px 0;color:#666;'>Student Name</td><td style='font-weight:bold;'>" + escHtml_(p.name) + " (" + escHtml_(p.studentId) + ")</td></tr>" +
        "<tr><td style='padding:3px 0;color:#666;'>Class / Section</td><td style='font-weight:bold;'>" + escHtml_(p.class) + (p.section ? " - " + escHtml_(p.section) : "") + "</td></tr>" +
        "<tr><td style='padding:3px 0;color:#666;'>Father's Name</td><td style='font-weight:bold;'>" + escHtml_(p.fatherName || "\u2014") + "</td></tr>" +
        "<tr><td style='padding:3px 0;color:#666;'>Attendance (overall)</td><td style='font-weight:bold;'>" + safeNum(attOverall.presentDays) + " / " + safeNum(attOverall.workingDays) + " (" + safeNum(attOverall.percentage) + "%)</td></tr>" +
      "</table>" +
      "<div style='font-size:13px;font-weight:bold;color:#0D2137;margin-bottom:6px;'>Subject-wise Marks</div>" +
      "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:14px;'>" +
        "<tr style='background:#0D2137;color:#fff;'><th style='padding:6px 8px;text-align:left;'>Subject</th><th style='padding:6px 8px;text-align:right;'>Marks</th></tr>" +
        subjectRows +
        "<tr style='background:#F3F7FA;font-weight:bold;'><td style='padding:6px 8px;'>Total</td><td style='padding:6px 8px;text-align:right;'>" + safeNum(examRec.total) + " / " + safeNum(examRec.maxPossible) + "</td></tr>" +
      "</table>" +
      "<table style='width:100%;border-collapse:collapse;margin-bottom:16px;'>" +
        "<tr>" +
          "<td style='width:49%;background:#0A7E8C;color:#fff;padding:10px;text-align:center;border-radius:4px;'><div style='font-size:10px;'>PERCENTAGE</div><div style='font-size:18px;font-weight:bold;'>" + safeNum(examRec.percentage) + "%</div></td>" +
          "<td style='width:2%;'>&nbsp;</td>" +
          "<td style='width:49%;background:" + resultColor + ";color:#fff;padding:10px;text-align:center;border-radius:4px;'><div style='font-size:10px;'>RESULT</div><div style='font-size:18px;font-weight:bold;'>" + resultText + "</div></td>" +
        "</tr>" +
      "</table>" +
      (examRec.failedSubjects && examRec.failedSubjects.length ? "<div style='font-size:11px;color:#C0392B;margin-bottom:10px;'>Needs improvement in: " + escHtml_(examRec.failedSubjects.join(", ")) + "</div>" : "") +
      "<div style='margin-top:24px;font-size:10px;color:#888;text-align:center;border-top:1px solid #ddd;padding-top:8px;'>" +
        "This is a system-generated report card. For any discrepancy, please contact the school office.<br>Generated via School DBMS &middot; Manha E-Solutions" +
      "</div>" +
    "</div>";
}

// ── DASHBOARD KPIs ────────────────────────────────────────────
// v3: now driven entirely by getFeeMap_() reading Tution_Fee only — fixes
// the wrong-column / stale Bus_Fee-sheet bug described in Notes item #1.
function getDashboardKPIs(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var profiles = getAllStudentProfiles_();
  var feeMap   = getFeeMap_();

  var totalBilled = 0, totalCollected = 0, totalDue = 0, totalAdvance = 0;
  var classSummary  = {};
  var transportCount = {};
  var headTotals = {}; // fee head -> {billed, paid}

  profiles.forEach(function(p) {
    var f = feeMap[p.studentId] || { billed: 0, paid: 0, due: 0, heads: {} };
    totalBilled    += f.billed;
    totalCollected += f.paid;
    // Positive-only, matching getOutstandingList_() / the Balance Sheet's
    // Fee Receivable exactly — a student who has paid MORE than billed
    // (advance/overpayment) has a negative due, which is a credit
    // balance, not a discount on what other students owe. It must never
    // net against genuine dues; it's tracked separately in totalAdvance.
    if (f.due > 0) totalDue += f.due; else totalAdvance += -f.due;

    var cls = p.class || "Unspecified";
    if (!classSummary[cls]) classSummary[cls] = { billed: 0, collected: 0, due: 0, count: 0 };
    classSummary[cls].billed    += f.billed;
    classSummary[cls].collected += f.paid;
    classSummary[cls].due       += (f.due > 0 ? f.due : 0); // positive-only, same reasoning as totalDue above
    classSummary[cls].count++;

    var tr = p.transport || "Other";
    transportCount[tr] = (transportCount[tr] || 0) + 1;

    Object.keys(f.heads || {}).forEach(function(h) {
      if (!headTotals[h]) headTotals[h] = { billed: 0, paid: 0 };
      headTotals[h].billed += f.heads[h].billed;
      headTotals[h].paid   += f.heads[h].paid;
    });
  });

  return {
    success: true,
    kpis: {
      totalStudents:  profiles.length,
      totalBilled:    totalBilled,
      totalCollected: totalCollected,
      totalDue:       totalDue,
      totalAdvance:   totalAdvance,
      collectionRate: totalBilled > 0 ? Math.round(totalCollected / totalBilled * 100) : 0,
      dueCount:       profiles.filter(function(p) {
                          var f = feeMap[p.studentId]; return f && f.due > 0;
                       }).length
    },
    transport:    transportCount,
    classSummary: classSummary,
    headTotals:   headTotals
  };
}

// ── SNAPSHOT INFOGRAPHIC — one-call bundle of headline numbers ──
// Powers the "Fee Flow Chain" infographic (Students → Billed → Collected
// → Outstanding → Net Worth) plus its four supporting stat cards. Reuses
// the same calculations already shown elsewhere (Dashboard, Balance
// Sheet) rather than recomputing anything independently, so the
// infographic can never drift out of sync with those numbers.
function getSnapshotInfographicData(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var kpis = getDashboardKPIs(token);
  if (kpis.error) return kpis;

  var bs = getBalanceSheetData_(token, null); // "as of today"
  if (bs.error) return bs;

  var classes = getClassList(token);
  var teachers = getTeacherList(token);
  var activeTeachers = (teachers.data || []).filter(function (t) { return t.status !== "Removed" && t.status !== "Inactive"; });

  // School-wide attendance %, trailing 30 days, across every student —
  // read directly since the built-in attendance reports are per-class or
  // per-student, not a single whole-school figure.
  var attData = getSheet(SHEETS.STUDENT_ATTENDANCE).getDataRange().getValues();
  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 30);
  var present = 0, total = 0;
  for (var i = 1; i < attData.length; i++) {
    var d = attData[i][STATT_COL.DATE];
    if (!(d instanceof Date) || d < cutoff) continue;
    total++;
    var status = safeStr(attData[i][STATT_COL.STATUS]).toLowerCase();
    if (status === "present" || status === "p") present++;
  }
  var attendancePct = total > 0 ? Math.round((present / total) * 1000) / 10 : 0;

  // Expenses for the current financial year to date — same window the
  // Balance Sheet's surplus figure uses, so the two stay consistent.
  var today = new Date();
  var fyStart = financialYearStart_(today);
  var incExp = getIncomeExpenditureSummary(token, _ymd_(fyStart), _ymd_(today));
  var expensesFY = incExp.error ? 0 : incExp.expenditure.total;

  return {
    success: true,
    schoolName: bs.school.name,
    asOn: bs.asOn,
    students: kpis.kpis.totalStudents,
    billed: kpis.kpis.totalBilled,
    collected: kpis.kpis.totalCollected,
    outstanding: kpis.kpis.totalDue,
    collectionRate: kpis.kpis.collectionRate,
    netWorth: bs.liabilities.capitalFund,
    classesCount: (classes.data || []).length,
    staffCount: activeTeachers.length,
    attendancePct: attendancePct,
    expensesFY: expensesFY
  };
}

// ── MONTHLY TREND (derived from Transactions dates) ─────────
function getMonthlyCollectionTrend(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var txSh   = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();
  var monthLabels = ["Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May"];
  var totals = new Array(12).fill(0);

  for (var i = 1; i < txData.length; i++) {
    var row = txData[i];
    if (isVoided_(row[TX_COL.VOID])) continue;
    var d = row[TX_COL.DATE];
    var amt = safeNum(row[TX_COL.AMOUNT]);
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    var m = d.getMonth(); // 0=Jan..11=Dec
    // Academic year starts June(5): map Jun->0 ... May->11
    var idx = (m - 5 + 12) % 12;
    totals[idx] += amt;
  }
  return { success: true, months: monthLabels, totals: totals };
}

// ============================================================
//  EXPENSES — Cash Deposits / Withdrawals / Stationery / Repairs &
//  Maintenance / Misc. / Salaries / Electricity / Others
//  New feature: tracks every outgoing entry against the school's cash
//  and bank position, separate from student fee Transactions, so
//  Financial Analytics can show true Income & Expenditure.
// ============================================================

// NOTE: getExpenseCategories() now lives in the SETTINGS module above,
// reading from the Settings sheet instead of a hardcoded list.

// Voucher numbering: PREFIX-EXP/ACADEMIC-YEAR/SEQ, e.g. AHS-EXP/26-27/0001.
// Mirrors getNextReceiptNo_() but keeps its own separate sequence so
// expense vouchers and fee receipts never share/collide on numbers.
function getNextVoucherNo_(forDate) {
  var ay     = getAcademicYear_(forDate);
  var prefix = getReceiptPrefix_() + "-EXP/" + ay + "/";
  var expSh  = getSheet(SHEETS.EXPENSES);
  var data   = expSh.getDataRange().getValues();
  var maxSeq = 0;
  for (var i = 1; i < data.length; i++) {
    var v = safeStr(data[i][EXP_COL.VOUCHER]);
    if (v.indexOf(prefix) === 0) {
      var seq = parseInt(v.substring(prefix.length), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  return prefix + ("0000" + (maxSeq + 1)).slice(-4);
}

function getNextVoucherNo(token, dateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  return { success: true, voucherNo: getNextVoucherNo_(parsePaymentDate_(dateStr)) };
}

// Records a new expense entry. amount > 0 always; the category itself
// (Cash Deposit vs Cash Withdrawal vs a real expense) determines how it's
// treated in cash-flow / income-expenditure calculations, not the sign.
function recordExpense(token, category, description, amount, mode, dateStr, bankAccount, toBankAccount) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };

  amount = safeNum(amount);
  if (amount <= 0) return { error: "Enter a valid amount." };
  if (!category) return { error: "Please select a category." };

  // Notes request: Bank to Bank Transfer needs a From account AND a To
  // account, and they must actually be different — transferring an
  // account into itself isn't a real movement and would silently break
  // the reconciliation (crediting and debiting the same account by the
  // same amount looks like a no-op, but is really a sign of a mistake).
  if (category === "Bank to Bank Transfer") {
    if (!bankAccount) return { error: "Select the FROM bank account for this transfer." };
    if (!toBankAccount) return { error: "Select the TO bank account for this transfer." };
    if (bankAccount === toBankAccount) return { error: "From and To accounts must be different." };
  }
  // Cash Deposit / Withdrawal must name the bank side, or the money would
  // leave one side without arriving on the other (breaking the reconciliation).
  if (category === "Cash Deposit" && !bankAccount) return { error: "Select the bank account the cash was deposited INTO." };
  if (category === "Cash Withdrawal" && !bankAccount) return { error: "Select the bank account the cash was withdrawn FROM." };

  var expDate = parsePaymentDate_(dateStr);
  if (expDate.getTime() > new Date().getTime() + 24 * 60 * 60 * 1000) {
    return { error: "Expense date cannot be in the future." };
  }

  var lock = LockService.getScriptLock();
  var voucherNo;
  try {
    lock.waitLock(10000);
    var expSh = getSheet(SHEETS.EXPENSES);
    if (expSh.getLastRow() === 0) {
      expSh.appendRow(["Date", "Voucher No", "Category", "Description", "Mode", "Amount", "By", "Voided", "Bank Account", "To Bank Account"]);
    }
    voucherNo = getNextVoucherNo_(expDate);
    expSh.appendRow([expDate, voucherNo, category, description || "", mode || "Cash", amount, sess.name, false, bankAccount || "", toBankAccount || ""]);
  } finally {
    lock.releaseLock();
  }

  return { success: true, voucherNo: voucherNo, message: "Voucher " + voucherNo + " · ₹" + amount + " recorded under " + category + "." };
}

// Same audit-trail pattern as cancelPayment() — marks the row voided
// rather than deleting it, so the history always shows what happened and
// who reversed it.
function cancelExpense(token, rowIndex, reason) {
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  if (rowIndex === undefined || rowIndex === null) return { error: "Expense not specified." };

  var expSh = getSheet(SHEETS.EXPENSES);
  var sheetRow = Number(rowIndex) + 1;
  if (sheetRow < 2 || sheetRow > expSh.getLastRow()) return { error: "Expense not found." };

  var existingVoucher = expSh.getRange(sheetRow, EXP_COL.VOUCHER + 1).getValue();
  if (!existingVoucher) return { error: "Expense not found." };

  expSh.getRange(sheetRow, EXP_COL.VOID + 1).setValue(true);
  var byCell = safeStr(expSh.getRange(sheetRow, EXP_COL.BY + 1).getValue());
  expSh.getRange(sheetRow, EXP_COL.BY + 1).setValue(byCell + " (cancelled by " + sess.name + (reason ? ": " + reason : "") + ")");

  return { success: true, message: "Voucher " + existingVoucher + " has been cancelled." };
}

// Returns expense rows, most recent first, optionally restricted to a
// date range (inclusive) for the Expenses tab's own date filter.
function getExpenseList(token, fromDateStr, toDateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var expSh = getSheet(SHEETS.EXPENSES);
  var data = expSh.getDataRange().getValues();

  var from = fromDateStr ? parsePaymentDate_(fromDateStr) : null;
  var to   = toDateStr ? endOfDay_(parsePaymentDate_(toDateStr)) : null;

  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[EXP_COL.VOUCHER]) continue;
    var d = r[EXP_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (from && d.getTime() < from.getTime()) continue;
    if (to && d.getTime() > to.getTime()) continue;
    out.push({
      rowIndex:    i,
      date:        formatDateSafe_(d),
      rawDate:     d.getTime(),
      voucher:     safeStr(r[EXP_COL.VOUCHER]),
      category:    safeStr(r[EXP_COL.CATEGORY]),
      description: safeStr(r[EXP_COL.DESCRIPTION]),
      mode:        safeStr(r[EXP_COL.MODE]),
      amount:      safeNum(r[EXP_COL.AMOUNT]),
      by:          safeStr(r[EXP_COL.BY]),
      bankAccount: safeStr(r[EXP_COL.BANK_ACCOUNT]),
      voided:      isVoided_(r[EXP_COL.VOID])
    });
  }
  out.sort(function(a, b) { return b.rawDate - a.rawDate; });
  return { success: true, data: out };
}

function endOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
}

// ── INCOME & EXPENDITURE ANALYSIS ───────────────────────────
// Combines fee collections (Transactions, income) with Expenses
// (expenditure), excluding TRANSFER_CATEGORIES (Cash Deposit / Cash
// Withdrawal) from "Total Expenditure" since those just move money
// between cash-in-hand and the bank rather than spending it.
function getIncomeExpenditureSummary(token, fromDateStr, toDateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var from = fromDateStr ? parsePaymentDate_(fromDateStr) : null;
  var to   = toDateStr ? endOfDay_(parsePaymentDate_(toDateStr)) : null;

  // Build set of dropout student IDs — their payments are still in the
  // main Transactions sheet but should be labeled separately in P&L.
  var dropoutIds = {};
  var dtxSh = getSheet(SHEETS.DROPOUT_TRANSACTIONS);
  if (dtxSh.getLastRow() > 1) {
    var dtxData = dtxSh.getDataRange().getValues();
    for (var di = 1; di < dtxData.length; di++) {
      var dsid = safeStr(dtxData[di][TX_COL.STUDENT_ID]);
      if (dsid) dropoutIds[dsid] = true;
    }
  }

  // Income side — fee receipts from Transactions (excluding voided rows).
  var txSh = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();
  var totalIncome = 0;
  var activeIncome = 0, dropoutIncome = 0;
  var incomeByMode = {};
  // Fee-head-wise receipts (Notes request: management wants to see WHICH
  // head the money was received against on the Receipts & Payments
  // statement / Balance Sheet, not just a consolidated cash/bank figure).
  // Ordered by first-seen fee head so the statement lists heads in a
  // stable, sensible order rather than alphabetically shuffled each run.
  var incomeByHead = {};
  var incomeHeadOrder = [];
  for (var i = 1; i < txData.length; i++) {
    var t = txData[i];
    if (isVoided_(t[TX_COL.VOID])) continue;
    var d = t[TX_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (from && d.getTime() < from.getTime()) continue;
    if (to && d.getTime() > to.getTime()) continue;
    var amt = safeNum(t[TX_COL.AMOUNT]);
    totalIncome += amt;
    if (dropoutIds[safeStr(t[TX_COL.STUDENT_ID])]) {
      dropoutIncome += amt;
    } else {
      activeIncome += amt;
    }
    var mode = safeStr(t[TX_COL.MODE]) || "Other";
    incomeByMode[mode] = (incomeByMode[mode] || 0) + amt;
    var head = safeStr(t[TX_COL.HEAD]) || "Other";
    if (!incomeByHead.hasOwnProperty(head)) { incomeByHead[head] = 0; incomeHeadOrder.push(head); }
    incomeByHead[head] += amt;
  }
  var incomeByHeadList = incomeHeadOrder.map(function(h) { return { head: h, amount: incomeByHead[h] }; });

  // Expenditure side — Expenses, split into real spend vs cash transfers.
  // The set of transfer categories (isTransfer=TRUE in Settings) is read
  // dynamically, so if the admin adds another transfer-type category later
  // it's excluded from Total Expenditure automatically, with no code change.
  var catRows = getSettingsRows_(SETTINGS_TYPE.EXPENSE_CATEGORY);
  var transferCatSet = {};
  catRows.forEach(function(r) { if (isTrueVal_(r.v1)) transferCatSet[r.key] = true; });

  var expSh = getSheet(SHEETS.EXPENSES);
  var expData = expSh.getDataRange().getValues();
  var totalExpenditure = 0;
  var cashDeposits = 0, cashWithdrawals = 0;
  var expenseByCategory = {};
  for (var k = 1; k < expData.length; k++) {
    var r = expData[k];
    if (!r[EXP_COL.VOUCHER] || isVoided_(r[EXP_COL.VOID])) continue;
    var ed = r[EXP_COL.DATE];
    if (Object.prototype.toString.call(ed) !== "[object Date]") continue;
    if (from && ed.getTime() < from.getTime()) continue;
    if (to && ed.getTime() > to.getTime()) continue;
    var eAmt = safeNum(r[EXP_COL.AMOUNT]);
    var cat  = safeStr(r[EXP_COL.CATEGORY]) || "Others";

    // "Cash Deposit" / "Cash Withdrawal" are the two recognized transfer
    // directions — their amounts feed the bank/cash impact figures.
    // Any OTHER category flagged isTransfer=TRUE in Settings is still
    // excluded from Total Expenditure, but isn't double-counted as a
    // deposit or withdrawal since its direction isn't known.
    if (cat === "Cash Deposit") { cashDeposits += eAmt; continue; }
    if (cat === "Cash Withdrawal") { cashWithdrawals += eAmt; continue; }
    // Bank to Bank Transfer is a protected built-in transfer category (see
    // PROTECTED_TRANSFER_CATEGORIES) — always exclude it from expenditure by
    // name as well as by flag, so a corrupted/duplicated Settings row can
    // never make it leak into the P&L / Balance Sheet as a real expense.
    if (cat === "Bank to Bank Transfer") { continue; }
    if (transferCatSet[cat]) { continue; }

    totalExpenditure += eAmt;
    expenseByCategory[cat] = (expenseByCategory[cat] || 0) + eAmt;
  }

  // Loan interest is a real expense kept in the Loans ledger (not the
  // Expenses sheet) so principal and interest stay separated.
  var _loanInterest = sumLoanInterest_(from, to);
  if (_loanInterest > 0) {
    totalExpenditure += _loanInterest;
    expenseByCategory["Interest on Loans"] = (expenseByCategory["Interest on Loans"] || 0) + _loanInterest;
  }

  return {
    success: true,
    income: { total: totalIncome, activeIncome: activeIncome, dropoutIncome: dropoutIncome, byMode: incomeByMode, byHead: incomeByHeadList },
    expenditure: { total: totalExpenditure, byCategory: expenseByCategory },
    // net here = deposits - withdrawals = the resulting impact on the BANK
    // balance (positive means more was deposited than withdrawn over the
    // period). The impact on cash-in-hand is the exact opposite sign.
    cashTransfers: { deposits: cashDeposits, withdrawals: cashWithdrawals, net: cashDeposits - cashWithdrawals },
    netSurplus: totalIncome - totalExpenditure
  };
}

// ============================================================
//  RECEIPTS TRANSACTION REPORT
//  Date-range report over Transactions: when, how much, by which mode
//  (Cash / Bank). Powers the Financial
//  Analytics "Receipts Transaction Report" tab with a calendar range
//  picker on the client side.
// ============================================================
function getReceiptsReport(token, fromDateStr, toDateStr, modeFilter, headFilter) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var from = fromDateStr ? parsePaymentDate_(fromDateStr) : null;
  var to   = toDateStr ? endOfDay_(parsePaymentDate_(toDateStr)) : null;

  var txSh = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();

  // Built once, used as a fallback below for any row where the name/class
  // stored on the transaction itself is blank (see getReceiptDetail() for
  // why that can happen on older rows).
  var profileById = {};
  getAllStudentProfiles_().forEach(function(p) { profileById[p.studentId] = p; });

  var rows = [];
  var totalsByMode = {};
  var grandTotal = 0;
  var voidedCount = 0;

  for (var i = 1; i < txData.length; i++) {
    var t = txData[i];
    if (!t[TX_COL.RECEIPT]) continue;
    var d = t[TX_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (from && d.getTime() < from.getTime()) continue;
    if (to && d.getTime() > to.getTime()) continue;

    var voided = isVoided_(t[TX_COL.VOID]);
    var mode = safeStr(t[TX_COL.MODE]) || "Other";
    if (modeFilter && mode !== modeFilter) continue;
    var head = safeStr(t[TX_COL.HEAD]);
    if (headFilter && head !== headFilter) continue;

    var sid = safeStr(t[TX_COL.STUDENT_ID]);
    var prof = profileById[sid];
    var amt = safeNum(t[TX_COL.AMOUNT]);
    rows.push({
      rowIndex:  i,
      date:      formatDateSafe_(d),
      rawDate:   d.getTime(),
      studentId: sid,
      name:      safeStr(t[TX_COL.STUDENT_NAME]) || (prof ? prof.name : ""),
      class:     safeStr(t[TX_COL.CLASS]) || (prof ? prof.class : ""),
      receipt:   safeStr(t[TX_COL.RECEIPT]),
      head:      head,
      mode:      mode,
      amount:    amt,
      by:        safeStr(t[TX_COL.BY]),
      bankAccount: safeStr(t[TX_COL.BANK_ACCOUNT]),
      voided:    voided
    });

    if (!voided) {
      totalsByMode[mode] = (totalsByMode[mode] || 0) + amt;
      grandTotal += amt;
    } else {
      voidedCount++;
    }
  }

  rows.sort(function(a, b) { return b.rawDate - a.rawDate; });

  return {
    success: true,
    data: rows,
    summary: {
      grandTotal: grandTotal,
      totalsByMode: totalsByMode,
      count: rows.length - voidedCount,
      voidedCount: voidedCount
    }
  };
}

// Distinct payment modes seen in Transactions, for the report's mode
// filter dropdown — always includes the standard set even if a mode
// hasn't been used yet, so the filter doesn't look empty on a fresh sheet.
function getPaymentModes(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var standard = ["Cash", "Bank"];
  var txSh = getSheet(SHEETS.TRANSACTIONS);
  var txData = txSh.getDataRange().getValues();
  var seen = {};
  standard.forEach(function(m) { seen[m] = true; });
  for (var i = 1; i < txData.length; i++) {
    var m = safeStr(txData[i][TX_COL.MODE]);
    if (m) seen[m] = true;
  }
  return { success: true, data: Object.keys(seen) };
}

// ── EXCEL / CSV EXPORT (Receipts Transaction Report) ─────────
// Builds a CSV string for the same date range + mode filter as
// getReceiptsReport(), which Excel/Google Sheets opens natively — this
// is the standard, dependency-free way to produce an "Excel download"
// from a GAS web app without spinning up and tearing down a temporary
// spreadsheet for every export. The client wraps this in a .csv blob
// and triggers a browser download.
function exportReceiptsCsv(token, fromDateStr, toDateStr, modeFilter, headFilter) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var report = getReceiptsReport(token, fromDateStr, toDateStr, modeFilter, headFilter);
  if (report.error) return report;

  var school = getSchoolInfo(token);
  var lines = [];
  lines.push(csvEscapeRow_([school.name || SCHOOL_NAME]));
  lines.push(csvEscapeRow_(["Receipts Transaction Report"]));
  lines.push(csvEscapeRow_(["Period:", (fromDateStr || "All dates"), "to", (toDateStr || "All dates")]));
  lines.push(csvEscapeRow_(["Mode Filter:", modeFilter || "All Modes", "Head Filter:", headFilter || "All Fee Heads"]));
  lines.push("");
  lines.push(csvEscapeRow_(["Date", "Receipt No", "Student ID", "Student Name", "Class", "Fee Head", "Mode", "Amount", "Recorded By", "Status"]));

  report.data.forEach(function(t) {
    lines.push(csvEscapeRow_([
      t.date, t.receipt, t.studentId, t.name, t.class, t.head, t.mode, t.amount, t.by,
      t.voided ? "Cancelled" : "Active"
    ]));
  });

  lines.push("");
  lines.push(csvEscapeRow_(["", "", "", "", "", "", "Total Collected:", report.summary.grandTotal, "", ""]));
  Object.keys(report.summary.totalsByMode || {}).forEach(function(mode) {
    lines.push(csvEscapeRow_(["", "", "", "", "", "", mode + ":", report.summary.totalsByMode[mode], "", ""]));
  });

  return { success: true, csv: lines.join("\r\n") };
}

function csvEscapeRow_(arr) {
  return arr.map(function(v) {
    var s = safeStr(v);
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }).join(",");
}

function exportExpensesCsv(token, fromDateStr, toDateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var result = getExpenseList(token, fromDateStr, toDateStr);
  if (result.error) return result;

  var school = getSchoolInfo(token);
  var lines = [];
  lines.push(csvEscapeRow_([school.name || SCHOOL_NAME]));
  lines.push(csvEscapeRow_(["Expenses Report"]));
  lines.push(csvEscapeRow_(["Period:", (fromDateStr || "All dates"), "to", (toDateStr || "All dates")]));
  lines.push("");
  lines.push(csvEscapeRow_(["Date", "Voucher No", "Category", "Description", "Mode", "Amount", "Recorded By", "Status"]));

  var total = 0;
  result.data.forEach(function(e) {
    lines.push(csvEscapeRow_([e.date, e.voucher, e.category, e.description, e.mode, e.amount, e.by, e.voided ? "Cancelled" : "Active"]));
    if (!e.voided) total += e.amount;
  });

  lines.push("");
  lines.push(csvEscapeRow_(["", "", "", "", "", "Total:", total, ""]));

  return { success: true, csv: lines.join("\r\n") };
}


/* ==================================================================
   LOANS + BALANCE SHEET  —  add-on for Code.gs
   Paste this ENTIRE block at the bottom of Code.gs.

   Accounting model (kept clean so the Balance Sheet always balances):
     • Loan RECEIVED  -> increases Cash/Bank and the loan liability.
                         It is NOT income (never touches the P&L).
     • Loan REPAYMENT -> principal reduces Cash/Bank and the loan
                         liability (NOT a P&L expense); interest is the
                         only part that is a real expense (P&L).
     • Loans live in their own "Loans" ledger sheet, so principal and
       interest stay separated and nothing distorts fee income.
     • Balance Sheet: Assets = Cash+Bank + Fixed/Current assets;
       Liabilities = Capital Fund (balancing) + Reserves & Surplus
       (incl. current-year surplus) + Loans + Current liabilities.
       A "Capital Fund (opening / balancing)" line makes both sides tie,
       exactly like Tally carries a difference in opening balances.
   ================================================================== */

// ---- schema additions (safe: just adds keys to existing objects) ----
SHEETS.LOANS = "Loans";
SETTINGS_TYPE.LOAN_ACCOUNT  = "LoanAccount";
SETTINGS_TYPE.ASSET_HEAD    = "AssetHead";
SETTINGS_TYPE.LIABILITY_HEAD = "LiabilityHead";

var LOAN_COL = {
  DATE: 0, VOUCHER: 1, TYPE: 2, LOAN_KEY: 3, PRINCIPAL: 4, INTEREST: 5,
  MODE: 6, BANK_ACCOUNT: 7, DESCRIPTION: 8, BY: 9, VOID: 10
};
var LOAN_NUM_COLS = 11;
var LOAN_CATEGORIES  = ["Bank Loan", "Loan from Management", "Unsecured Loan"];
var ASSET_GROUPS     = ["Fixed Asset", "Investment", "Current Asset"];
var LIABILITY_GROUPS = ["Capital Account", "Reserves & Surplus", "Current Liability"];

// ---- low-level helpers -------------------------------------------------
function ensureLoansSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.LOANS);
  if (!sh) sh = ss.insertSheet(SHEETS.LOANS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(["Date", "Voucher No", "Type", "Loan Account", "Principal",
                  "Interest", "Mode", "Bank Account", "Description", "By", "Voided"]);
  }
  return sh;
}

function getLoanVoucherNo_(forDate) {
  var ay = getAcademicYear_(forDate);
  var prefix = getReceiptPrefix_() + "-LOAN/" + ay + "/";
  var sh = ensureLoansSheet_();
  var data = sh.getDataRange().getValues();
  var maxSeq = 0;
  for (var i = 1; i < data.length; i++) {
    var v = safeStr(data[i][LOAN_COL.VOUCHER]);
    if (v.indexOf(prefix) === 0) {
      var seq = parseInt(v.substring(prefix.length), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  return prefix + ("0000" + (maxSeq + 1)).slice(-4);
}

// Sum of loan interest (real expense) over a date window — used by the
// Income & Expenditure summary so interest shows in the P&L.
function sumLoanInterest_(from, to) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.LOANS);
  if (!sh) return 0;
  var data = sh.getDataRange().getValues(), total = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[LOAN_COL.VOUCHER] || isVoided_(r[LOAN_COL.VOID])) continue;
    if (safeStr(r[LOAN_COL.TYPE]) !== "Repayment") continue;
    var d = r[LOAN_COL.DATE];
    if (Object.prototype.toString.call(d) !== "[object Date]") continue;
    if (from && d.getTime() < from.getTime()) continue;
    if (to && d.getTime() > to.getTime()) continue;
    total += safeNum(r[LOAN_COL.INTEREST]);
  }
  return total;
}

// Outstanding per loan account as on a date = opening (Settings) +
// Σ received principal − Σ repaid principal (un-voided, on/before asOn).
function getLoanOutstanding_(asOn) {
  var accounts = {};
  var rows = getSettingsRows_(SETTINGS_TYPE.LOAN_ACCOUNT);
  rows.forEach(function (r) {
    accounts[r.key] = { category: r.v1 || "Unsecured Loan", outstanding: safeNum(r.v2) };
  });
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.LOANS);
  if (sh) {
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var t = data[i];
      if (!t[LOAN_COL.VOUCHER] || isVoided_(t[LOAN_COL.VOID])) continue;
      var d = t[LOAN_COL.DATE];
      if (Object.prototype.toString.call(d) !== "[object Date]") continue;
      if (asOn && d.getTime() > asOn.getTime()) continue;
      var key = safeStr(t[LOAN_COL.LOAN_KEY]);
      if (!accounts[key]) accounts[key] = { category: "Unsecured Loan", outstanding: 0 };
      var p = safeNum(t[LOAN_COL.PRINCIPAL]);
      var typ = safeStr(t[LOAN_COL.TYPE]);
      if (typ === "Received") accounts[key].outstanding += p;
      else if (typ === "Repayment") accounts[key].outstanding -= p;
    }
  }
  return accounts;
}

function financialYearStart_(d) {
  var y = d.getFullYear(), m = d.getMonth();      // Apr (3) … Mar
  var fy = (m >= 3) ? y : y - 1;
  return new Date(fy, 3, 1);
}

// ---- LOAN TRANSACTIONS --------------------------------------------------
function recordLoanReceived(token, payload) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  payload = payload || {};
  var loanKey = safeStr(payload.loanKey);
  if (!loanKey) return { error: "Select a loan account." };
  var amount = safeNum(payload.amount);
  if (amount <= 0) return { error: "Enter a valid loan amount." };
  var mode = safeStr(payload.mode) || "Cash";
  var bankAccount = (mode !== "Cash") ? safeStr(payload.bankAccount) : "";
  if (mode !== "Cash" && !bankAccount) return { error: "Select the bank account the loan was credited to." };
  var d = parsePaymentDate_(payload.dateStr);
  if (d.getTime() > new Date().getTime() + 86400000) return { error: "Date cannot be in the future." };

  var lock = LockService.getScriptLock(), v;
  try {
    lock.waitLock(10000);
    var sh = ensureLoansSheet_();
    v = getLoanVoucherNo_(d);
    sh.appendRow([d, v, "Received", loanKey, amount, 0, mode, bankAccount,
                  safeStr(payload.description), sess.name, false]);
  } finally { lock.releaseLock(); }
  return { success: true, voucherNo: v,
           message: "Loan received " + v + " · \u20B9" + amount + " into " + (mode === "Cash" ? "Cash" : bankAccount) + "." };
}

function recordLoanRepayment(token, payload) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  payload = payload || {};
  var loanKey = safeStr(payload.loanKey);
  if (!loanKey) return { error: "Select a loan account." };
  var principal = safeNum(payload.principal);
  var interest = safeNum(payload.interest);
  if (principal < 0 || interest < 0) return { error: "Amounts cannot be negative." };
  if (principal + interest <= 0) return { error: "Enter a principal and/or interest amount." };
  var mode = safeStr(payload.mode) || "Cash";
  var bankAccount = (mode !== "Cash") ? safeStr(payload.bankAccount) : "";
  if (mode !== "Cash" && !bankAccount) return { error: "Select the bank account the repayment was paid from." };
  var d = parsePaymentDate_(payload.dateStr);
  if (d.getTime() > new Date().getTime() + 86400000) return { error: "Date cannot be in the future." };

  var lock = LockService.getScriptLock(), v;
  try {
    lock.waitLock(10000);
    var sh = ensureLoansSheet_();
    v = getLoanVoucherNo_(d);
    sh.appendRow([d, v, "Repayment", loanKey, principal, interest, mode, bankAccount,
                  safeStr(payload.description), sess.name, false]);
  } finally { lock.releaseLock(); }
  return { success: true, voucherNo: v,
           message: "Repayment " + v + " · principal \u20B9" + principal + (interest ? " + interest \u20B9" + interest : "") + " recorded." };
}

function cancelLoan(token, rowIndex, reason) {
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  if (rowIndex === undefined || rowIndex === null) return { error: "Entry not specified." };
  var sh = ensureLoansSheet_();
  var sr = Number(rowIndex) + 1;
  if (sr < 2 || sr > sh.getLastRow()) return { error: "Loan entry not found." };
  if (!sh.getRange(sr, LOAN_COL.VOUCHER + 1).getValue()) return { error: "Loan entry not found." };
  sh.getRange(sr, LOAN_COL.VOID + 1).setValue(true);
  SpreadsheetApp.flush();
  return { success: true, message: "Loan entry cancelled." };
}

// Ledger + live outstanding summary (grouped by category).
function getLoans(token, fromDateStr, toDateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var from = fromDateStr ? parsePaymentDate_(fromDateStr) : null;
  var to   = toDateStr ? endOfDay_(parsePaymentDate_(toDateStr)) : null;

  var sh = ensureLoansSheet_();
  var data = sh.getDataRange().getValues();
  var ledger = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[LOAN_COL.VOUCHER]) continue;
    var d = r[LOAN_COL.DATE];
    var t = (Object.prototype.toString.call(d) === "[object Date]") ? d.getTime() : 0;
    if (from && t && t < from.getTime()) continue;
    if (to && t && t > to.getTime()) continue;
    ledger.push({
      rowIndex: i,
      date: formatDateSafe_(d),
      voucher: safeStr(r[LOAN_COL.VOUCHER]),
      type: safeStr(r[LOAN_COL.TYPE]),
      loanKey: safeStr(r[LOAN_COL.LOAN_KEY]),
      principal: safeNum(r[LOAN_COL.PRINCIPAL]),
      interest: safeNum(r[LOAN_COL.INTEREST]),
      mode: safeStr(r[LOAN_COL.MODE]),
      bankAccount: safeStr(r[LOAN_COL.BANK_ACCOUNT]),
      description: safeStr(r[LOAN_COL.DESCRIPTION]),
      voided: isVoided_(r[LOAN_COL.VOID])
    });
  }
  ledger.reverse(); // newest first

  var outMap = getLoanOutstanding_(null);
  var outstanding = [], totalOut = 0;
  for (var k in outMap) {
    if (!outMap.hasOwnProperty(k)) continue;
    outstanding.push({ name: k, category: outMap[k].category, outstanding: outMap[k].outstanding });
    totalOut += outMap[k].outstanding;
  }

  return { success: true, ledger: ledger, outstanding: outstanding, totalOutstanding: totalOut };
}

// ---- FINANCE HEADS (Loan accounts / Asset heads / Liability heads) ------
function _financeType_(kind) {
  if (kind === "loan") return SETTINGS_TYPE.LOAN_ACCOUNT;
  if (kind === "asset") return SETTINGS_TYPE.ASSET_HEAD;
  if (kind === "liability") return SETTINGS_TYPE.LIABILITY_HEAD;
  return null;
}
function _financeGroups_(kind) {
  if (kind === "loan") return LOAN_CATEGORIES;
  if (kind === "asset") return ASSET_GROUPS;
  if (kind === "liability") return LIABILITY_GROUPS;
  return [];
}

function getFinanceHeads(token, kind) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var type = _financeType_(kind);
  if (!type) return { error: "Unknown list." };
  var rows = getSettingsRows_(type);
  var out = rows.map(function (r) {
    return { rowIndex: r.rowIndex, name: r.key, group: r.v1, amount: safeNum(r.v2), date: r.v3 };
  });
  return { success: true, kind: kind, groups: _financeGroups_(kind), data: out };
}

function saveFinanceHead(token, kind, payload) {
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  var type = _financeType_(kind);
  if (!type) return { error: "Unknown list." };
  payload = payload || {};
  var name = safeStr(payload.name).trim();
  if (!name) return { error: "Name is required." };
  var group = safeStr(payload.group).trim() || _financeGroups_(kind)[0];
  var amount = safeNum(payload.amount);
  var dateVal = payload.date ? parsePaymentDate_(payload.date) : "";
  ensureSettingsSeeded_();
  var sh = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var rowIndex = (payload.rowIndex === 0 || payload.rowIndex) ? Number(payload.rowIndex) : -1;

  for (var i = 1; i < data.length; i++) {
    if (i === rowIndex) continue;
    var act = data[i][SET_COL.ACTIVE];
    var inactive = (act === false || safeStr(act).toUpperCase() === "FALSE");
    if (!inactive && safeStr(data[i][SET_COL.TYPE]) === type && safeStr(data[i][SET_COL.KEY]) === name) {
      return { error: "\"" + name + "\" already exists." };
    }
  }
  if (rowIndex >= 1) {
    var sr = rowIndex + 1;
    if (sr < 2 || sr > sh.getLastRow() || safeStr(sh.getRange(sr, SET_COL.TYPE + 1).getValue()) !== type) {
      return { error: "Item not found." };
    }
    sh.getRange(sr, SET_COL.KEY + 1, 1, 6).setValues([[name, group, amount, dateVal, "", ""]]);
    sh.getRange(sr, SET_COL.ACTIVE + 1).setValue(true);
    SpreadsheetApp.flush();
    return { success: true, message: "\"" + name + "\" updated." };
  }
  sh.appendRow([type, name, group, amount, dateVal, "", "", true]);
  SpreadsheetApp.flush();
  return { success: true, message: "\"" + name + "\" added." };
}

function deleteFinanceHead(token, kind, rowIndex) {
  var sess = validateSession(token);
  var permErr = requireAdmin_(sess);
  if (permErr) return permErr;
  var sh = getSheet(SHEETS.SETTINGS);
  var sr = Number(rowIndex) + 1;
  if (sr < 2 || sr > sh.getLastRow()) return { error: "Item not found." };
  sh.getRange(sr, SET_COL.ACTIVE + 1).setValue(false);
  SpreadsheetApp.flush();
  return { success: true, message: "Removed." };
}

// ---- BALANCE SHEET --------------------------------------
// Format a Date as yyyy-MM-dd — the ONLY format parsePaymentDate_ accepts.
// (formatDateSafe_ returns "dd-MMM-yyyy", which parsePaymentDate_ cannot
// read and silently treats as today — so never feed its output back into
// a date-parsing function; use this for that.)
function _ymd_(d) {
  if (Object.prototype.toString.call(d) !== "[object Date]" || isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd");
}

function getBalanceSheetData_(token, asOnStr) {
  var asOn = asOnStr ? endOfDay_(parsePaymentDate_(asOnStr)) : endOfDay_(new Date());
  var asOnDisp = formatDateSafe_(asOn);
  var school = getSchoolInfo(token);

  // cash & bank closing as on date (reuse the reconciliation engine)
  var recon = getCashBankReconciliation(token, null, _ymd_(asOn));
  if (recon.error) return recon;

  // current-year surplus (FY-start .. asOn)
  var fyStart = financialYearStart_(asOn);
  var incExp = getIncomeExpenditureSummary(token, _ymd_(fyStart), _ymd_(asOn));
  if (incExp.error) return incExp;
  var surplus = incExp.netSurplus;

  // manual heads
  function group(kind, grp) {
    var rows = getSettingsRows_(_financeType_(kind));
    return rows.filter(function (r) { return (r.v1 || "") === grp; })
               .map(function (r) { return { name: r.key, amount: safeNum(r.v2) }; });
  }
  function sum(arr) { return arr.reduce(function (s, x) { return s + x.amount; }, 0); }

  var capitalHeads   = group("liability", "Capital Account");
  var reservesHeads  = group("liability", "Reserves & Surplus");
  var currLiabHeads  = group("liability", "Current Liability");
  var fixedHeads     = group("asset", "Fixed Asset");
  var investHeads    = group("asset", "Investment");
  var currAssetHeads = group("asset", "Current Asset");

  // loans outstanding grouped by category
  var outMap = getLoanOutstanding_(asOn);
  var loanByCat = {};
  LOAN_CATEGORIES.forEach(function (c) { loanByCat[c] = []; });
  for (var k in outMap) {
    if (!outMap.hasOwnProperty(k)) continue;
    var cat = outMap[k].category;
    if (!loanByCat[cat]) loanByCat[cat] = [];
    if (Math.round(outMap[k].outstanding) !== 0) loanByCat[cat].push({ name: k, amount: outMap[k].outstanding });
  }
  var totalLoans = 0;
  LOAN_CATEGORIES.forEach(function (c) { (loanByCat[c] || []).forEach(function (x) { totalLoans += x.amount; }); });

  // Notes request: outstanding student fee dues (money billed but not yet
  // collected) is a receivable — it belongs on the Assets side, same as
  // cash/bank, since it's value the school is owed. For the CURRENT
  // (active, un-archived) academic year this is necessarily a "current,
  // as of today" figure — fee billing isn't stored with a per-transaction
  // date the way payments are, so it can't be reconstructed for an
  // arbitrary past date within a year that's still live. But once a year
  // has been archived (Settings → Promote Students), its true year-end
  // Fee Receivable is frozen in that Students_<year> snapshot, so a
  // Balance Sheet run for a past year now reads THAT instead of today's
  // live figure — see getFeeReceivableAsOf_().
  var feeReceivableInfo = getFeeReceivableAsOf_(token, asOn);
  var feeReceivable = feeReceivableInfo.total;

  // assets
  var cashInHand = recon.cash.closing;
  var bankRows = recon.banks.map(function (b) { return { name: b.label, amount: b.closing }; });
  var totalBank = bankRows.reduce(function (s, b) { return s + b.amount; }, 0);
  var totalCurrentAssets = cashInHand + totalBank + feeReceivable + sum(currAssetHeads);
  var totalAssets = sum(fixedHeads) + sum(investHeads) + totalCurrentAssets;

  // liabilities — capital fund is the balancing figure. Adding
  // feeReceivable above increases totalAssets, which (since knownLiab is
  // unchanged) automatically increases Capital Fund by the same amount —
  // so the school's recorded net worth grows to include money it's owed,
  // not just cash already in hand, with no separate adjustment needed.
  var knownLiab = sum(capitalHeads) + sum(reservesHeads) + surplus + totalLoans + sum(currLiabHeads);
  var capitalFund = totalAssets - knownLiab;       // makes both sides tie (Tally-style)
  var totalLiabilities = totalAssets;              // by construction

  return {
    success: true,
    school: { name: school.name, place: school.address || "", logo: school.logo || "" },
    asOn: asOnDisp,
    fyLabel: formatDateSafe_(fyStart) + " to " + asOnDisp,
    liabilities: {
      capitalHeads: capitalHeads, capitalFund: capitalFund,
      reservesHeads: reservesHeads, surplus: surplus,
      loanByCat: loanByCat, currLiabHeads: currLiabHeads
    },
    assets: {
      fixedHeads: fixedHeads, investHeads: investHeads,
      cashInHand: cashInHand, bankRows: bankRows,
      feeReceivable: feeReceivable, feeReceivableByHead: feeReceivableInfo.byHead,
      feeReceivableHeadWiseAvailable: feeReceivableInfo.headWiseAvailable,
      currAssetHeads: currAssetHeads
    },
    totalLiabilities: totalLiabilities, totalAssets: totalAssets
  };
}

function getBalanceSheet(token, asOnStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var data = getBalanceSheetData_(token, asOnStr);
  if (data.error) return data;
  var built = buildBalanceSheetHtml_(data, function (x) { return escHtml_(x); });
  return { success: true, css: built.css, html: built.body,
           totalAssets: data.totalAssets, totalLiabilities: data.totalLiabilities,
           fileName: "Balance_Sheet_" + safeStr(data.asOn).replace(/-/g, "") };
}

function getBalanceSheetPdf(token, asOnStr) {
  var r = getBalanceSheet(token, asOnStr);
  if (r.error) return r;
  var fullHtml = "<!doctype html><html><head><meta charset='utf-8'><style>" + r.css +
                 "</style></head><body>" + r.html + "</body></html>";
  var bytes;
  try { bytes = htmlToPdfBlob_(fullHtml, r.fileName); }
  catch (e) { return { error: "Could not generate the PDF: " + e.message }; }
  return { success: true, fileName: r.fileName + ".pdf", pdfBase64: Utilities.base64Encode(bytes) };
}

function _bsMoney_(n) {
  n = Math.round(n || 0);
  var neg = n < 0; n = Math.abs(n);
  var s = String(n), l3 = s.slice(-3), r = s.slice(0, -3);
  if (r) l3 = "," + l3;
  r = r.replace(/\B(?=(\d{2})+(?!\d))/g, ",");
  var out = "\u20B9" + r + l3;
  return neg ? "(" + out + ")" : out;
}

function buildBalanceSheetHtml_(data, esc) {
  esc = esc || function (x) { return String(x == null ? "" : x); };
  var NAVY = "#0D2137", TEAL = "#0A7E8C", GOLD = "#F0C040", LINE = "#9AA9B2",
      GREY = "#EEF3F5", RED = "#C0392B";
  var FF = "Calibri,'Segoe UI',Arial,sans-serif", SERIF = "Cambria,Georgia,'Times New Roman',serif";
  var cellL = "border:1px solid " + LINE + ";padding:3px 9px;font-size:11.5px;font-family:" + FF + ";";
  var cellR = cellL + "text-align:right;white-space:nowrap;";

  // build a side as rows: {t:'group'|'sub'|'item'|'blank', label, amount}
  function sideRows(which) {
    var rows = [], L = data.liabilities, A = data.assets;
    if (which === "L") {
      // Capital Account
      rows.push({ t: "group", label: "Capital Account" });
      L.capitalHeads.forEach(function (x) { rows.push({ t: "item", label: x.name, amount: x.amount }); });
      rows.push({ t: "item", label: "Capital Fund (opening / balancing)", amount: L.capitalFund });
      // Reserves & Surplus
      rows.push({ t: "group", label: "Reserves & Surplus" });
      L.reservesHeads.forEach(function (x) { rows.push({ t: "item", label: x.name, amount: x.amount }); });
      rows.push({ t: "item", label: "Surplus / (Deficit) for the year", amount: L.surplus });
      // Loans
      var anyLoan = false;
      LOAN_CATEGORIES.forEach(function (c) { if ((L.loanByCat[c] || []).length) anyLoan = true; });
      if (anyLoan) {
        rows.push({ t: "group", label: "Loans (Liability)" });
        LOAN_CATEGORIES.forEach(function (c) {
          var arr = L.loanByCat[c] || [];
          if (!arr.length) return;
          rows.push({ t: "subhead", label: c });
          arr.forEach(function (x) { rows.push({ t: "item", label: "   " + x.name, amount: x.amount }); });
        });
      }
      // Current Liabilities
      if (L.currLiabHeads.length) {
        rows.push({ t: "group", label: "Current Liabilities" });
        L.currLiabHeads.forEach(function (x) { rows.push({ t: "item", label: x.name, amount: x.amount }); });
      }
    } else {
      var A2 = data.assets;
      if (A2.fixedHeads.length) {
        rows.push({ t: "group", label: "Fixed Assets" });
        A2.fixedHeads.forEach(function (x) { rows.push({ t: "item", label: x.name, amount: x.amount }); });
      }
      if (A2.investHeads.length) {
        rows.push({ t: "group", label: "Investments" });
        A2.investHeads.forEach(function (x) { rows.push({ t: "item", label: x.name, amount: x.amount }); });
      }
      rows.push({ t: "group", label: "Current Assets" });
      rows.push({ t: "item", label: "Cash-in-Hand", amount: A2.cashInHand });
      if (A2.bankRows.length) {
        rows.push({ t: "subhead", label: "Bank Accounts" });
        A2.bankRows.forEach(function (x) { rows.push({ t: "item", label: "   " + x.name, amount: x.amount }); });
      }
      if (A2.feeReceivable) {
        if (A2.feeReceivableHeadWiseAvailable && A2.feeReceivableByHead && A2.feeReceivableByHead.length) {
          // Notes request: show WHICH fee head the outstanding is against,
          // not just one consolidated number — a subhead banner plus one
          // indented line per head, then the head-wise total.
          rows.push({ t: "subhead", label: "Fee Receivable (Outstanding Student Dues) \u2014 Head-wise" });
          A2.feeReceivableByHead.forEach(function (x) {
            rows.push({ t: "item", label: "   " + x.head, amount: x.amount });
          });
          rows.push({ t: "item", label: "Total Fee Receivable", amount: A2.feeReceivable });
        } else {
          // Archived academic year — the frozen snapshot only stored each
          // student's total Due, not a per-head split, so only the
          // consolidated figure is available for this "As On" date.
          rows.push({ t: "item", label: "Fee Receivable (Outstanding Student Dues)", amount: A2.feeReceivable });
        }
      }
      A2.currAssetHeads.forEach(function (x) { rows.push({ t: "item", label: x.name, amount: x.amount }); });
    }
    return rows;
  }

  function renderSide(title, rows, total) {
    var h = '<td valign="top" style="width:50%;padding:0;border:1px solid ' + LINE + ';">';
    h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">';
    h += '<tr><td style="' + cellL + 'background:' + NAVY + ';color:#fff;font-weight:bold;">' + title + '</td>'
       + '<td style="' + cellR + 'background:' + NAVY + ';color:#fff;font-weight:bold;width:120px;">Amount</td></tr>';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r.t === "group") {
        h += '<tr><td colspan="2" style="' + cellL + 'background:' + TEAL + ';color:#fff;font-weight:bold;font-family:' + SERIF + ';">' + esc(r.label) + '</td></tr>';
      } else if (r.t === "subhead") {
        h += '<tr><td colspan="2" style="' + cellL + 'background:' + GREY + ';font-weight:bold;color:' + NAVY + ';">' + esc(r.label) + '</td></tr>';
      } else {
        var amtCol = (r.amount < 0) ? RED : NAVY;
        h += '<tr><td style="' + cellL + '">' + esc(r.label) + '</td>'
           + '<td style="' + cellR + 'color:' + amtCol + ';">' + _bsMoney_(r.amount) + '</td></tr>';
      }
    }
    h += '</table></td>';
    return h;
  }

  var lRows = sideRows("L"), rRows = sideRows("R");
  var h = "";
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-family:' + FF + ';">';
  h += '<tr><td style="border:2px solid ' + NAVY + ';padding:12px 16px 10px;">';

  // header
  h += '<div style="text-align:center;border-bottom:2px solid ' + GOLD + ';padding-bottom:6px;margin-bottom:10px;">';
  h += '<div style="font-size:22px;font-weight:bold;color:' + NAVY + ';font-family:' + SERIF + ';">' + esc(data.school.name) + '</div>';
  if (data.school.place) h += '<div style="font-size:12px;color:#5E6E7C;text-transform:uppercase;letter-spacing:1px;">' + esc(data.school.place) + '</div>';
  h += '<div style="margin-top:4px;"><span style="background:' + TEAL + ';color:#fff;font-size:12px;font-weight:bold;padding:3px 16px;border-radius:12px;">Balance Sheet as on ' + esc(data.asOn) + '</span></div>';
  h += '</div>';

  // two columns
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;table-layout:fixed;"><tr>';
  h += renderSide("LIABILITIES", lRows);
  h += renderSide("ASSETS", rRows);
  h += '</tr>';
  // totals row (equal by construction)
  h += '<tr>';
  h += '<td style="border:1px solid ' + LINE + ';padding:0;"><table width="100%" style="border-collapse:collapse;"><tr>'
     + '<td style="' + cellL + 'background:' + GOLD + ';font-weight:bold;color:' + NAVY + ';">Total</td>'
     + '<td style="' + cellR + 'background:' + GOLD + ';font-weight:bold;color:' + NAVY + ';width:120px;">' + _bsMoney_(data.totalLiabilities) + '</td></tr></table></td>';
  h += '<td style="border:1px solid ' + LINE + ';padding:0;"><table width="100%" style="border-collapse:collapse;"><tr>'
     + '<td style="' + cellL + 'background:' + GOLD + ';font-weight:bold;color:' + NAVY + ';">Total</td>'
     + '<td style="' + cellR + 'background:' + GOLD + ';font-weight:bold;color:' + NAVY + ';width:120px;">' + _bsMoney_(data.totalAssets) + '</td></tr></table></td>';
  h += '</tr></table>';

  // notes
  h += '<div style="margin-top:8px;font-size:10.5px;color:#5E6E7C;font-style:italic;">'
     + 'Cash &amp; bank balances are live as on the date above. Surplus is for the current financial year (' + esc(data.fyLabel) + '). '
     + 'Loan outstanding = opening + received \u2212 principal repaid. "Capital Fund (opening / balancing)" carries any opening-balance difference so the two sides tie, as in Tally.'
     + '</div>';

  h += '</td></tr></table>';

  var css = ''
    + '@page{size:A4 portrait;margin:9mm;}'
    + 'html,body{margin:0;padding:0;background:#fff;font-family:' + FF + ';color:#16242F;}'
    + '@media print{.noprint{display:none!important;}}';
  return { css: css, body: h };
}


/* ==================================================================
   U-DISE CASTE CENSUS  +  STUDENT DATA AUDIT  —  add-on for Code.gs
   Paste this ENTIRE block at the bottom of your existing Code.gs.
   Reuses existing helpers (validateSession, safeStr, escHtml_, SHEETS,
   SP_COL, TEACHER_COL, SETTINGS_TYPE, getSettingsRows_, getSchoolInfo,
   getCurrentAcademicYear_, getAllStudentProfiles_, getAllTeacherProfiles_,
   htmlToPdfBlob_). Nothing is redefined — no name clashes.

   Caste rule: SC/SC-A.. -> SC, ST -> ST, BC-A..E/OBC/MBC -> BC,
   OC/General/FC -> OC.  Blank / unrecognised caste is NOT counted as
   OC — it is left out of the census and listed in the Data Audit so
   the office can fix it.
   ================================================================== */

/* ============================================================
   CASTE-WISE CENSUS — shared core (ES5, Apps-Script compatible)
   These functions are pasted verbatim into Code.gs. They are kept
   ES5 (var, no arrow fns, no template literals) so the SAME code
   runs in Google Apps Script and in this Node preview.
   ============================================================ */

// Normalise any caste / sub-caste spelling into one of the four
// statutory U-Dise buckets. Examples handled:
//   SC, SC-A, SC-B, SC-C, "Scheduled Caste"        -> SC
//   ST, ST-A, "Scheduled Tribe"                     -> ST
//   BC, BC-A..BC-E, OBC, MBC, EBC, "Backward Class" -> BC
//   OC, FC, GEN, General, Open, blank/unknown       -> OC
function casteCategory_(caste, subCaste) {
  var raw = String(caste == null ? "" : caste);
  if (!raw.replace(/[^A-Za-z]/g, "")) raw = String(subCaste == null ? "" : subCaste);
  var s = raw.toUpperCase().replace(/[^A-Z]/g, ""); // strip spaces, hyphens, digits
  if (!s) return "";                                   // no caste recorded -> NOT counted (left blank)
  if (s.indexOf("SCHEDULEDTRIBE") === 0) return "ST";  // match ST word before the "SC..." rule
  if (s.indexOf("SCHEDULEDCASTE") === 0 || s.indexOf("SC") === 0) return "SC";
  if (s.indexOf("ST") === 0) return "ST";
  if (s.indexOf("BC") === 0 || s.indexOf("OBC") === 0 || s.indexOf("MBC") === 0 ||
      s.indexOf("EBC") === 0 || s.indexOf("BACKWARD") === 0) return "BC";
  if (s.indexOf("OC") === 0 || s.indexOf("OPEN") === 0 || s.indexOf("GEN") === 0 ||
      s.indexOf("GENERAL") === 0 || s === "FC" || s.indexOf("FORWARD") === 0) return "OC";
  return "";                                           // unrecognised caste -> NOT counted (only real data is tallied)
}

// Is a government-ID / data field actually filled in? Treats common
// placeholders ("-", "NA", "Nil", "0", "None") as not-present.
function fieldPresent_(v) {
  var s = String(v == null ? "" : v).trim().toLowerCase();
  if (!s) return false;
  if (s === "-" || s === "--" || s === "na" || s === "n/a" || s === "nil" ||
      s === "null" || s === "none" || s === "0") return false;
  return true;
}

// Boys / Girls bucket. Returns "B", "G", or "" (skip — e.g. "Other").
function genderBucket_(g) {
  var s = String(g == null ? "" : g).trim().toUpperCase();
  if (s === "MALE" || s === "BOY" || s === "M" || s === "BOYS") return "B";
  if (s === "FEMALE" || s === "GIRL" || s === "F" || s === "GIRLS") return "G";
  return "";
}

// Group a class name into the three census sections.
function classGroup_(name) {
  var n = String(name == null ? "" : name).trim();
  var up = n.toUpperCase();
  var num = parseInt(n, 10);
  if (!isNaN(num) && /^[0-9]+$/.test(n)) {
    if (num <= 5) return "Primary";
    return "High School"; // 6,7,8,9,10 (and 11/12 if a school adds them)
  }
  if (/NUR|LKG|UKG|PRE|KG|PP|MONT/.test(up)) return "Pre-Primary";
  return "Pre-Primary"; // any other non-numeric label
}

// Is this staff member teaching? Uses Department first, then Designation
// keywords. Defaults to "teaching" when nothing matches.
function isTeachingStaff_(department, designation) {
  var dep = String(department == null ? "" : department).toLowerCase();
  var des = String(designation == null ? "" : designation).toLowerCase();
  if (dep.indexOf("non") >= 0) return false;                 // "Non-Teaching"
  if (dep.indexOf("teach") >= 0) return true;                 // "Teaching"
  if (dep.indexOf("admin") >= 0 || dep.indexOf("office") >= 0 ||
      dep.indexOf("support") >= 0 || dep.indexOf("ministerial") >= 0) return false;
  var nonTeach = ["accountant", "clerk", "attender", "attendant", "watchman", "watch man",
    "sweeper", "scavenger", "ayah", "aaya", "helper", "peon", "driver", "gardener",
    "security", "cook", "office", "lab assistant", "data entry", "computer operator"];
  for (var i = 0; i < nonTeach.length; i++) if (des.indexOf(nonTeach[i]) >= 0) return false;
  var teach = ["teacher", "sgt", "school assistant", "pet", "tgt", "pgt", "lecturer",
    "headmaster", "head master", "head mistress", "principal", "pandit", "craft",
    "drawing", "music", "lp", "up ", "sa "];
  for (var j = 0; j < teach.length; j++) if (des.indexOf(teach[j]) >= 0) return true;
  return true; // sensible default for a school roster
}

// Active student? (present strength). Removed/TC/Dropout are excluded.
function isActiveStudent_(status) {
  var s = String(status == null ? "" : status).trim().toLowerCase();
  return s === "" || s === "active";
}

// ---------- HTML escape (kept local so the Node preview is standalone;
// in Code.gs the project's existing escHtml_ is used instead) ----------
function _censusEsc(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/* ------------------------------------------------------------
   buildCasteCensusHtml_(data, esc)
   Builds the single-A4 census report. `esc` lets Code.gs pass its
   own escHtml_; defaults to the local _censusEsc for the Node preview.
   data = {
     school:{name, place, logo}, academicYear, asOf,
     staff:{ teaching:{male,female}, nonTeaching:{male,female} },
     groups:[ { label, rows:[ {cls, SC:{b,g}, ST:{b,g}, BC:{b,g}, OC:{b,g}} ] } ]
   }
   ------------------------------------------------------------ */
function buildCasteCensusHtml_(data, esc) {
  esc = esc || _censusEsc;
  var CATS = ["SC", "ST", "BC", "OC"];
  var NAVY = "#0D2137", TEAL = "#0A7E8C", GOLD = "#F0C040",
      LINE = "#9AA9B2", GREY = "#EEF3F5", ZEBRA = "#F4F9FA";
  var FF = "Calibri,'Segoe UI',Arial,sans-serif", SERIF = "Cambria,Georgia,'Times New Roman',serif";

  // reusable inline cell styles (inline = renders in browser print, Google
  // Docs PDF export, and any other consumer — no reliance on <style>)
  var base   = "border:1px solid " + LINE + ";padding:2px 7px;font-size:11px;line-height:1.15;text-align:center;font-family:" + FF + ";";
  var head   = base + "background:" + NAVY + ";color:#ffffff;font-weight:bold;";
  var tealh  = base + "background:" + TEAL + ";color:#ffffff;font-weight:bold;";
  var banner = base + "background:" + TEAL + ";color:#ffffff;font-weight:bold;letter-spacing:1px;font-family:" + SERIF + ";text-transform:uppercase;";
  var clsCell= base + "text-align:left;padding-left:12px;font-weight:bold;color:" + NAVY + ";white-space:nowrap;";
  var rtot   = base + "font-weight:bold;background:" + GREY + ";";
  var gold   = base + "background:" + GOLD + ";color:" + NAVY + ";font-weight:bold;";

  var colTot = { SC: { b: 0, g: 0 }, ST: { b: 0, g: 0 }, BC: { b: 0, g: 0 }, OC: { b: 0, g: 0 } };
  function rowTotal(r) { var tt = 0; for (var i = 0; i < CATS.length; i++) { var c = r[CATS[i]] || { b: 0, g: 0 }; tt += (c.b || 0) + (c.g || 0); } return tt; }
  var dcell = function (n) { return n > 0 ? n : ""; };

  var h = "";
  // ---- frame (outer bordered table) ----
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-family:' + FF + ';">';
  h += '<tr><td style="border:2px solid ' + NAVY + ';padding:9px 16px 7px;">';

  // ---- title row (logo right) ----
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;"><tr>';
  h += '<td width="64"></td><td align="center" style="padding-bottom:4px;border-bottom:2px solid ' + GOLD + ';">';
  h += '<div style="font-size:20px;font-weight:bold;letter-spacing:.4px;color:' + NAVY + ';font-family:' + SERIF + ';">' + esc(data.school.name) + '</div>';
  if (data.school.place) h += '<div style="font-size:13px;color:#5E6E7C;letter-spacing:1px;text-transform:uppercase;margin-top:2px;">' + esc(data.school.place) + '</div>';
  h += '<div style="font-size:13px;font-weight:bold;color:' + NAVY + ';margin-top:3px;">ACADEMIC YEAR 20' + esc(data.academicYear) + '</div>';
  h += '<div style="margin-top:4px;"><span style="background:' + TEAL + ';color:#fff;font-size:12px;font-weight:bold;padding:3px 16px;border-radius:12px;letter-spacing:.4px;">Caste-wise Strength Report &mdash; Boys &amp; Girls</span></div>';
  h += '</td><td width="64" align="right" valign="top">';
  if (data.school.logo) h += '<img src="' + esc(data.school.logo) + '" width="54" height="54" style="border-radius:50%;object-fit:contain;"/>';
  h += '</td></tr></table>';

  // ---- staff block ----
  var s = data.staff;
  h += '<div style="height:8px;"></div>';
  h += '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;">';
  h += '<tr><td width="170" nowrap="nowrap" style="' + head + '"></td><td width="72" nowrap="nowrap" style="' + head + '">Male</td><td width="78" nowrap="nowrap" style="' + head + '">Female</td><td width="66" nowrap="nowrap" style="' + head + '">Total</td></tr>';
  h += '<tr><td nowrap="nowrap" style="' + base + 'text-align:left;padding-left:12px;font-weight:bold;color:' + NAVY + ';">Teaching Staff</td><td style="' + base + '">' + s.teaching.male + '</td><td style="' + base + '">' + s.teaching.female + '</td><td style="' + rtot + '">' + (s.teaching.male + s.teaching.female) + '</td></tr>';
  h += '<tr><td nowrap="nowrap" style="' + base + 'text-align:left;padding-left:12px;font-weight:bold;color:' + NAVY + ';">Non-Teaching Staff</td><td style="' + base + '">' + s.nonTeaching.male + '</td><td style="' + base + '">' + s.nonTeaching.female + '</td><td style="' + rtot + '">' + (s.nonTeaching.male + s.nonTeaching.female) + '</td></tr>';
  h += '</table>';

  // ---- census table ----
  h += '<div style="height:8px;"></div>';
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">';
  h += '<tr><td rowspan="2" style="' + head + 'vertical-align:middle;">Class</td>';
  for (var ci = 0; ci < CATS.length; ci++) h += '<td colspan="2" style="' + head + 'font-size:13px;letter-spacing:.5px;">' + CATS[ci] + '</td>';
  h += '<td rowspan="2" style="' + tealh + 'vertical-align:middle;">TOTAL</td></tr>';
  h += '<tr>';
  for (var cj = 0; cj < CATS.length; cj++) h += '<td style="' + head + '">Boys</td><td style="' + head + '">Girls</td>';
  h += '</tr>';

  for (var gi = 0; gi < data.groups.length; gi++) {
    var grp = data.groups[gi];
    if (!grp.rows.length) continue;
    h += '<tr><td colspan="10" style="' + banner + '">' + esc(grp.label) + '</td></tr>';
    for (var ri = 0; ri < grp.rows.length; ri++) {
      var r = grp.rows[ri];
      var zb = (ri % 2 === 1) ? "background:" + ZEBRA + ";" : "";
      h += '<tr><td nowrap="nowrap" style="' + clsCell + zb + '">' + esc(r.cls) + '</td>';
      for (var k = 0; k < CATS.length; k++) {
        var c = r[CATS[k]] || { b: 0, g: 0 };
        colTot[CATS[k]].b += (c.b || 0); colTot[CATS[k]].g += (c.g || 0);
        h += '<td style="' + base + zb + '">' + dcell(c.b || 0) + '</td><td style="' + base + zb + '">' + dcell(c.g || 0) + '</td>';
      }
      var rt = rowTotal(r);
      h += '<td style="' + rtot + '">' + (rt > 0 ? rt : "") + '</td></tr>';
    }
  }

  var gBoys = 0, gGirls = 0;
  h += '<tr><td style="' + gold + 'text-align:left;padding-left:12px;">TOTAL</td>';
  for (var m = 0; m < CATS.length; m++) { var ct = colTot[CATS[m]]; gBoys += ct.b; gGirls += ct.g; h += '<td style="' + gold + '">' + ct.b + '</td><td style="' + gold + '">' + ct.g + '</td>'; }
  h += '<td style="' + gold + 'font-size:14px;">' + (gBoys + gGirls) + '</td></tr>';
  h += '</table>';

  // ---- footer ----
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin-top:8px;"><tr>';
  h += '<td style="font-size:12px;font-family:' + FF + ';color:' + NAVY + ';"><b>Total Boys:</b> ' + gBoys + ' &nbsp;&nbsp; <b>Total Girls:</b> ' + gGirls + ' &nbsp;&nbsp; <b>Grand Total:</b> ' + (gBoys + gGirls) + '</td>';
  h += '<td align="right" style="font-size:11px;font-style:italic;color:#5E6E7C;font-family:' + FF + ';">Generated on ' + esc(data.asOf) + '</td>';
  h += '</tr></table>';

  if (data.uncounted) {
    h += '<div style="margin-top:5px;font-size:11px;color:#C0392B;font-style:italic;font-family:' + FF + ';">'
       + 'Note: ' + data.uncounted + ' active student(s) are not shown above because their caste is not recorded. '
       + 'Open the Data Audit report to find and update them.</div>';
  }

  h += '</td></tr></table>'; // frame

  var css = ''
    + '@page{size:A4 portrait;margin:6mm;}'
    + 'html,body{margin:0;padding:0;background:#ffffff;font-family:' + FF + ';color:#16242F;}'
    + '@media print{.noprint{display:none!important;}}';

  return { css: css, body: h };
}

/* ------------------------------------------------------------
   buildDataAuditHtml_(data, esc)
   Student data-completeness report: how many active students have
   Aadhar / PEN / APAAR / Caste on file, plus a list of every student
   who is missing one or more — so the office knows whose data to update.
   data = {
     school:{name,place,logo}, academicYear, asOf,
     summary:{ total, aadhar, pen, apaar, caste },     // available counts
     missing:[ {studentId, name, cls, aadhar, pen, apaar, caste, mobile} ] // booleans = present
   }
   ------------------------------------------------------------ */
function buildDataAuditHtml_(data, esc) {
  esc = esc || _censusEsc;
  var NAVY = "#0D2137", TEAL = "#0A7E8C", GOLD = "#F0C040", GREEN = "#2E9E5B",
      RED = "#C0392B", LINE = "#9AA9B2", ZEBRA = "#F4F9FA";
  var FF = "Calibri,'Segoe UI',Arial,sans-serif", SERIF = "Cambria,Georgia,'Times New Roman',serif";
  var base = "border:1px solid " + LINE + ";padding:4px 7px;font-size:11px;font-family:" + FF + ";";
  var head = base + "background:" + NAVY + ";color:#fff;font-weight:bold;text-align:center;";

  function pct(n, d) { return d > 0 ? Math.round(n / d * 100) : 0; }
  var sm = data.summary, total = sm.total || 0;

  var h = "";
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-family:' + FF + ';">';
  h += '<tr><td style="border:2px solid ' + NAVY + ';padding:10px 16px 8px;">';

  // title
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;"><tr>';
  h += '<td width="64"></td><td align="center" style="padding-bottom:4px;border-bottom:2px solid ' + GOLD + ';">';
  h += '<div style="font-size:20px;font-weight:bold;letter-spacing:.4px;color:' + NAVY + ';font-family:' + SERIF + ';">' + esc(data.school.name) + '</div>';
  if (data.school.place) h += '<div style="font-size:12px;color:#5E6E7C;letter-spacing:1px;text-transform:uppercase;">' + esc(data.school.place) + '</div>';
  h += '<div style="font-size:12px;font-weight:bold;color:' + NAVY + ';margin-top:2px;">ACADEMIC YEAR 20' + esc(data.academicYear) + '</div>';
  h += '<div style="margin-top:4px;"><span style="background:' + TEAL + ';color:#fff;font-size:11px;font-weight:bold;padding:3px 16px;border-radius:12px;">Student Data Audit &mdash; Missing Government IDs</span></div>';
  h += '</td><td width="64" align="right" valign="top">';
  if (data.school.logo) h += '<img src="' + esc(data.school.logo) + '" width="50" height="50" style="border-radius:50%;object-fit:contain;"/>';
  h += '</td></tr></table>';

  // summary cards
  function cardCell(label, avail, isTotal) {
    var miss = total - avail, p = pct(avail, total);
    var s = '<td style="' + base + 'text-align:center;width:20%;vertical-align:top;padding:8px 6px;">';
    s += '<div style="font-size:10.5px;color:#5E6E7C;font-weight:bold;text-transform:uppercase;letter-spacing:.4px;">' + esc(label) + '</div>';
    if (isTotal) {
      s += '<div style="font-size:24px;font-weight:bold;color:' + NAVY + ';font-family:' + SERIF + ';margin-top:3px;">' + avail + '</div>';
      s += '<div style="font-size:10.5px;color:#5E6E7C;">active students</div>';
    } else {
      var col = p >= 100 ? GREEN : (p >= 75 ? TEAL : (p >= 40 ? "#C8941A" : RED));
      s += '<div style="font-size:22px;font-weight:bold;color:' + col + ';font-family:' + SERIF + ';margin-top:3px;">' + avail + '<span style="font-size:12px;color:#5E6E7C;"> / ' + total + '</span></div>';
      s += '<div style="font-size:10.5px;color:#5E6E7C;">' + p + '% on file</div>';
      s += '<div style="font-size:10.5px;color:' + (miss ? RED : GREEN) + ';font-weight:bold;">' + (miss ? miss + ' missing' : 'complete') + '</div>';
    }
    return s + '</td>';
  }
  h += '<div style="height:10px;"></div>';
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;"><tr>';
  h += cardCell("Active Students", total, true);
  h += cardCell("Aadhar", sm.aadhar || 0, false);
  h += cardCell("PEN", sm.pen || 0, false);
  h += cardCell("APAAR", sm.apaar || 0, false);
  h += cardCell("Caste", sm.caste || 0, false);
  h += '</tr></table>';

  // missing list
  h += '<div style="height:12px;"></div>';
  var miss = data.missing || [];
  if (!miss.length) {
    h += '<div style="border:1px solid ' + GREEN + ';background:#EAF7EF;color:' + GREEN + ';font-weight:bold;text-align:center;padding:14px;border-radius:6px;font-family:' + FF + ';">All active students have Aadhar, PEN, APAAR and Caste on file. Nothing to update.</div>';
  } else {
    h += '<div style="font-size:12px;font-weight:bold;color:' + NAVY + ';font-family:' + SERIF + ';margin-bottom:5px;">Students with missing details (' + miss.length + ')</div>';
    h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;">';
    h += '<thead><tr>';
    var cols = ["#", "Student ID", "Name", "Class", "Aadhar", "PEN", "APAAR", "Caste", "Father Mobile"];
    var al = ["center", "left", "left", "center", "center", "center", "center", "center", "left"];
    for (var i = 0; i < cols.length; i++) h += '<td style="' + head + '">' + cols[i] + '</td>';
    h += '</tr></thead><tbody>';
    function mark(ok) {
      return ok ? '<span style="color:' + GREEN + ';font-weight:bold;">&#10004;</span>'
                : '<span style="color:' + RED + ';font-weight:bold;">&#10008;</span>';
    }
    for (var r = 0; r < miss.length; r++) {
      var m = miss[r], zb = (r % 2 === 1) ? "background:" + ZEBRA + ";" : "";
      var cells = [
        String(r + 1), esc(m.studentId), esc(m.name), esc(m.cls),
        mark(m.aadhar), mark(m.pen), mark(m.apaar), mark(m.caste), esc(m.mobile || "")
      ];
      h += '<tr>';
      for (var c = 0; c < cells.length; c++) {
        h += '<td style="' + base + zb + 'text-align:' + al[c] + ';' + (c === 2 ? 'font-weight:bold;color:' + NAVY + ';' : '') + '">' + cells[c] + '</td>';
      }
      h += '</tr>';
    }
    h += '</tbody></table>';
  }

  // footer
  h += '<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin-top:10px;"><tr>';
  h += '<td style="font-size:11px;color:' + NAVY + ';font-family:' + FF + ';"><b>&#10004;</b> on file &nbsp;&nbsp; <b style="color:' + RED + ';">&#10008;</b> missing &mdash; update these in Student Records.</td>';
  h += '<td align="right" style="font-size:11px;font-style:italic;color:#5E6E7C;font-family:' + FF + ';">Generated on ' + esc(data.asOf) + '</td>';
  h += '</tr></table>';

  h += '</td></tr></table>';

  var css = ''
    + '@page{size:A4 portrait;margin:8mm;}'
    + 'html,body{margin:0;padding:0;background:#fff;font-family:' + FF + ';color:#16242F;}'
    + 'thead{display:table-header-group;}'
    + '@media print{.noprint{display:none!important;}}';
  return { css: css, body: h };
}

/* ---- PUBLIC ENDPOINT: caste census data + single-A4 HTML --------- */
function getCasteCensusReport(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var school = getSchoolInfo(token);
  var ay = (typeof getCurrentAcademicYear_ === "function") ? getCurrentAcademicYear_() : "";
  var tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  var asOf = Utilities.formatDate(new Date(), tz, "dd-MMM-yyyy");

  function blank() { return { SC:{b:0,g:0}, ST:{b:0,g:0}, BC:{b:0,g:0}, OC:{b:0,g:0} }; }
  var counts = {}, order = [], seen = {}, uncounted = 0;

  var clsRows = getSettingsRows_(SETTINGS_TYPE.CLASS);
  for (var i = 0; i < clsRows.length; i++) {
    var k = safeStr(clsRows[i].key);
    if (k && !seen[k]) { seen[k] = true; order.push(k); counts[k] = blank(); }
  }

  var students = getAllStudentProfiles_();
  for (var s = 0; s < students.length; s++) {
    var p = students[s];
    if (!isActiveStudent_(p.status)) continue;          // present strength only
    var cls = safeStr(p.class);
    if (!cls) continue;
    if (!seen[cls]) { seen[cls] = true; order.push(cls); counts[cls] = blank(); }
    var cat = casteCategory_(p.caste, p.subCaste);      // SC/ST/BC/OC, or "" when not recorded
    var gb  = genderBucket_(p.gender);                  // "B"/"G", or "" when not recorded
    if (!cat) uncounted++;                              // caste missing -> reported, not forced into OC
    if (!cat || !gb) continue;                          // count ONLY real caste + gender
    counts[cls][cat][gb === "B" ? "b" : "g"]++;
  }

  var buckets = { "Pre-Primary": [], "Primary": [], "High School": [] };
  for (var o = 0; o < order.length; o++) {
    var cn = order[o], grp = classGroup_(cn), c = counts[cn];
    if (!buckets[grp]) buckets[grp] = [];
    buckets[grp].push({ cls: cn, SC: c.SC, ST: c.ST, BC: c.BC, OC: c.OC });
  }
  var groups = [
    { label: "Pre-Primary", rows: buckets["Pre-Primary"] },
    { label: "Primary",     rows: buckets["Primary"] },
    { label: "High School", rows: buckets["High School"] }
  ];

  var staff = { teaching: { male: 0, female: 0 }, nonTeaching: { male: 0, female: 0 } };
  var teachers = getAllTeacherProfiles_();
  for (var t = 0; t < teachers.length; t++) {
    var tt = teachers[t], stt = safeStr(tt.status);
    if (stt && stt.toLowerCase() !== "active") continue;
    var gb2 = genderBucket_(tt.gender);
    if (!gb2) continue;
    var bucket = isTeachingStaff_(tt.department, tt.designation) ? staff.teaching : staff.nonTeaching;
    bucket[gb2 === "B" ? "male" : "female"]++;
  }

  var data = {
    school: { name: school.name, place: school.address || "", logo: school.logo || "" },
    academicYear: ay, asOf: asOf, staff: staff, groups: groups, uncounted: uncounted
  };

  var built = buildCasteCensusHtml_(data, function (x) { return escHtml_(x); });
  return { success: true, css: built.css, html: built.body,
           fileName: "Caste_Census_" + asOf.replace(/-/g, "") };
}

/* ---- PUBLIC ENDPOINT: caste census as a downloadable PDF --------- */
function getCasteCensusPdf(token) {
  var r = getCasteCensusReport(token);
  if (r.error) return r;
  var fullHtml = "<!doctype html><html><head><meta charset='utf-8'><style>" + r.css +
                 "</style></head><body>" + r.html + "</body></html>";
  var bytes;
  try { bytes = htmlToPdfBlob_(fullHtml, r.fileName); }
  catch (e) { return { error: "Could not generate the PDF: " + e.message }; }
  return { success: true, fileName: r.fileName + ".pdf", pdfBase64: Utilities.base64Encode(bytes) };
}

/* ---- helpers for the Data Audit ---- */
function _censusCsvCell_(v) {
  v = String(v == null ? "" : v);
  if (/[",\n\r]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
  return v;
}

/* ---- PUBLIC ENDPOINT: student data audit (missing PEN/APAAR/etc) - */
function getDataAuditReport(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };

  var school = getSchoolInfo(token);
  var ay = (typeof getCurrentAcademicYear_ === "function") ? getCurrentAcademicYear_() : "";
  var tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  var asOf = Utilities.formatDate(new Date(), tz, "dd-MMM-yyyy");

  // class display order (so the missing list reads top-class-first)
  var clsRows = getSettingsRows_(SETTINGS_TYPE.CLASS), clsIdx = {};
  for (var i = 0; i < clsRows.length; i++) clsIdx[safeStr(clsRows[i].key)] = i;

  var total = 0, aA = 0, pP = 0, aP = 0, cC = 0, missing = [];
  var students = getAllStudentProfiles_();
  for (var s = 0; s < students.length; s++) {
    var p = students[s];
    if (!isActiveStudent_(p.status)) continue;
    total++;
    var hasA  = fieldPresent_(p.aadhar);
    var hasP  = fieldPresent_(p.pen);
    var hasAp = fieldPresent_(p.apaar);
    var hasC  = casteCategory_(p.caste, p.subCaste) !== "";
    if (hasA)  aA++;
    if (hasP)  pP++;
    if (hasAp) aP++;
    if (hasC)  cC++;
    if (!(hasA && hasP && hasAp && hasC)) {
      var ck = safeStr(p.class);
      var cls = ck + (safeStr(p.section) ? "-" + safeStr(p.section) : "");
      missing.push({ studentId: safeStr(p.studentId), name: safeStr(p.name), cls: cls,
        aadhar: hasA, pen: hasP, apaar: hasAp, caste: hasC,
        mobile: safeStr(p.fatherMobile), _ck: ck });
    }
  }
  missing.sort(function (a, b) {
    var ai = (clsIdx[a._ck] == null ? 999 : clsIdx[a._ck]);
    var bi = (clsIdx[b._ck] == null ? 999 : clsIdx[b._ck]);
    if (ai !== bi) return ai - bi;
    return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0);
  });

  var data = {
    school: { name: school.name, place: school.address || "", logo: school.logo || "" },
    academicYear: ay, asOf: asOf,
    summary: { total: total, aadhar: aA, pen: pP, apaar: aP, caste: cC },
    missing: missing
  };
  var built = buildDataAuditHtml_(data, function (x) { return escHtml_(x); });

  // CSV (worklist of students to update)
  var lines = [];
  lines.push(["Student ID", "Name", "Class", "Father Mobile",
              "Aadhar", "PEN", "APAAR", "Caste", "Missing Fields"].join(","));
  for (var m = 0; m < missing.length; m++) {
    var x = missing[m], miss = [];
    if (!x.aadhar) miss.push("Aadhar");
    if (!x.pen)    miss.push("PEN");
    if (!x.apaar)  miss.push("APAAR");
    if (!x.caste)  miss.push("Caste");
    lines.push([
      _censusCsvCell_(x.studentId), _censusCsvCell_(x.name), _censusCsvCell_(x.cls),
      _censusCsvCell_(x.mobile),
      x.aadhar ? "Yes" : "No", x.pen ? "Yes" : "No", x.apaar ? "Yes" : "No", x.caste ? "Yes" : "No",
      _censusCsvCell_(miss.join("; "))
    ].join(","));
  }

  return { success: true, css: built.css, html: built.body, csv: lines.join("\r\n"),
           fileName: "Data_Audit_" + asOf.replace(/-/g, ""), summary: data.summary };
}

/* ---- PUBLIC ENDPOINT: data audit as a downloadable PDF ----------- */
function getDataAuditPdf(token) {
  var r = getDataAuditReport(token);
  if (r.error) return r;
  var fullHtml = "<!doctype html><html><head><meta charset='utf-8'><style>" + r.css +
                 "</style></head><body>" + r.html + "</body></html>";
  var bytes;
  try { bytes = htmlToPdfBlob_(fullHtml, r.fileName); }
  catch (e) { return { error: "Could not generate the PDF: " + e.message }; }
  return { success: true, fileName: r.fileName + ".pdf", pdfBase64: Utilities.base64Encode(bytes) };
}


/* ==================================================================
   WHATSAPP NOTIFICATIONS  —  add-on for Code.gs
   Free click-to-chat (wa.me) approach: the app builds a WhatsApp link
   with the message pre-filled for each parent; the sender taps Send in
   WhatsApp. Templates are editable any time, and every notification is
   written to the "WhatsApp Log" sheet for audit.
   Reuses: validateSession, safeStr, safeNum, getAllStudentProfiles_,
   isActiveStudent_, fieldPresent_, getSettingsRows_, SETTINGS_TYPE,
   getSchoolInfo, PERM_MSG.
   ================================================================== */

var WA_TPL_SHEET = "WhatsApp Templates";
var WA_LOG_SHEET = "WhatsApp Log";
var WA_SENT_SHEET = "WhatsApp Sent";

function ensureWaSheets_() {
  var ss = SpreadsheetApp.getActive();
  var t = ss.getSheetByName(WA_TPL_SHEET);
  if (!t) { t = ss.insertSheet(WA_TPL_SHEET); t.appendRow(["Name", "Category", "Body", "Active", "UpdatedAt", "UpdatedBy"]); }
  var l = ss.getSheetByName(WA_LOG_SHEET);
  if (!l) { l = ss.insertSheet(WA_LOG_SHEET); l.appendRow(["Timestamp", "Sent By", "Type", "Scope", "Template", "Recipients", "Numbers", "Message", "Count"]); }
  var s = ss.getSheetByName(WA_SENT_SHEET);
  if (!s) { s = ss.insertSheet(WA_SENT_SHEET); s.appendRow(["RefDate", "StudentId", "Name", "TemplateKey", "SentBy", "Timestamp"]); }
  return { t: t, l: l, s: s };
}

// Sent-tracking column map for "WhatsApp Sent".
var WSENT_COL = { REFDATE: 0, STUDENT_ID: 1, NAME: 2, TEMPLATE_KEY: 3, SENT_BY: 4, TIMESTAMP: 5 };

// Returns { studentId: true } of everyone already marked sent for this
// templateKey on this reference date — used to auto-hide them from the list.
function getWhatsappSentSet_(templateKey, dateStr) {
  if (!templateKey) return {};
  var sh = ensureWaSheets_().s;
  var data = sh.getDataRange().getValues();
  var set = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (safeStr(r[WSENT_COL.TEMPLATE_KEY]) !== templateKey) continue;
    if (dateStr && safeStr(r[WSENT_COL.REFDATE]) !== dateStr) continue;
    set[safeStr(r[WSENT_COL.STUDENT_ID])] = true;
  }
  return set;
}

// Marks one student as notified for (templateKey, refDate). Idempotent.
function markWhatsappSent(token, studentId, studentName, templateKey, dateStr) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (!studentId || !templateKey) return { error: "Missing student or template." };
  var refDate = dateStr || _ymd_(new Date());
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = ensureWaSheets_().s;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (safeStr(data[i][WSENT_COL.STUDENT_ID]) === safeStr(studentId) &&
          safeStr(data[i][WSENT_COL.TEMPLATE_KEY]) === templateKey &&
          safeStr(data[i][WSENT_COL.REFDATE]) === refDate) {
        return { success: true, already: true }; // already recorded
      }
    }
    sh.appendRow([refDate, safeStr(studentId), safeStr(studentName), templateKey, sess.name || sess.user, _waNow_()]);
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

// Removes a sent record (undo) so the student reappears in the list.
function unmarkWhatsappSent(token, studentId, templateKey, dateStr) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var refDate = dateStr || _ymd_(new Date());
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = ensureWaSheets_().s;
    var data = sh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (safeStr(data[i][WSENT_COL.STUDENT_ID]) === safeStr(studentId) &&
          safeStr(data[i][WSENT_COL.TEMPLATE_KEY]) === templateKey &&
          safeStr(data[i][WSENT_COL.REFDATE]) === refDate) {
        sh.deleteRow(i + 1);
      }
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

function _waBool_(v) { return v === "" || v === true || String(v).toLowerCase() === "true"; }
function _waNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy HH:mm"); }

/* ---------- TEMPLATES (create / read / update / delete) ---------- */
function getWhatsappTemplates(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var sh = ensureWaSheets_().t;
  var data = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!safeStr(r[0])) continue;
    out.push({ rowIndex: i + 1, name: safeStr(r[0]), category: safeStr(r[1]), body: safeStr(r[2]), active: _waBool_(r[3]) });
  }
  // first run: seed a few useful school templates so the tab isn't empty
  if (!out.length) {
    var seed = [
      ["Fee Reminder", "Fees", "Dear {fatherName}, this is a gentle reminder that the school fee for {studentName} (Class {class}) is pending. Kindly clear it at the earliest. - {schoolName}"],
      ["Holiday Notice", "General", "Dear Parent, {schoolName} will remain closed on {date} on account of a holiday. Regular classes resume the next working day."],
      ["PTM Invitation", "Meeting", "Dear {fatherName}, you are invited to the Parent-Teacher Meeting for {studentName} (Class {class}) on {date}. Your presence is requested. - {schoolName}"],
      ["Absent Notice", "Attendance", "Dear {fatherName}, your ward {studentName} (Class {class}) was absent on {date}. Please inform the school of the reason. - {schoolName}"],
      ["Exam Schedule", "Exams", "Dear Parent, the examination schedule for Class {class} has been announced. Please ensure {studentName} is well prepared. - {schoolName}"],
      ["Report Card", "ReportCard", "Dear {fatherName}, please find attached the report card of {studentName} (Class {class}). - {schoolName}"]
    ];
    seed.forEach(function (s) { sh.appendRow([s[0], s[1], s[2], true, _waNow_(), "system"]); });
    return getWhatsappTemplates(token);
  }
  // Backfill: schools whose templates sheet predates the report-card
  // feature won't have it (seeding only runs on an empty sheet), so add
  // it once if missing.
  var hasReportCard = out.some(function (t) { return String(t.category).toLowerCase() === "reportcard" || /report\s*card/i.test(t.name); });
  if (!hasReportCard) {
    sh.appendRow(["Report Card", "ReportCard", "Dear {fatherName}, please find attached the report card of {studentName} (Class {class}). - {schoolName}", true, _waNow_(), "system"]);
    return getWhatsappTemplates(token);
  }
  return { success: true, data: out };
}

function saveWhatsappTemplate(token, payload) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin" && sess.role !== "principal") return { error: PERM_MSG.MGMT_ONLY };
  payload = payload || {};
  var name = safeStr(payload.name), body = safeStr(payload.body);
  if (!name) return { error: "Template name is required." };
  if (!body) return { error: "Message body is required." };
  var sh = ensureWaSheets_().t;
  var active = payload.active === false ? false : true;
  var row = [name, safeStr(payload.category), body, active, _waNow_(), sess.name || sess.user];
  if (payload.rowIndex) sh.getRange(payload.rowIndex, 1, 1, 6).setValues([row]);
  else sh.appendRow(row);
  return { success: true };
}

function deleteWhatsappTemplate(token, rowIndex) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  if (sess.role !== "admin" && sess.role !== "principal") return { error: PERM_MSG.MGMT_ONLY };
  if (!rowIndex) return { error: "Invalid template row." };
  ensureWaSheets_().t.deleteRow(rowIndex);
  return { success: true };
}

/* ---------- RECIPIENTS (active students + parent mobile) ---------- */
function getWhatsappRecipients(token, filter, dateStr, templateKey) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  filter = safeStr(filter) || "all";
  var refDate = dateStr || _ymd_(new Date());

  // Pre-compute filter sets so we only build them when needed.
  var absentSet = null, dueMap = null;
  if (filter === "absent") {
    ensureDailyStudentAttendance_();
    absentSet = getAbsentStudentsForDate_(refDate);
  } else if (filter === "due") {
    dueMap = {};
    getOutstandingList_(token).forEach(function(o) { dueMap[safeStr(o.studentId)] = o.due; });
  }

  // Who has already been notified for this template + reference date.
  var sentSet = templateKey ? getWhatsappSentSet_(templateKey, refDate) : {};

  var students = getAllStudentProfiles_(), out = [], sent = [];
  for (var i = 0; i < students.length; i++) {
    var p = students[i];
    if (!isActiveStudent_(p.status)) continue;
    var sid = safeStr(p.studentId);
    if (filter === "absent" && !absentSet[sid]) continue;
    if (filter === "due" && !(dueMap[sid] > 0)) continue;
    var mobile = fieldPresent_(p.fatherMobile) ? safeStr(p.fatherMobile)
               : (fieldPresent_(p.motherMobile) ? safeStr(p.motherMobile) : "");
    var rec = {
      studentId: sid, name: safeStr(p.name),
      classKey: safeStr(p.class),
      cls: safeStr(p.class) + (safeStr(p.section) ? "-" + safeStr(p.section) : ""),
      mobile: mobile, fatherName: safeStr(p.fatherName),
      due: dueMap ? (dueMap[sid] || 0) : 0
    };
    if (sentSet[sid]) sent.push(rec); else out.push(rec);
  }
  var clsRows = getSettingsRows_(SETTINGS_TYPE.CLASS), idx = {};
  for (var c = 0; c < clsRows.length; c++) idx[safeStr(clsRows[c].key)] = c;
  var sorter = function (a, b) {
    var ai = idx[a.classKey] == null ? 999 : idx[a.classKey];
    var bi = idx[b.classKey] == null ? 999 : idx[b.classKey];
    if (ai !== bi) return ai - bi;
    return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0);
  };
  out.sort(sorter); sent.sort(sorter);
  var school = getSchoolInfo(token);
  return { success: true, data: out, sent: sent, schoolName: school.name || "", filter: filter, refDate: refDate };
}

/* ---------- AUDIT LOG (write + read) ---------- */
function logWhatsappNotification(token, payload) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  payload = payload || {};
  var recips = payload.recipients || []; // [{name, number}]
  if (!recips.length) return { error: "No recipients with a valid mobile number to record." };
  var names = recips.map(function (r) { return safeStr(r.name); }).join(", ");
  var numbers = recips.map(function (r) { return safeStr(r.number); }).join(", ");
  ensureWaSheets_().l.appendRow([
    _waNow_(), sess.name || sess.user, safeStr(payload.type), safeStr(payload.scope),
    safeStr(payload.templateName), names, numbers, safeStr(payload.message), recips.length
  ]);
  return { success: true, count: recips.length };
}

function getWhatsappLog(token, limit) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var sh = ensureWaSheets_().l;
  var data = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!safeStr(r[0])) continue;
    out.push({
      ts: safeStr(r[0]), sentBy: safeStr(r[1]), type: safeStr(r[2]), scope: safeStr(r[3]),
      template: safeStr(r[4]), recipients: safeStr(r[5]), numbers: safeStr(r[6]),
      message: safeStr(r[7]), count: safeNum(r[8])
    });
  }
  out.reverse(); // newest first
  if (limit && out.length > limit) out = out.slice(0, limit);
  return { success: true, data: out };
}
// ============================================================
//  BIRTHDAY GREETINGS — auto-generates customised messages
//  for students and staff whose birthday (month+day) matches
//  the requested date range (single day, week, or month).
//  DOB is stored in "dd-MMM-yyyy" format via formatDateSafe_;
//  age is calculated from the birth year vs today.
// ============================================================

// Parse a DOB string (dd-MMM-yyyy or yyyy-MM-dd or a Date object)
// and return { month (1-12), day, year } or null on failure.
function parseDobParts_(dob) {
  if (!dob) return null;
  if (Object.prototype.toString.call(dob) === '[object Date]') {
    if (isNaN(dob.getTime()) || dob.getFullYear() < 1920) return null;
    return { month: dob.getMonth() + 1, day: dob.getDate(), year: dob.getFullYear() };
  }
  var s = String(dob);
  // dd-MMM-yyyy  e.g. 15-Aug-2005
  var m1 = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (m1) {
    var MONTHS = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
    var mo = MONTHS[m1[2].toLowerCase()];
    return mo ? { month: mo, day: Number(m1[1]), year: Number(m1[3]) } : null;
  }
  // yyyy-MM-dd
  var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) return { month: Number(m2[2]), day: Number(m2[3]), year: Number(m2[1]) };
  // dd/MM/yyyy
  var m3 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m3) return { month: Number(m3[2]), day: Number(m3[1]), year: Number(m3[3]) };
  return null;
}

// Returns true if the (month, day) pair falls within the inclusive date range.
function birthdayInRange_(parts, fromObj, toObj) {
  if (!parts) return false;
  // Compare only month+day against each calendar date in the range
  var d = new Date(fromObj.getFullYear(), fromObj.getMonth(), fromObj.getDate());
  var limit = new Date(toObj.getFullYear(), toObj.getMonth(), toObj.getDate());
  var guard = 0;
  while (d.getTime() <= limit.getTime() && guard < 400) {
    if ((d.getMonth() + 1) === parts.month && d.getDate() === parts.day) return true;
    d.setDate(d.getDate() + 1);
    guard++;
  }
  return false;
}

function calcAge_(birthYear, referenceDate) {
  if (!birthYear || birthYear < 1920) return '';
  var age = referenceDate.getFullYear() - birthYear;
  if (referenceDate.getMonth() + 1 < 0) age--; // conservative
  return String(age);
}

// Main endpoint: returns birthday matches in a date range.
// mode: 'day' (single date), 'week' (next 7 days), 'month' (full month)
// filterType: 'students', 'staff', or 'all'
// ── Birthday AI Greeting Generator ──────────────────────────────────────
// Calls the Claude API (claude-haiku-4-5-20251001) via UrlFetchApp to
// produce a unique, personalized birthday message for each person.
// The API key is stored in Script Properties under "ANTHROPIC_API_KEY"
// so it never appears in code. Set it once via: Project Settings →
// Script Properties → Add property: ANTHROPIC_API_KEY = sk-ant-...

function setAnthropicApiKey(token, apiKey) {
  var sess = validateSession(token);
  if (!sess) return { error: 'Session expired. Please log in again.' };
  var permErr = requireManagement_(sess); if (permErr) return permErr;
  if (!apiKey || !apiKey.trim()) return { error: 'Enter a valid API key.' };
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', apiKey.trim());
  return { success: true, message: 'API key saved.' };
}

function getAnthropicApiKeyStatus(token) {
  if (!validateSession(token)) return { error: 'Session expired.' };
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') || '';
  return { success: true, configured: !!key, preview: key ? ('sk-ant-...' + key.slice(-4)) : '' };
}

function generateBirthdayGreeting(token, name, age, classLabel, type, schoolName) {
  if (!validateSession(token)) return { error: 'Session expired. Please log in again.' };
  if (!name) return { error: 'Name is required.' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return {
    error: 'Anthropic API key not configured. Go to Settings → School Info → set the API Key, or add ANTHROPIC_API_KEY in Apps Script Project Settings → Script Properties.'
  };

  var isStaff = (type === 'staff');
  var ageStr  = age && age !== '?' ? age + ' years old' : '';

  var prompt =
    'Write a warm, genuine birthday greeting message for a ' +
    (isStaff ? 'school staff member' : 'school student') + '.\n\n' +
    'Details:\n' +
    '- Name: ' + name + '\n' +
    (ageStr ? '- Age: ' + ageStr + '\n' : '') +
    (isStaff ? '- Designation: ' + classLabel + '\n' : '- Class: ' + classLabel + '\n') +
    '- School: ' + schoolName + '\n\n' +
    'Guidelines:\n' +
    '- Start with a birthday emoji and the person\'s name\n' +
    '- Keep it warm and personal — 3 to 4 sentences\n' +
    (isStaff
      ? '- Appreciate their hard work and dedication to the students\n'
      : '- Encourage them in their studies and wish them bright success\n') +
    '- Sign off warmly from ' + schoolName + '\n' +
    '- Use 2-3 relevant emojis naturally within the text\n' +
    '- Return ONLY the message, no preamble or explanation.';

  var payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 300,
    messages: [{ role: 'user', content: prompt }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var res  = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    var body = JSON.parse(res.getContentText());
    if (body.error) return { error: 'AI error: ' + (body.error.message || JSON.stringify(body.error)) };
    var text = body.content && body.content[0] && body.content[0].text;
    if (!text) return { error: 'AI returned an empty response. Please try again.' };
    return { success: true, message: text.trim() };
  } catch (e) {
    return { error: 'Network error calling Claude API: ' + e.message };
  }
}

function getBirthdays(token, dateStr, mode, filterType) {
  if (!validateSession(token)) return { error: 'Session expired. Please log in again.' };
  var school = getSchoolInfo(token);
  var schoolName = (school && school.name) ? school.name : SCHOOL_NAME;

  var base  = dateStr ? parsePaymentDate_(dateStr) : new Date();
  var from  = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  var to;
  if (mode === 'week')       { to = new Date(from); to.setDate(to.getDate() + 6); }
  else if (mode === 'month') { to = new Date(from.getFullYear(), from.getMonth() + 1, 0); from = new Date(from.getFullYear(), from.getMonth(), 1); }
  else                       { to = new Date(from); }

  var results = [];

  if (filterType === 'students' || filterType === 'all') {
    var students = getAllStudentProfiles_().filter(function(p) { return p.status !== 'Removed'; });
    students.forEach(function(p) {
      var parts = parseDobParts_(p.dob);
      if (!birthdayInRange_(parts, from, to)) return;
      // Find the actual matching date(s) in range
      var matchDate = new Date(from.getFullYear(), (parts.month - 1), parts.day);
      // Could be next year if already passed this year in the range
      if (matchDate.getTime() < from.getTime()) matchDate.setFullYear(matchDate.getFullYear() + 1);
      results.push({
        type: 'student',
        id: p.studentId,
        name: p.name,
        classLabel: (p.class || '') + (p.section ? ' – ' + p.section : ''),
        dob: p.dob,
        age: parts.year ? String(matchDate.getFullYear() - parts.year) : '',
        phone: p.fatherMobile || p.motherMobile || '',
        photo: p.photo || '',
        matchDate: _ymd_(matchDate)
      });
    });
  }

  if (filterType === 'staff' || filterType === 'all') {
    var staff = getAllTeacherProfiles_().filter(function(t) { return t.status === 'Active'; });
    staff.forEach(function(t) {
      var parts = parseDobParts_(t.dob);
      if (!birthdayInRange_(parts, from, to)) return;
      var matchDate = new Date(from.getFullYear(), (parts.month - 1), parts.day);
      if (matchDate.getTime() < from.getTime()) matchDate.setFullYear(matchDate.getFullYear() + 1);
      results.push({
        type: 'staff',
        id: t.teacherId,
        name: t.name,
        classLabel: t.designation || 'Staff',
        dob: t.dob,
        age: parts.year ? String(matchDate.getFullYear() - parts.year) : '',
        phone: t.contactNo || '',
        photo: t.photo || '',
        matchDate: _ymd_(matchDate)
      });
    });
  }

  // Sort by matching date, then name
  results.sort(function(a, b) {
    if (a.matchDate < b.matchDate) return -1;
    if (a.matchDate > b.matchDate) return 1;
    return a.name < b.name ? -1 : 1;
  });

  return {
    success: true, schoolName: schoolName,
    fromDate: _ymd_(from), toDate: _ymd_(to),
    mode: mode, filterType: filterType,
    count: results.length, data: results
  };
}
// Minimal inline birthday lookup that doesn't require a session token —
// mirrors getBirthdays()'s single-day logic, used by both get_info and
// the unattended daily digest trigger.
function _agentBirthdaysOn_(dateStr) {
  var base = dateStr ? parsePaymentDate_(dateStr) : new Date();
  var from = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  var names = [];
  getAllStudentProfiles_().filter(function(p) { return p.status !== 'Removed'; }).forEach(function(p) {
    var parts = parseDobParts_(p.dob);
    if (birthdayInRange_(parts, from, from)) names.push(p.name);
  });
  getAllTeacherProfiles_().filter(function(t) { return t.status === 'Active'; }).forEach(function(t) {
    var parts = parseDobParts_(t.dob);
    if (birthdayInRange_(parts, from, from)) names.push(t.name);
  });
  return names;
}

// ── DAILY DIGEST — unattended email summary ─────────────────
// True automated WhatsApp sending isn't possible without a paid
// WhatsApp Business API provider, so this stays inside what's
// actually reliable: a scheduled trigger that emails an admin a
// morning summary (absentees, fee dues, birthdays). Free to run —
// it only reads data and sends one email; no external AI API is
// called. Configured from Settings → School Identity.
function getAgentDigestSettings(token) {
  if (!validateSession(token)) return { error: "Session expired. Please log in again." };
  var props = PropertiesService.getScriptProperties();
  return {
    success: true,
    enabled: props.getProperty('AGENT_DIGEST_ENABLED') === 'true',
    email: props.getProperty('AGENT_DIGEST_EMAIL') || '',
    hour: Number(props.getProperty('AGENT_DIGEST_HOUR') || 7)
  };
}

function setAgentDigestSettings(token, enabled, email, hour) {
  var sess = validateSession(token);
  if (!sess) return { error: "Session expired. Please log in again." };
  var permErr = requireAdmin_(sess); if (permErr) return permErr;
  hour = Math.max(0, Math.min(23, Number(hour) || 7));

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runAgentDailyDigest_') ScriptApp.deleteTrigger(t);
  });

  var props = PropertiesService.getScriptProperties();
  if (enabled) {
    if (!email || !email.trim()) return { error: "Enter an email address to receive the daily digest." };
    ScriptApp.newTrigger('runAgentDailyDigest_').timeBased().everyDays(1).atHour(hour).nearMinute(0).create();
    props.setProperty('AGENT_DIGEST_ENABLED', 'true');
    props.setProperty('AGENT_DIGEST_EMAIL', email.trim());
    props.setProperty('AGENT_DIGEST_HOUR', String(hour));
  } else {
    props.setProperty('AGENT_DIGEST_ENABLED', 'false');
  }
  return { success: true };
}

// Runs unattended via the time-driven trigger — no session/token
// available, so it only reads data and emails a summary. It never
// touches WhatsApp on its own.
function runAgentDailyDigest_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('AGENT_DIGEST_ENABLED') !== 'true') return;
  var email = props.getProperty('AGENT_DIGEST_EMAIL');
  if (!email) return;

  var todayStr = _ymd_(new Date());
  var school = getSchoolInfo('') || {};
  var schoolName = school.name || SCHOOL_NAME;
  var lines = [];
  lines.push('Good morning \u2014 here is the ' + schoolName + ' daily digest for ' + todayStr + '.', '');

  try {
    ensureDailyStudentAttendance_();
    var absentCount = Object.keys(getAbsentStudentsForDate_(todayStr)).length;
    lines.push('\uD83D\uDCCB Attendance: ' + (absentCount ? (absentCount + " student(s) already marked absent today.") : "Today's attendance isn't marked yet."));
  } catch (e) { lines.push('\uD83D\uDCCB Attendance: could not be checked.'); }

  try {
    var due = getOutstandingList_('');
    var totalDue = 0; due.forEach(function(s) { totalDue += s.due; });
    lines.push('\uD83D\uDCB0 Fees: ' + due.length + ' student(s) with pending dues, totalling \u20B9' + totalDue.toLocaleString('en-IN') + '.');
  } catch (e) { lines.push('\uD83D\uDCB0 Fees: could not be checked.'); }

  try {
    var bdNames = _agentBirthdaysOn_(todayStr);
    lines.push('\uD83C\uDF82 Birthdays today: ' + bdNames.length + (bdNames.length ? (' \u2014 ' + bdNames.join(', ')) : ''));
  } catch (e) {}

  lines.push('', 'Open WhatsApp Notifications in the app to queue absent notices or fee reminders for one-tap sending \u2014 nothing sends automatically.');

  try {
    MailApp.sendEmail({ to: email, subject: schoolName + ' \u2014 Daily Digest \u2014 ' + todayStr, body: lines.join('\n') });
  } catch (e) {}
}
