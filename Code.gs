/**
 * ============================================================
 *  Help Me Grow — Feedback Form Backend
 *  Google Apps Script Web App → Google Sheets
 * ============================================================
 *
 *  This script receives POST requests containing JSON-encoded
 *  form answers, validates them, and appends a new row to the
 *  connected Google Sheet — one row per submission.
 *
 *  DEPLOYMENT: See the "Deployment Guide" provided separately.
 * ============================================================
 */

/** -----------------------------------------------------------
 *  CONFIG
 *  ----------------------------------------------------------- */

// Name of the sheet (tab) inside your spreadsheet where responses go.
const SHEET_NAME = "Responses";

// Ordered list of every question ID expected in the payload.
// This drives both the header row AND the row-writing order,
// so the sheet columns always line up correctly even if the
// incoming JSON key order varies.
const QUESTION_IDS = [
  "q1","q2","q3","q4","q5","q6","q7","q8","q9","q10",
  "q11","q12","q13","q14","q15","q16","q17","q18","q19","q20",
  "q21","q22","q23","q24","q25"
];

// Which questions are checkbox/multi-select (their answers arrive
// as arrays and should be flattened to comma-separated text).
const MULTI_SELECT_IDS = new Set([
  "q1","q2","q5","q6","q7","q8","q9","q10","q11",
  "q12","q13","q14","q15","q16","q17",
  "q18","q19","q20","q21"
]);

/** -----------------------------------------------------------
 *  ENTRY POINT — handles POST requests from the frontend
 *  ----------------------------------------------------------- */
function doPost(e) {
  try {
    // --- 1. Validate that a request body was actually sent ---
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse(false, "No data received in request.");
    }

    // --- 2. Parse the incoming JSON safely ---
    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse(false, "Malformed JSON payload.");
    }

    const answers = payload.answers;
    if (!answers || typeof answers !== "object") {
      return jsonResponse(false, "Missing or invalid 'answers' object.");
    }

    // --- 3. Validate required questions are present and non-empty ---
    const validationError = validateAnswers(answers);
    if (validationError) {
      return jsonResponse(false, validationError);
    }

    // --- 4. Build the row in the correct column order ---
    const row = buildRow(answers);

    // --- 5. Append the row to the sheet ---
    const sheet = getOrCreateSheet();
    sheet.appendRow(row);

    // --- 6. Respond with success ---
    return jsonResponse(true, "Response saved successfully.");

  } catch (err) {
    // Catch-all: never let the script throw an unhandled error back
    // to the client — always return a clean JSON error instead.
    return jsonResponse(false, "Server error: " + err.message);
  }
}

/** -----------------------------------------------------------
 *  Optional: lets you sanity-check the deployed URL in a browser
 *  (GET requests aren't used for real submissions).
 *  ----------------------------------------------------------- */
function doGet(e) {
  return jsonResponse(true, "Web App is live. Use POST to submit responses.");
}

/** -----------------------------------------------------------
 *  Validates that every required question has a real answer.
 *  Returns an error message string, or null if valid.
 *  ----------------------------------------------------------- */
function validateAnswers(answers) {
  for (const id of QUESTION_IDS) {
    const value = answers[id];
    const isMulti = MULTI_SELECT_IDS.has(id);

    if (isMulti) {
      // Multi-select answers must be a non-empty array.
      if (!Array.isArray(value) || value.length === 0) {
        return `Question ${id} is required and was left empty.`;
      }
    } else {
      // Single-select answers must be a non-empty string.
      if (typeof value !== "string" || value.trim() === "") {
        return `Question ${id} is required and was left empty.`;
      }
    }
  }
  return null; // all good
}

/** -----------------------------------------------------------
 *  Converts the answers object into a flat array matching the
 *  sheet's column order: [Timestamp, Q1, Q2, ... Q25]
 *  ----------------------------------------------------------- */
function buildRow(answers) {
  const row = [new Date()]; // Timestamp column, auto-generated server-side

  for (const id of QUESTION_IDS) {
    const value = answers[id];

    if (MULTI_SELECT_IDS.has(id)) {
      // Checkbox answers: join array into a comma-separated string.
      // Each option is trimmed to strip any stray whitespace.
      const cleaned = Array.isArray(value)
        ? value.map(v => String(v).trim()).filter(v => v !== "")
        : [];
      row.push(cleaned.join(", "));
    } else {
      // Single-select answers: trimmed plain string.
      row.push(typeof value === "string" ? value.trim() : "");
    }
  }

  return row;
}

/** -----------------------------------------------------------
 *  Gets the target sheet, creating it (with headers) if it
 *  doesn't exist yet. Keeps the script safe to run on a fresh
 *  spreadsheet without manual setup.
 *  ----------------------------------------------------------- */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ["Timestamp", ...QUESTION_IDS.map(id => id.toUpperCase())];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/** -----------------------------------------------------------
 *  Builds a consistent JSON response for the frontend to parse.
 *  Keeping this in one place avoids inconsistent response shapes.
 *  ----------------------------------------------------------- */
function jsonResponse(success, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success, message }))
    .setMimeType(ContentService.MimeType.JSON);
}
