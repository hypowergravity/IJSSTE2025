/**
 * =====================================================
 * IJSSTE Hokkaido Symposium 2025 — Reminder Email Script
 * Time-driven reminder to ALL registrants + PDF attachment
 * =====================================================
 */

/**
 * -------------------------
 * CONFIG (EDIT THESE)
 * -------------------------
 */

// Paste either the Drive URL or just the file ID
const ABSTRACT_PDF_URL_OR_ID =
  "https://drive.google.com/file/d/";

// If the script is NOT bound to the form, set FORM_ID.
// If bound to the form, leave as "" and it will use getActiveForm().
const FORM_ID = ""; // e.g. "1FAIpQLSe....."

// Question matching (only used if collected email is not available)
const FIRST_NAME_QUESTION_REGEX = /First Name/i;
const EMAIL_QUESTION_REGEX = /Email Address/i;

// Email metadata
const SENDER_NAME = "IJSSTE Hokkaido Symposium 2025";
const EMAIL_SUBJECT = "Reminder: IJSSTE Hokkaido Symposium 2025 - Tomorrow";

// Behavior
const DEDUPE_BY_EMAIL = true;     // only send once per email
const SLEEP_MS = 500;            // throttle
const SEND_ONLY_TO_LATEST = true; // if true, keep latest response per email when deduping

/**
 * -------------------------
 * MAIN REMINDER (TIME-DRIVEN)
 * -------------------------
 * Create a time trigger to run this function.
 */
function sendConferenceReminder() {
  try {
    const form = getForm_();
    const responses = form.getResponses();

    Logger.log("Total responses: " + responses.length);

    const pdfBlob = getAbstractPdfBlob_(ABSTRACT_PDF_URL_OR_ID);

    const recipients = buildRecipientList_(responses, {
      dedupeByEmail: DEDUPE_BY_EMAIL,
      keepLatest: SEND_ONLY_TO_LATEST
    });

    Logger.log("Valid recipients: " + recipients.length);

    let sent = 0;
    for (let i = 0; i < recipients.length; i++) {
      const r = recipients[i];

      const htmlBody = buildShortReminderHTML_(r.greetingName);

      try {
        MailApp.sendEmail({
          to: r.email,
          subject: EMAIL_SUBJECT,
          htmlBody: htmlBody,
          attachments: [pdfBlob],
          name: SENDER_NAME
        });

        sent++;
        Logger.log("Sent " + sent + "/" + recipients.length + " to: " + r.email);
        Utilities.sleep(SLEEP_MS);

      } catch (mailError) {
        Logger.log("FAILED to send to " + r.email + " | " + mailError);
      }
    }

    Logger.log("Done. Sent: " + sent + " / " + recipients.length);

  } catch (error) {
    Logger.log("Error in sendConferenceReminder(): " + error);
    throw error;
  }
}

/**
 * -------------------------
 * TEST (SEND ONE EMAIL ONLY)
 * -------------------------
 */
function testReminderToMyself() {
  const TEST_TO = "ijsste.hokkaido.symposium.2025@gmail.com";

  const pdfBlob = getAbstractPdfBlob_(ABSTRACT_PDF_URL_OR_ID);

  MailApp.sendEmail({
    to: TEST_TO,
    subject: "TEST — " + EMAIL_SUBJECT,
    htmlBody: buildShortReminderHTML_("Test"),
    attachments: [pdfBlob],
    name: SENDER_NAME
  });

  Logger.log("Test reminder sent to: " + TEST_TO);
}

/**
 * -------------------------
 * TRIGGER SETUP (TIME-DRIVEN)
 * -------------------------
 * Creates a daily time trigger.
 * If you want "tomorrow at 09:00" specifically, see notes below.
 */
function setupDailyReminderTrigger() {
  // Delete existing triggers for this handler (avoid duplicates)
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendConferenceReminder") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create a daily trigger at ~9am (Apps Script chooses a time within the hour)
  ScriptApp.newTrigger("sendConferenceReminder")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log("Daily reminder trigger created (runs around 09:00).");
}

/**
 * Optional: one-time trigger "tomorrow at 09:00".
 * Run this once; it will schedule one execution.
 */
function setupOneTimeReminderTomorrowAt9() {
  // Delete existing triggers for this handler (avoid duplicates)
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendConferenceReminder") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const tz = Session.getScriptTimeZone();
  const now = new Date();

  // Compute "tomorrow at 09:00" in script timezone
  const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  const runAt = new Date(
    Utilities.formatDate(tomorrow, tz, "yyyy/MM/dd") + " 09:00:00"
  );

  // If parsing above behaves oddly in your locale, replace with:
  // const runAt = new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate(), 9, 0, 0);

  ScriptApp.newTrigger("sendConferenceReminder")
    .timeBased()
    .at(runAt)
    .create();

  Logger.log("One-time reminder trigger created for: " + runAt.toString());
}

/**
 * -------------------------
 * FORM + RESPONSE PARSING
 * -------------------------
 */

function getForm_() {
  if (FORM_ID && FORM_ID.trim()) {
    return FormApp.openById(FORM_ID.trim());
  }
  const active = FormApp.getActiveForm();
  if (!active) {
    throw new Error("Active form not found. Set FORM_ID if this is a standalone script.");
  }
  return active;
}

function buildRecipientList_(responses, opts) {
  const dedupeByEmail = !!opts.dedupeByEmail;
  const keepLatest = !!opts.keepLatest;

  if (!dedupeByEmail) {
    const list = [];
    for (let i = 0; i < responses.length; i++) {
      const r = parseResponseForNameAndEmail_(responses[i]);
      if (isValidEmail_(r.email)) {
        list.push({
          email: r.email,
          greetingName: r.firstName ? r.firstName : "Participant",
          timestamp: responses[i].getTimestamp()
        });
      }
    }
    return list;
  }

  // Deduped map: email -> record
  const byEmail = {};

  for (let i = 0; i < responses.length; i++) {
    const resp = responses[i];
    const ts = resp.getTimestamp();
    const parsed = parseResponseForNameAndEmail_(resp);

    if (!isValidEmail_(parsed.email)) continue;

    const record = {
      email: parsed.email.trim(),
      greetingName: parsed.firstName ? parsed.firstName : "Participant",
      timestamp: ts
    };

    if (!byEmail[record.email]) {
      byEmail[record.email] = record;
    } else if (keepLatest && byEmail[record.email].timestamp < record.timestamp) {
      byEmail[record.email] = record;
    }
  }

  const list = Object.keys(byEmail).map(k => byEmail[k]);
  // Sort optional (oldest -> newest)
  list.sort((a, b) => a.timestamp - b.timestamp);
  return list;
}

function parseResponseForNameAndEmail_(response) {
  let email = "";
  let firstName = "";

  // Prefer collected email if available
  try {
    const collected = response.getRespondentEmail();
    if (collected) email = String(collected).trim();
  } catch (e) {
    // ignore
  }

  const itemResponses = response.getItemResponses();
  for (let i = 0; i < itemResponses.length; i++) {
    const itemResponse = itemResponses[i];
    const question = itemResponse.getItem().getTitle();
    const answer = itemResponse.getResponse();

    if (!firstName && FIRST_NAME_QUESTION_REGEX.test(question)) {
      firstName = String(answer || "").trim();
    }

    if (!email && EMAIL_QUESTION_REGEX.test(question)) {
      email = String(answer || "").trim();
    }
  }

  return { firstName, email };
}

/**
 * -------------------------
 * DRIVE ATTACHMENT
 * -------------------------
 */

function getAbstractPdfBlob_(urlOrId) {
  const fileId = extractDriveFileId_(urlOrId);
  if (!fileId) throw new Error("Could not extract Drive file ID from: " + urlOrId);

  const file = DriveApp.getFileById(fileId);
  return file.getBlob().setName(file.getName());
}

function extractDriveFileId_(urlOrId) {
  const s = String(urlOrId || "").trim();
  if (!s) return null;
  const m = s.match(/[-\w]{25,}/); // typical Drive ID length/pattern
  return m ? m[0] : null;
}

function isValidEmail_(email) {
  const e = String(email || "").trim();
  if (!e) return false;
  if (e.indexOf("@") < 1) return false;
  if (e.indexOf(" ") !== -1) return false;
  // light sanity check; avoid being overly strict
  if (e.indexOf(".") === -1) return true; // still allow internal domains if any
  return true;
}

/**
 * -------------------------
 * EMAIL HTML
 * -------------------------
 */

function buildShortReminderHTML_(greeting) {
  let html = "<html><body style='font-family: Arial, sans-serif; color: #333; line-height: 1.6;'>";
  html += "<div style='max-width: 600px; margin: 0 auto; padding: 20px;'>";

  html += "<h2 style='color: #2c5aa0; border-bottom: 2px solid #2c5aa0; padding-bottom: 10px; margin-bottom: 5px;'>IJSSTE Hokkaido Symposium 2025</h2>";
  html += "<p style='color: #666; font-style: italic; margin-top: 5px; margin-bottom: 25px;'>AI and Experiment: Shaping the Next Generation of Molecular and Materials Design</p>";

  html += "<p>Dear " + escapeHtml_(greeting) + ",</p>";

  html += "<p>We hope this message finds you well. This is a friendly reminder that the IJSSTE Hokkaido Symposium 2025 is taking place tomorrow. We look forward to seeing you.</p>";

  html += "<p>Attached to this email, you will find a PDF containing all the presentation abstracts for your kind perusal.</p>";

  html += "<div style='background-color: #fff3cd; padding: 20px; border-left: 4px solid #ffc107; margin: 25px 0; border-radius: 5px;'>";
  html += "<p style='margin: 0 0 12px 0;'><strong>For Kind Attention:</strong></p>";
  html += "<p style='margin: 0;'><strong>Please bring exact change for the registration fee if possible.</strong> We will have limited change available at the registration desk, so having the exact amount ready will help us process smoothly. Thank you for your effort; it is greatly appreciated and will help us start the symposium on time.</p>";
  html += "</div>";

  html += "<div style='background-color: #f8f9fa; padding: 20px; margin: 25px 0; border-radius: 5px;'>";
  html += "<h4 style='color: #2c5aa0; margin-top: 0;'>What You Need to Know:</h4>";
  html += "<p style='margin: 10px 0;'><strong>Registration Fee:</strong> Please bring the exact amount in cash if possible. The fee will be collected at the registration desk upon arrival.</p>";
  html += "<p style='margin: 10px 0;'><strong>What’s Included:</strong> Your registration includes abstract booklets, lunch, and complimentary tea.</p>";
  html += "</div>";

  html += "<div style='background-color: #d4edda; padding: 15px; margin: 25px 0; border-left: 4px solid #28a745; border-radius: 5px;'>";
  html += "<p style='margin: 0;'><strong>Tip:</strong> We recommend reviewing the attached abstracts and creating a personal schedule of the presentations you would most like to attend.</p>";
  html += "</div>";

  html += "<p style='margin-top: 25px;'>We are excited about tomorrow’s symposium and grateful that you will be joining us. We are honored to have you as part of this gathering.</p>";

  html += "<p style='margin-top: 20px; font-size: 16px;'><strong>See you tomorrow, " + escapeHtml_(greeting) + ".</strong></p>";

  html += "<hr style='border: none; border-top: 1px solid #ddd; margin: 30px 0 20px 0;'>";
  html += "<p style='color: #666; margin: 5px 0;'>Warm regards,</p>";
  html += "<p style='color: #666; margin: 5px 0; font-weight: bold;'>IJSSTE Hokkaido Symposium 2025</p>";
  html += "<p style='color: #666; margin: 5px 0;'>Organizing Committee</p>";

  html += "</div></body></html>";
  return html;
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}