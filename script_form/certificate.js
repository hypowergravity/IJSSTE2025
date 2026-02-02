/**
 * =====================================================
 * IJSSTE Hokkaido Chapter Symposium 2025 — Certificate Distribution + Thank You Email Script
 * =====================================================
 */

/**
 * -------------------------
 * CONFIG (EDIT THESE ONLY)
 * -------------------------
 */
const EMAIL_CONFIG = {
  senderName: "IJSSTE Hokkaido Chapter Symposium Organizing Team",
  subject: "Certificate - IJSSTE Hokkaido Chapter Symposium 2025",

  headerTitle: "Thank you for attending IJSSTE Hokkaido Chapter Symposium 2025",
  headerSubtitle: "We sincerely appreciate your participation and support.",

  defaultGreetingName: "Participant",

  thankYouParagraph:
    "Thank you very much for attending the IJSSTE Hokkaido Chapter Symposium 2025 and making the event a great success. We hope such events foster greater academic collaboration between India and Japan.",

  certificateNote:
    "Please find your certificate attached to this email as recognition of your valuable participation in the symposium.",

  attachmentNote: "Please find your certificate attached for your kind perusal.",

  closingLine: "Warm regards,",
  signatureLine1: "IJSSTE Hokkaido Chapter",
  signatureLine2: "Organizing Committee"
};

// Certificate folder in Google Drive (folder URL or folder ID)
const CERTIFICATE_FOLDER = "https://drive.google.com/drive/folders/";

// Additional attachments (Drive URLs or file IDs) - optional
const ATTACHMENT_FILES = ""
// [
//   "https://drive.google.com/file/d/usp=sharing"
// ];

/**
 * Paste your Google Form EDIT URL (recommended) or just the Form ID.
 * Example:
 *   https://docs.google.com/forms/d/<FORM_ID>/edit
 */
const FORM_URL_OR_ID ="";

// Question matching (used for parsing form responses)
const FIRST_NAME_REGEX = /First Name/i;
const EMAIL_REGEX = /Email Address/i;


// Behavior
// const DEDUPE_BY_EMAIL = true;
// const SEND_ONLY_TO_LATEST = true;
// const SLEEP_MS = 500;

/**
 * -------------------------
 * MAIN SENDER WITH CERTIFICATES
 * -------------------------
 * Sends certificates and thank you emails to all registrants
 */
function sendCertificatesAndThankYouEmail() {
  try {
    const form = getForm_();
    const responses = form.getResponses();
    Logger.log("IJSSTE: Total responses: " + responses.length);

    const recipients = buildRecipientList_(responses, {
      dedupeByEmail: DEDUPE_BY_EMAIL,
      keepLatest: SEND_ONLY_TO_LATEST
    });
    Logger.log("IJSSTE: Valid recipients: " + recipients.length);

    // Load certificate files from Drive folder
    const certificateFiles = loadCertificateFiles_();
    Logger.log("IJSSTE: Loaded " + Object.keys(certificateFiles).length + " certificates");

    // Load additional attachments
    const attachments = getDriveBlobs_(ATTACHMENT_FILES);

    let sent = 0, notFound = 0;
    for (let i = 0; i < recipients.length; i++) {
      const r = recipients[i];
      
      // Find matching certificate
      const certificate = findCertificateForRecipient_(r, certificateFiles);
      
      if (!certificate) {
        Logger.log("IJSSTE: Certificate not found for: " + r.firstName + " " + r.lastName + " (" + r.email + ")");
        notFound++;
        continue;
      }

      // Prepare email attachments (certificate + additional files)
      const emailAttachments = attachments.slice(); // Copy array
      emailAttachments.push(certificate);

      const htmlBody = buildCertificateAndThankYouHTML_(r.greetingName, EMAIL_CONFIG);

      try {
        MailApp.sendEmail({
          to: r.email,
          subject: EMAIL_CONFIG.subject,
          htmlBody: htmlBody,
          attachments: emailAttachments,
          name: EMAIL_CONFIG.senderName
        });

        sent++;
        Logger.log("IJSSTE: Sent " + sent + "/" + recipients.length + " to: " + r.email + " (Certificate: " + certificate.getName() + ")");
        Utilities.sleep(SLEEP_MS);

      } catch (mailError) {
        Logger.log("IJSSTE: FAILED to send to " + r.email + " | " + mailError);
      }
    }

    Logger.log("IJSSTE: Done. Sent: " + sent + " / " + recipients.length + " (Not found: " + notFound + ")");

  } catch (error) {
    Logger.log("IJSSTE: Error in sendCertificatesAndThankYouEmail(): " + error);
    throw error;
  }
}


/**
 * -------------------------
 * TEST (SEND ONE EMAIL WITH CERTIFICATE)
 * -------------------------
 * Reads from form document but sends test email to you
 */
function testCertificateToMyself() {
  const TEST_TO = "ijsste.hokkaido.symposium.2025@gmail.com";

  try {
    // Get form and responses like real function
    const form = getForm_();
    const responses = form.getResponses();
    Logger.log("IJSSTE: Test - Total responses: " + responses.length);

    if (responses.length === 0) {
      throw new Error("No form responses found for testing");
    }

    // Use the first response for testing
    const testResponse = responses[0];
    const parsed = parseResponseForNameAndEmail_(testResponse);
    
    if (!isValidEmail_(parsed.email)) {
      throw new Error("First form response does not have valid email for testing");
    }

    // Load certificate files from Drive folder
    const certificateFiles = loadCertificateFiles_();
    Logger.log("IJSSTE: Test - Loaded " + Object.keys(certificateFiles).length + " certificates");

    // Try to find certificate for this participant
    let testCertificate = findCertificateForRecipient_({
      firstName: parsed.firstName,
      lastName: "", // Will be extracted if available
      email: parsed.email
    }, certificateFiles);

    // If no specific certificate found, use first available
    if (!testCertificate) {
      const allCertificates = Object.values(certificateFiles);
      if (allCertificates.length === 0) {
        throw new Error("No certificates found in Drive folder for testing");
      }
      testCertificate = allCertificates[0];
      Logger.log("IJSSTE: Test - No specific match, using first certificate: " + testCertificate.getName());
    } else {
      Logger.log("IJSSTE: Test - Found matching certificate: " + testCertificate.getName());
    }

    const attachments = getDriveBlobs_(ATTACHMENT_FILES);
    attachments.push(testCertificate);

    const greetingName = parsed.firstName || "Test Participant";
    const htmlBody = buildCertificateAndThankYouHTML_(greetingName, EMAIL_CONFIG);

    // Send to your email instead of participant's email
    MailApp.sendEmail({
      to: TEST_TO,
      subject: "TEST — " + EMAIL_CONFIG.subject + " (for " + parsed.firstName + ")",
      htmlBody: htmlBody,
      attachments: attachments,
      name: EMAIL_CONFIG.senderName
    });

    Logger.log("IJSSTE: Test email sent to: " + TEST_TO + " (certificate for: " + greetingName + ")");

  } catch (error) {
    Logger.log("IJSSTE: Test failed: " + error);
    throw error;
  }
}


/**
 * -------------------------
 * FORM ACCESS
 * -------------------------
 */
function getForm_() {
  if (FORM_URL_OR_ID && FORM_URL_OR_ID.trim()) {
    const formId = extractFormId_(FORM_URL_OR_ID);
    if (!formId) throw new Error("IJSSTE: Could not extract Form ID from: " + FORM_URL_OR_ID);
    return FormApp.openById(formId);
  }

  // Optional fallback if (and only if) this script is bound directly to the Form:
  // const active = FormApp.getActiveForm();
  // if (active) return active;

  throw new Error("IJSSTE: FORM_URL_OR_ID is empty. Paste the Form edit URL or Form ID.");
}

function extractFormId_(urlOrId) {
  const s = String(urlOrId || "").trim();
  if (!s) return null;

  // Handles:
  // /forms/d/<ID>/edit
  // /forms/d/e/<ID>/viewform
  let m = s.match(/\/forms\/d\/e\/([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];

  m = s.match(/\/forms\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];

  // Raw ID
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;

  return null;
}

/**
 * -------------------------
 * RECIPIENT PARSING
 * -------------------------
 */
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
          firstName: r.firstName,
          lastName: "",
          greetingName: r.firstName ? r.firstName : EMAIL_CONFIG.defaultGreetingName,
          timestamp: responses[i].getTimestamp()
        });
      }
    }
    return list;
  }

  const byEmail = {};
  for (let i = 0; i < responses.length; i++) {
    const resp = responses[i];
    const ts = resp.getTimestamp();
    const parsed = parseResponseForNameAndEmail_(resp);
    if (!isValidEmail_(parsed.email)) continue;

    const record = {
      email: parsed.email.trim(),
      firstName: parsed.firstName,
      lastName: "",
      greetingName: parsed.firstName ? parsed.firstName : EMAIL_CONFIG.defaultGreetingName,
      timestamp: ts
    };

    if (!byEmail[record.email]) {
      byEmail[record.email] = record;
    } else if (keepLatest && byEmail[record.email].timestamp < record.timestamp) {
      byEmail[record.email] = record;
    }
  }

  const list = Object.keys(byEmail).map(k => byEmail[k]);
  list.sort((a, b) => a.timestamp - b.timestamp);
  return list;
}


function parseResponseForNameAndEmail_(response) {
  let email = "";
  let firstName = "";

  // Prefer collected email if available (requires "Collect email addresses" enabled)
  try {
    const collected = response.getRespondentEmail();
    if (collected) email = String(collected).trim();
  } catch (e) {}

  const itemResponses = response.getItemResponses();
  for (let i = 0; i < itemResponses.length; i++) {
    const itemResponse = itemResponses[i];
    const question = itemResponse.getItem().getTitle();
    const answer = itemResponse.getResponse();

    if (!firstName && FIRST_NAME_REGEX.test(question)) {
      firstName = String(answer || "").trim();
    }
    if (!email && EMAIL_REGEX.test(question)) {
      email = String(answer || "").trim();
    }
  }

  return { firstName, email };
}

function isValidEmail_(email) {
  const e = String(email || "").trim();
  if (!e) return false;
  if (e.indexOf("@") < 1) return false;
  if (e.indexOf(" ") !== -1) return false;
  return true;
}

/**
 * -------------------------
 * CERTIFICATE MANAGEMENT
 * -------------------------
 */
function loadCertificateFiles_() {
  const certificateFiles = {};

  // Load all certificates from single folder
  if (CERTIFICATE_FOLDER && CERTIFICATE_FOLDER !== "PASTE_CERTIFICATES_DRIVE_FOLDER_URL_HERE") {
    const folderId = extractDriveFolderId_(CERTIFICATE_FOLDER);
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFilesByType(MimeType.JPEG);
      
      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        // Extract name from filename: "FirstName_LastName_poster_certificate.jpeg" or "FirstName_LastName_participation_certificate.jpeg"
        const nameKey = extractNameFromCertificateFilename_(fileName);
        if (nameKey) {
          certificateFiles[nameKey] = file.getBlob().setName(fileName);
        }
      }
    }
  }

  return certificateFiles;
}

function extractNameFromCertificateFilename_(filename) {
  // Expected formats: 
  // "FirstName_LastName_poster_certificate.jpeg"
  // "FirstName_LastName_participation_certificate.jpeg"
  // Extract "FirstName_LastName" part
  
  const match = filename.match(/^(.+)_(poster|participation)_certificate\./);
  if (match) {
    return match[1].toLowerCase(); // Return normalized name key
  }
  return null;
}

function findCertificateForRecipient_(recipient, certificateFiles) {
  // Create search key from recipient name
  const searchKey = createNameSearchKey_(recipient.firstName, recipient.lastName);
  
  // Try exact match first
  if (certificateFiles[searchKey]) {
    return certificateFiles[searchKey];
  }

  // Try fuzzy matching - look for partial matches
  const allKeys = Object.keys(certificateFiles);
  for (let i = 0; i < allKeys.length; i++) {
    const key = allKeys[i];
    if (nameMatches_(searchKey, key)) {
      Logger.log("IJSSTE: Fuzzy match found: " + searchKey + " -> " + key);
      return certificateFiles[key];
    }
  }

  Logger.log("IJSSTE: No certificate found for: " + searchKey);
  return null;
}

function createNameSearchKey_(firstName, lastName) {
  // Create normalized search key: "firstname_lastname"
  const first = String(firstName || "").trim().toLowerCase().replace(/\s+/g, "_");
  const last = String(lastName || "").trim().toLowerCase().replace(/\s+/g, "_");
  return first + "_" + last;
}

function nameMatches_(searchKey, fileKey) {
  // Simple fuzzy matching - check if the names are similar
  // Remove common variations and check
  const normalize = (str) => str.replace(/[._\s-]/g, "").toLowerCase();
  
  const normalizedSearch = normalize(searchKey);
  const normalizedFile = normalize(fileKey);
  
  // Check if they're the same when normalized
  if (normalizedSearch === normalizedFile) {
    return true;
  }
  
  // Check if search is contained in file (handles cases like "john_smith" vs "john_d_smith")
  if (normalizedFile.includes(normalizedSearch) || normalizedSearch.includes(normalizedFile)) {
    return true;
  }
  
  return false;
}

function extractDriveFolderId_(urlOrId) {
  const s = String(urlOrId || "").trim();
  if (!s) return null;
  
  // Handle folder URLs like: https://drive.google.com/drive/folders/FOLDER_ID
  let m = s.match(/\/folders\/([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];
  
  // Handle folder URLs with additional parameters
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
  if (m && m[1]) return m[1];
  
  // Raw folder ID
  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;
  
  return null;
}

/**
 * -------------------------
 * DRIVE ATTACHMENTS
 * -------------------------
 */
function getDriveBlobs_(urlOrIds) {
  const blobs = [];
  for (let i = 0; i < urlOrIds.length; i++) {
    const fileId = extractDriveFileId_(urlOrIds[i]);
    if (!fileId) throw new Error("IJSSTE: Could not extract Drive file ID from: " + urlOrIds[i]);

    const file = DriveApp.getFileById(fileId);
    blobs.push(file.getBlob().setName(file.getName()));
  }
  return blobs;
}

function extractDriveFileId_(urlOrId) {
  const s = String(urlOrId || "").trim();
  if (!s) return null;
  const m = s.match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

/**
 * -------------------------
 * EMAIL HTML
 * -------------------------
 */
function buildCertificateAndThankYouHTML_(greetingName, cfg) {
  const name = greetingName ? greetingName : cfg.defaultGreetingName;

  let html = "<html><body style='font-family: Arial, sans-serif; color: #333; line-height: 1.6;'>";
  html += "<div style='max-width: 640px; margin: 0 auto; padding: 20px;'>";

  html += "<h2 style='margin: 0; color: #2c5aa0; border-bottom: 2px solid #2c5aa0; padding-bottom: 10px;'>";
  html += escapeHtml_(cfg.headerTitle);
  html += "</h2>";

  if (cfg.headerSubtitle) {
    html += "<p style='margin: 10px 0 20px 0; color: #666; font-style: italic;'>";
    html += escapeHtml_(cfg.headerSubtitle);
    html += "</p>";
  }

  html += "<p>Dear " + escapeHtml_(name) + ",</p>";

  html += "<p>" + escapeHtml_(cfg.thankYouParagraph) + "</p>";
  
  if (cfg.certificateNote) {
    html += "<p>" + escapeHtml_(cfg.certificateNote) + "</p>";
  }

  // Add link to IJSSTE images
  html += "<p><strong>Note:</strong> Images from IJSSTE events can be found at: ";
  html += "<a href='' target='_blank' rel='noopener noreferrer'>";
  html += "IJSSTE Event Images";
  html += "</a></p>";

  if (cfg.attachmentNote) {
    html += "<p>" + escapeHtml_(cfg.attachmentNote) + "</p>";
  }

  html += "<hr style='border: none; border-top: 1px solid #ddd; margin: 28px 0 18px 0;'>";
  html += "<p style='margin: 0; color: #666;'>" + escapeHtml_(cfg.closingLine) + "</p>";
  html += "<p style='margin: 6px 0 0 0; color: #666; font-weight: bold;'>" + escapeHtml_(cfg.signatureLine1) + "</p>";
  if (cfg.signatureLine2) {
    html += "<p style='margin: 4px 0 0 0; color: #666;'>" + escapeHtml_(cfg.signatureLine2) + "</p>";
  }

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