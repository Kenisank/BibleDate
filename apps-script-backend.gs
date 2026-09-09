// Bible Date registration backend.
// Paste this whole file into the Apps Script editor attached to your Google Sheet
// (see SETUP.md for step-by-step instructions), then deploy it as a Web App.

// ---- EDIT THESE BEFORE DEPLOYING ----
const EVENT_NAME = "Bible Date — Beyond the Pulpit";
const EVENT_DATE = "Friday 25th September 2026, 3PM";
const EVENT_VENUE =
  "KSDPC ESTATE MAIGERO AFTER MAMAS BUKA BEFORE FREEDOM HOTEL, TYK SCHOOL, U/Maigero, Kaduna State";
const SITE_URL = "https://bibledate.vercel.app"; // e.g. https://your-site.vercel.app/bible-date-registration.html
const ADMIN_KEY = "biblearomdate2026"; // must match ADMIN_PASSCODE in the HTML file, exactly
// --------------------------------------

const WA_NUMBER = "2348034503111"; // Crave & Crumb WhatsApp for event orders
const SUPPORT_PHONE = "0803 875 8275";

const HEADERS = [
  "Submitted At",
  "ID",
  "Name",
  "Phone",
  "Email",
  "Address",
  "Question",
  "Food",
  "Total",
  "Verified",
  "Access Code",
  "Purchased",
];

// Column indexes (1-based) into the sheet — mirror HEADERS order.
const COL = {
  SUBMITTED: 1,
  ID: 2,
  NAME: 3,
  PHONE: 4,
  EMAIL: 5,
  ADDRESS: 6,
  QUESTION: 7,
  FOOD: 8,
  TOTAL: 9,
  VERIFIED: 10,
  CODE: 11,
  PURCHASED: 12,
};

function esc(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function findRowById(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][COL.ID - 1] === id) return i + 1; // sheet rows are 1-indexed
  }
  return -1;
}

function findRowByCode(sheet, code) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (
      values[i][COL.CODE - 1] &&
      String(values[i][COL.CODE - 1]).toUpperCase() ===
        String(code).toUpperCase()
    )
      return i + 1;
  }
  return -1;
}

function generateAccessCode() {
  return Math.random().toString(36).slice(2, 8).toUpperCase();
}

function rowCell(sheet, rowNum, colNum) {
  try {
    return sheet.getRange(rowNum, colNum).getValue();
  } catch (err) {
    return "";
  }
}

function setCell(sheet, rowNum, colNum, value) {
  sheet.getRange(rowNum, colNum).setValue(value);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}

// Builds and sends the "registration received" email. Used on submit.
function sendRegistrationReceivedEmail(email, name) {
  const plainBody =
    `Hi ${name},\n\nThanks for registering for ${EVENT_NAME}.\n\n` +
    `Your registration is currently pending verification. Once verified, we'll send you a follow-up email with your access code and the event location, date and time.\n\n` +
    `For questions in the meantime, call ${SUPPORT_PHONE}.\n\nSee you soon.`;

  const htmlBody =
    `<div style="font-family:Arial,Helvetica,sans-serif;color:#111111;line-height:1.6;max-width:480px;">` +
    `<p>Hi ${esc(name)},</p>` +
    `<p>Thanks for registering for <strong>${esc(EVENT_NAME)}</strong>.</p>` +
    `<p>Your registration is currently <strong>pending verification</strong>. Once verified, we'll send you a follow-up email with your access code and the event location, date and time.</p>` +
    `<p>For questions in the meantime, call <strong>${esc(SUPPORT_PHONE)}</strong>.</p>` +
    `<p>See you soon.</p>` +
    `</div>`;

  MailApp.sendEmail({
    to: email,
    subject: `${EVENT_NAME} — Registration received`,
    body: plainBody,
    htmlBody: htmlBody,
  });
}

// Builds and sends the "you're verified" email with access code + deep link.
function sendVerifiedEmail(email, name, code) {
  const portalLink = `${SITE_URL}?view=mine&code=${code}`;

  const plainBody =
    `Hi ${name},\n\nYou're verified for ${EVENT_NAME}.\n\n` +
    `Access code: ${code}\nDate & time: ${EVENT_DATE}\nVenue: ${EVENT_VENUE}\n\n` +
    `Please keep your access code handy for entry.\n\n` +
    `For questions, call ${SUPPORT_PHONE}.\n\nSee you there.\n\n` +
    `------------------------------\nYou can also check your registration status any time here:\n${portalLink}` +
    `\n\nIf you ordered food, you can also reach Crave & Crumb on WhatsApp:\nhttps://wa.me/${WA_NUMBER}`;

  const htmlBody =
    `<div style="font-family:Arial,Helvetica,sans-serif;color:#111111;line-height:1.6;max-width:480px;">` +
    `<p>Hi ${esc(name)},</p>` +
    `<p>You're verified for <strong>${esc(EVENT_NAME)}</strong>.</p>` +
    `<p>` +
    `<strong>Access code:</strong> ${esc(code)}<br>` +
    `<strong>Date &amp; time:</strong> ${esc(EVENT_DATE)}<br>` +
    `<strong>Venue:</strong> ${esc(EVENT_VENUE)}` +
    `</p>` +
    `<p>Please keep your access code handy for entry.</p>` +
    `<p>For questions, call <strong>${esc(SUPPORT_PHONE)}</strong>.</p>` +
    `<p>See you there.</p>` +
    `<p style="margin:28px 0 0 0;padding-top:22px;border-top:1px solid #e5e5e5;">` +
    `You can also check your registration status any time here:<br><br>` +
    `<a href="${esc(portalLink)}" style="background-color:#6FCB53;color:#08170F;padding:14px 26px;border-radius:8px;text-decoration:none;font-weight:bold;display:inline-block;">View my registration</a>` +
    `</p>` +
    `</div>`;

  MailApp.sendEmail({
    to: email,
    subject: `${EVENT_NAME} — You're verified! Here are your details`,
    body: plainBody,
    htmlBody: htmlBody,
  });
}

function doPost(e) {
  const sheet = getSheet();
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: "Invalid request body" });
  }
  const action = data.action || "register";

  if (action === "register") {
    const name = String(data.name || "").trim();
    const email = String(data.email || "").trim();
    const phone = String(data.phone || "").trim();
    const address = String(data.address || "").trim();
    const id = String(data.id || "").trim();

    if (!name || !isValidEmail(email) || !phone || !address) {
      return jsonOut({ ok: false, error: "Missing required details" });
    }
    if (!id) {
      return jsonOut({ ok: false, error: "Missing registration id" });
    }

    // Idempotency: if this exact id already exists (e.g. a retried send), don't create a duplicate.
    const existing = findRowById(sheet, id);
    if (existing !== -1) {
      return jsonOut({ ok: true, alreadySaved: true });
    }

    sheet.appendRow([
      data.submittedAt || new Date().toISOString(),
      id,
      name,
      phone,
      email,
      address,
      String(data.question || "").trim(),
      (data.food || []).join(", "),
      Number(data.total) || 0,
      false, // Verified
      "", // Access Code
      false, // Purchased
    ]);

    try {
      sendRegistrationReceivedEmail(email, name);
    } catch (err) {
      // Email failed (e.g. quota); registration row is still saved — let the client report this so
      // the user isn't told the email definitely went out.
      return jsonOut({ ok: true, emailQueued: false });
    }

    return jsonOut({ ok: true, emailQueued: true });
  }

  // Every action below this point changes admin-controlled data, so it requires the admin key.
  if (data.key !== ADMIN_KEY) {
    return jsonOut({ ok: false, error: "Unauthorized" });
  }

  if (action === "verify" || action === "resend") {
    const row = findRowById(sheet, data.id);
    if (row === -1)
      return jsonOut({ ok: false, error: "Registration not found" });

    const alreadyVerified = rowCell(sheet, row, COL.VERIFIED) === true;
    let code = rowCell(sheet, row, COL.CODE);

    // Verify: generate + store a code only if the row isn't verified yet.
    if (!alreadyVerified) {
      code = generateAccessCode();
      setCell(sheet, row, COL.VERIFIED, true);
      setCell(sheet, row, COL.CODE, code);
    }

    if (action === "resend" && !alreadyVerified) {
      // Nothing meaningful to resend for an unverified row.
      return jsonOut({ ok: false, error: "Not verified yet" });
    }

    if (!code) {
      return jsonOut({ ok: false, error: "No access code on file" });
    }

    const email = rowCell(sheet, row, COL.EMAIL);
    const name = rowCell(sheet, row, COL.NAME);
    if (!isValidEmail(email)) {
      return jsonOut({ ok: false, error: "Invalid email on file" });
    }

    try {
      sendVerifiedEmail(email, name, code);
    } catch (err) {
      return jsonOut({ ok: false, error: "Could not send email" });
    }

    return jsonOut({ ok: true, alreadyDone: action === "resend" });
  }

  if (action === "purchase") {
    const row = findRowById(sheet, data.id);
    if (row === -1)
      return jsonOut({ ok: false, error: "Registration not found" });
    const already = rowCell(sheet, row, COL.PURCHASED) === true;
    if (already) return jsonOut({ ok: true, alreadyDone: true });
    setCell(sheet, row, COL.PURCHASED, true);
    return jsonOut({ ok: true });
  }

  return jsonOut({ ok: false, error: "Unknown action" });
}

function doGet(e) {
  const sheet = getSheet();

  // Self-service lookup by access code — a guest checking their own status.
  // Only returns a single, already-verified record; never the full list.
  if (e.parameter.code) {
    const row = findRowByCode(sheet, e.parameter.code);
    if (row === -1) return jsonOut({ error: "Not found" });

    const verified = rowCell(sheet, row, COL.VERIFIED) === true;
    if (!verified) return jsonOut({ error: "Not found" });

    return jsonOut({
      name: rowCell(sheet, row, COL.NAME),
      food: String(rowCell(sheet, row, COL.FOOD) || "")
        .split(", ")
        .filter(Boolean),
      total: rowCell(sheet, row, COL.TOTAL),
      verified: true,
      purchased: rowCell(sheet, row, COL.PURCHASED) === true,
      venue: EVENT_VENUE,
      eventDate: EVENT_DATE,
    });
  }

  // Full list — organizer view only, requires the admin key.
  if (e.parameter.key !== ADMIN_KEY) {
    return jsonOut({ error: "Unauthorized" });
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonOut([]);

  const rows = values
    .slice(1)
    .filter((r) => String(r[COL.ID - 1] || "").trim() !== "") // skip stray/empty trailing rows
    .map((row) => ({
      submittedAt: row[COL.SUBMITTED - 1],
      id: row[COL.ID - 1],
      name: row[COL.NAME - 1],
      phone: row[COL.PHONE - 1],
      email: row[COL.EMAIL - 1],
      address: row[COL.ADDRESS - 1],
      question: row[COL.QUESTION - 1],
      food: String(row[COL.FOOD - 1] || "")
        .split(", ")
        .filter(Boolean),
      total: row[COL.TOTAL - 1],
      verified: row[COL.VERIFIED - 1] === true,
      accessCode: row[COL.CODE - 1],
      purchased: row[COL.PURCHASED - 1] === true,
    }));

  return jsonOut(rows);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
