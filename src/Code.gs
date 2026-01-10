/***** CONFIG *****/
const SHEET_NAME = "Sheet1"; // rename if your tab is different
const TIMEZONE = "America/Toronto";

// Put your calendar IDs here
const CALENDAR_IDS = {
  Gryffindor: "c_08692d5205f8c5b8b7f7a2ccf9e10ee5101960ea9df40447feac59db88c3b3a5@group.calendar.google.com",
  Hufflepuff: "c_fc75a6aa5d3d70223c225dac5d102fb1c8320c9fc389b1ae88d85b797bd685ab@group.calendar.google.com",
  Ravenclaw: "c_37a015d269eb93ed78c238ecccdf74db15453079c3f2e40463ab7758fdbe5eb0@group.calendar.google.com"
};

// Set these in Project Settings → Script Properties
// API_KEY: a random string (site must send it)
// ADMIN_EMAIL: where approvals go (your email)
function getProps_() {
  const p = PropertiesService.getScriptProperties();
  return {
    apiKey: p.getProperty("API_KEY"),
    adminEmail: p.getProperty("ADMIN_EMAIL"),
    baseUrl: p.getProperty("BASE_URL") // optional, can be set after deployment
  };
}

/***** WEB APP ENDPOINT: accept request from site *****/

function doPost(e) {
  const data = e.parameter || {};

  // Validate basics
  const name = String(data.name || "").trim();
  const email = String(data.email || "").trim();
  const room = String(data.room || "").trim();
  const checkInStr = String(data.checkIn || "").trim();
  const checkOutStr = String(data.checkOut || "").trim();
  const notes = String(data.notes || "").trim();

  if (!name || !email || !room || !checkInStr || !checkOutStr) {
    return HtmlService.createHtmlOutput("Missing required fields.");
  }
  if (!CALENDAR_IDS[room]) {
    return HtmlService.createHtmlOutput("Invalid room.");
  }

  const requestId = Utilities.getUuid();
  const adminToken = Utilities.getUuid();
  const status = "Pending";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

  sh.appendRow([
    new Date(),
    requestId,
    status,
    name,
    email,
    room,
    checkInStr,
    checkOutStr,
    notes,
    adminToken,
    "" // EventId
  ]);

  // Email you the approval links
  sendApprovalEmail_({
    requestId,
    adminToken,
    name,
    email,
    room,
    checkInStr,
    checkOutStr,
    notes
  });

  // Return link back to your site (no auto-redirect)
  // Set BASE_URL in Script Properties to "https://healeycottage.com" when you go live.
  const { baseUrl } = getProps_();
  const siteBase = (baseUrl && String(baseUrl).trim()) || "http://localhost:4321";
  const redirectUrl = `${siteBase.replace(/\/+$/, "")}/request/success`;

  const html = `
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Request received</title>
        <style>
          :root { color-scheme: light; }
          body {
            margin: 0;
            font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
            line-height: 1.4;
            background: #f7f7f7;
            color: #111;
          }
          .wrap {
            min-height: 100vh;
            display: grid;
            place-items: center;
            padding: 24px;
          }
          .card {
            width: 100%;
            max-width: 520px;
            background: #fff;
            border: 1px solid rgba(0,0,0,.12);
            border-radius: 16px;
            padding: 18px 18px 16px;
            box-shadow: 0 10px 25px rgba(0,0,0,.06);
          }
          h1 {
            margin: 0 0 8px;
            font-size: 18px;
          }
          p { margin: 0 0 10px; }
          .muted { color: rgba(0,0,0,.65); font-size: 14px; }
          .row {
            margin-top: 14px;
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            align-items: center;
          }
          a.btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 10px 12px;
            border-radius: 12px;
            border: 1px solid rgba(0,0,0,.12);
            background: #111;
            color: #fff;
            text-decoration: none;
            font-weight: 600;
          }
          a.btn:hover { filter: brightness(1.05); }
        </style>
      </head>
      <body>
        <div class="wrap">
          <div class="card" role="status" aria-live="polite">
            <h1>Request received</h1>
            <p class="muted">Thanks. We’ll review it and email you once it’s approved or declined. You can return to the site using the button below.</p>

            <div class="row">
              <a class="btn" href="${redirectUrl}" target="_top" rel="noopener">Return to Healey Cottage</a>
            </div>
          </div>
        </div>
      </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***** APPROVE / DECLINE HANDLER *****/

function doGet(e) {
  const action = ((e && e.parameter && e.parameter.action) ? e.parameter.action : "").toLowerCase();

  // If someone just opens the web app URL in a browser, show a simple OK page.
  if (!action) {
    return HtmlService.createHtmlOutput("OK 1333");
  }

  const requestId = e.parameter.id;
  const token = e.parameter.token;

  if (!["approve", "decline"].includes(action) || !requestId || !token) {
    return HtmlService.createHtmlOutput("Invalid link.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sh.getDataRange().getValues();

  // Find row by RequestID
  const header = data[0];
  const idx = (name) => header.indexOf(name);

  const idCol = idx("RequestID");
  const statusCol = idx("Status");
  const tokenCol = idx("AdminToken");
  const eventIdCol = idx("EventId");
  const nameCol = idx("Name");
  const emailCol = idx("Email");
  const roomCol = idx("Room");
  const checkInCol = idx("CheckIn");
  const checkOutCol = idx("CheckOut");
  const notesCol = idx("Notes");

  if ([idCol, statusCol, tokenCol, eventIdCol, nameCol, emailCol, roomCol, checkInCol, checkOutCol].some(c => c === -1)) {
    return HtmlService.createHtmlOutput("Sheet headers don't match expected format.");
  }

  let rowIndex = -1;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idCol]) === String(requestId)) {
      rowIndex = r + 1; // sheet is 1-indexed
      break;
    }
  }
  if (rowIndex === -1) return HtmlService.createHtmlOutput("Request not found.");

  const row = sh.getRange(rowIndex, 1, 1, header.length).getValues()[0];

  if (String(row[tokenCol]) !== String(token)) {
    return HtmlService.createHtmlOutput("Invalid token.");
  }

  const currentStatus = String(row[statusCol]);
  if (currentStatus !== "Pending") {
    return HtmlService.createHtmlOutput(`Already ${currentStatus}.`);
  }

  const guestName = String(row[nameCol]);
  const guestEmail = String(row[emailCol]);
  const room = String(row[roomCol]);
  const checkInStr = String(row[checkInCol]);
  const checkOutStr = String(row[checkOutCol]);
  const notes = String(row[notesCol] || "");

  if (action === "decline") {
    sh.getRange(rowIndex, statusCol + 1).setValue("Declined");
    sendGuestEmail_(guestEmail, "Request declined", declineBody_({ guestName, room, checkInStr, checkOutStr }));
    return HtmlService.createHtmlOutput("Declined. Guest has been emailed.");
  }

  // Approve: create calendar event (busy)
  const calId = CALENDAR_IDS[room];
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) return HtmlService.createHtmlOutput("Calendar not found or not accessible.");

const start = parseDate_(row[checkInCol]);
const end = parseDate_(row[checkOutCol]);

if (!start || !end || end <= start) {
  return HtmlService.createHtmlOutput(`Bad dates in sheet: checkIn="${row[checkInCol]}", checkOut="${row[checkOutCol]}"`);
}

  const title = `Booked: ${guestName}`;
  const desc = notes ? `Notes: ${notes}\n\nRequestID: ${requestId}` : `RequestID: ${requestId}`;

  const event = cal.createAllDayEvent(title, start, end, { description: desc });

  sh.getRange(rowIndex, statusCol + 1).setValue("Approved");
  sh.getRange(rowIndex, eventIdCol + 1).setValue(event.getId());

  const ics = makeStayIcs_({
    requestId,
    guestName,
    room,
    startDate: start,
    endDate: end
  });

  sendGuestEmail_(
    guestEmail,
    "Request approved",
    approveBody_({ guestName, room, checkInStr, checkOutStr }),
    { attachments: [ics] }
  );

  return HtmlService.createHtmlOutput("Approved. Calendar updated and guest emailed.");
}

/***** EMAILS *****/
function sendApprovalEmail_(req) {
  const { adminEmail } = getProps_();
  if (!adminEmail) throw new Error("ADMIN_EMAIL is not set in Script Properties.");

  const webAppUrl = ScriptApp.getService().getUrl();
  const approveUrl = `${webAppUrl}?action=approve&id=${encodeURIComponent(req.requestId)}&token=${encodeURIComponent(req.adminToken)}`;
  const declineUrl = `${webAppUrl}?action=decline&id=${encodeURIComponent(req.requestId)}&token=${encodeURIComponent(req.adminToken)}`;

  const subject = `Healey Cottage request: ${req.room} (${req.checkInStr} → ${req.checkOutStr})`;

  const html = `
    <div style="font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; line-height:1.4;">
      <p style="margin:0 0 10px;">
        <b>${escapeHtml_(req.name)}</b> requested <b>${escapeHtml_(req.room)}</b>.
      </p>

      <p style="margin:0 0 10px;">
        <b>Dates:</b> ${escapeHtml_(req.checkInStr)} → ${escapeHtml_(req.checkOutStr)}<br/>
        <b>Email:</b> ${escapeHtml_(req.email)}
      </p>

      ${req.notes ? `<p style="margin:0 0 10px;"><b>Notes:</b><br/>${escapeHtml_(req.notes)}</p>` : ""}

      <p style="margin:14px 0 12px;">
        <a href="${approveUrl}" style="display:inline-block; padding:10px 14px; background:#16a34a; color:white; text-decoration:none; border-radius:10px; font-weight:600;">Approve</a>
        <span style="display:inline-block; width:10px;"></span>
        <a href="${declineUrl}" style="display:inline-block; padding:10px 14px; background:#dc2626; color:white; text-decoration:none; border-radius:10px; font-weight:600;">Decline</a>
      </p>

      <p style="margin:0; color:#6b7280; font-size:12px;">
        Tip: approval will create a busy block on the room calendar and email the guest (with an optional calendar hold attachment).
      </p>

      <p style="margin:8px 0 0; color:#6b7280; font-size:12px;">RequestID: ${escapeHtml_(req.requestId)}</p>
    </div>
  `;

  MailApp.sendEmail({
    to: adminEmail,
    subject,
    htmlBody: html
  });
}

function sendGuestEmail_(to, subject, body, opts = {}) {
  const payload = {
    to,
    subject: `Healey Cottage: ${subject}`,
    htmlBody: body,
    name: "Healey Cottage"
  };

  if (opts.attachments && Array.isArray(opts.attachments) && opts.attachments.length) {
    payload.attachments = opts.attachments;
  }
  if (opts.replyTo) {
    payload.replyTo = opts.replyTo;
  }

  MailApp.sendEmail(payload);
}

function approveBody_({ guestName, room, checkInStr, checkOutStr }) {
  return `
    <p>Hi ${escapeHtml_(guestName)},</p>

    <p>Your request has been approved. We’re looking forward to having you.</p>

    <p>
      <b>Room:</b> ${escapeHtml_(room)}<br/>
      <b>Dates:</b> ${escapeHtml_(checkInStr)} → ${escapeHtml_(checkOutStr)}
    </p>

    <p>
      <b>Address:</b><br/>
      1107 Peter Road<br/>
      Bracebridge, ON, P1L 1X3
    </p>

    <p>
      I’ve attached a calendar hold you can add if helpful.
    </p>

    <p>
      Please remember to bring sunscreen, bug spray, bathing suits, and flip flops or water shoes.
      And feel free to reach out if you’d like to coordinate anything else to bring, such as food or alcohol.
    </p>

    <p>If anything changes, just reply to this email.</p>

    <p>See you up north,<br/>Kyle & Brandon</p>
  `;
}

// UPDATED decline copy
function declineBody_({ guestName, room, checkInStr, checkOutStr }) {
  return `
    <p>Hi ${escapeHtml_(guestName)},</p>

    <p>Thanks for the request. Unfortunately we can’t accommodate these dates as requested.</p>

    <p>
      <b>Room:</b> ${escapeHtml_(room)}<br/>
      <b>Dates:</b> ${escapeHtml_(checkInStr)} → ${escapeHtml_(checkOutStr)}
    </p>

    <p>
      There's a scheduling constraint on our side.
      We’ll reach out shortly to discuss options that work.
    </p>

    <p>Talk soon,<br/>Kyle & Brandon</p>
  `;
}

/***** HELPERS *****/
function parseDate_(value) {
  // If Sheets stored it as a Date object already
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const s = String(value || "").trim();
  if (!s) return null;

  // If it's already YYYY-MM-DD (from the HTML date input)
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (iso) {
    const y = Number(iso[1]);
    const m = Number(iso[2]);
    const d = Number(iso[3]);
    return new Date(y, m - 1, d);
  }

  // Fallback: try letting JS parse things like 7/10/2026
  const d2 = new Date(s);
  if (!isNaN(d2.getTime())) {
    return new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
  }

  return null;
}

function json_(obj, code) {
  const out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  // Apps Script doesn’t let us set status codes directly for web apps reliably;
  // we include ok/error in payload. Kept 'code' for readability.
  return out;
}

function escapeHtml_(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

/***** CALENDAR INVITE (.ics) *****/
function makeStayIcs_({ requestId, guestName, room, startDate, endDate }) {
  // All-day event: DTEND is exclusive (ideal for stays).
  const uid = `${requestId}@healeycottage.com`;
  const dtstamp = formatUtcTimestamp_(new Date());

  const dtstart = formatIcsDate_(startDate);
  const dtend = formatIcsDate_(endDate);

  const summary = `Healey Cottage — ${room}`;
  const descriptionLines = [
    `Guest: ${guestName}`,
    `Room: ${room}`,
    ``,
    `Address:`,
    `1107 Peter Road`,
    `Bracebridge, ON, P1L 1X3`,
    ``,
    `This is a hold for your stay at Healey Cottage.`,
    `If plans change, reply to the approval email.`,
    ``,
    `Bring: sunscreen, bug spray, bathing suits, flip flops or water shoes.`,
    `Feel free to reach out to coordinate food or alcohol.`,
    ``,
    `healeycottage.com`
  ];

  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Healey Cottage//Booking//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${dtstamp}`,
    `DTSTART;VALUE=DATE:${dtstart}`,
    `DTEND;VALUE=DATE:${dtend}`,
    `SUMMARY:${escapeIcsText_(summary)}`,
    `LOCATION:${escapeIcsText_("1107 Peter Road, Bracebridge, ON, P1L 1X3")}`,
    `DESCRIPTION:${escapeIcsText_(descriptionLines.join("\\n"))}`,
    "STATUS:CONFIRMED",
    "TRANSP:OPAQUE",
    "END:VEVENT",
    "END:VCALENDAR"
  ];

  const icsText = lines.map(foldIcsLine_).join("\r\n") + "\r\n";
  return Utilities.newBlob(icsText, "text/calendar; charset=utf-8", "healey-cottage.ics");
}

function formatIcsDate_(d) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d, tz, "yyyyMMdd");
}

function formatUtcTimestamp_(d) {
  return Utilities.formatDate(d, "UTC", "yyyyMMdd'T'HHmmss'Z'");
}

function escapeIcsText_(s) {
  return String(s)
    .replace(/\\/g, "\\\\")
    .replace(/\r?\n/g, "\\n")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,");
}

// Basic RFC line folding (good enough for our short fields)
function foldIcsLine_(line) {
  const max = 73;
  if (line.length <= max) return line;
  const out = [];
  let i = 0;
  while (i < line.length) {
    out.push(line.slice(i, i + max));
    i += max;
  }
  return out.join("\r\n ");
}