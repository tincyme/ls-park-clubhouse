// ═══════════════════════════════════════════════════════════════════════
// FOLLOW-UPS & STATUS FORMATTING (reconstructed 27 Aug 2026)
// ═══════════════════════════════════════════════════════════════════════
//
// WHY THIS FILE EXISTS: your live Apps Script project had 4 menu items
// wired up in roles.gs -- sendFollowUpReminders, setupStatusManagement,
// installFollowUpTrigger, removeFollowUpTrigger -- with no matching code
// anywhere in code.gs or roles.gs, and no copy in git history either.
// The functions themselves were gone from the live project too, so this
// is a REBUILT replacement based on the menu labels and how the rest of
// the system works, not a recovery of the original code. Two design
// choices below were genuinely unknowable and I picked a reasonable
// default -- both are called out inline. Read a booking or two through
// once after installing, before you trust it on autopilot.
//
// SETUP STEPS:
// 1. In the Apps Script editor, add a new file named exactly: FollowUps
// 2. Paste this entire file's contents into it.
// 3. Save. Deploy > Manage deployments > pencil icon > Version: New
//    version > Deploy (same two-step rule as every other backend change
//    -- saving alone does not update the live /exec URL, and this file
//    doesn't touch doGet/doPost anyway, so no redeploy is strictly
//    required for the menu items to work, but do it anyway so the
//    Apps Script "current version" stays in sync with what's pasted in).
// 4. Reload the Sheet. The 4 menu items now work.
// 5. Nothing else needs to change in code.gs or roles.gs -- this file
//    reuses their existing helpers (requireRole_, findCustomerById_,
//    sendCustomerEmail_, sendOwnerNotificationEmail_, ALL_FACILITY_IDS,
//    facilityNameById_, SPREADSHEET_ID).

// How long to wait after a reminder before that same request/booking is
// eligible for another one. Prevents the 12-hour trigger from emailing
// the same customer every 12 hours forever until someone actions it.
const FOLLOWUP_REMINDER_COOLDOWN_HOURS = 24;

// Don't nag a customer about a request the moment they submit it --
// only once it's genuinely been sitting a while.
const PENDING_REQUEST_REMINDER_MIN_AGE_HOURS = 24;

// ── Detecting whether we're running from a menu click or a trigger ─────
// A time-based trigger has no UI to show -- SpreadsheetApp.getUi() throws
// in that context. Used to skip UI-only steps (role check, the on-screen
// WhatsApp dialog) when this runs unattended overnight.
function isInteractive_() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

// Reads a header cell; if it's blank, writes the given label into it.
// Self-healing schema, same pattern as the Dashboard's Amount Charged
// check -- lets this file safely add its own tracking column on first
// run instead of requiring you to edit the sheet by hand first.
function ensureColumnHeader_(sheet, colIndex1Based, label) {
  const headerCell = sheet.getRange(1, colIndex1Based);
  if (!String(headerCell.getValue()).trim()) {
    headerCell.setValue(label);
  }
}

function hoursSince_(dateValue) {
  if (!dateValue) return Infinity;
  const d = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
  if (isNaN(d.getTime())) return Infinity;
  return (new Date().getTime() - d.getTime()) / (1000 * 60 * 60);
}

// ── Send follow-up reminder now (menu item + the 12-hr trigger) ────────
//
// DESIGN CHOICE (this part was genuinely unknowable, not recovered):
// reminds about (a) Booking Requests still New/Under Review after 24hrs,
// and (b) Confirmed bookings with Payment Status Partial/Unpaid whose
// Booking Date hasn't passed yet -- a booking that already happened
// without full payment is a collections conversation, not something to
// auto-email about, so those are left out and only listed for you.
//
// Sends the customer an automatic email (if they have one on file) and,
// when run interactively, shows you a one-tap WhatsApp link per item --
// true unattended WhatsApp sending isn't available (see roles.gs), so
// the trigger-driven overnight run only sends the email half.
function sendFollowUpReminders() {
  const interactive = isInteractive_();
  if (interactive && !requireRole_(["Admin", "Manager"])) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date();
  const items = []; // { id, kind, customerName, phone, email, message, whatsappText }

  // -- Pending Booking Requests --------------------------------------
  const reqSheet = ss.getSheetByName("Booking Requests");
  ensureColumnHeader_(reqSheet, 14, "Last Reminder Sent"); // column N
  const reqLastRow = reqSheet.getLastRow();
  if (reqLastRow >= 2) {
    const reqRows = reqSheet.getRange(2, 1, reqLastRow - 1, 14).getValues();
    reqRows.forEach(function (row, i) {
      const status = String(row[10]).trim(); // K Status
      if (status !== "New" && status !== "Under Review") return;
      if (hoursSince_(row[11]) < PENDING_REQUEST_REMINDER_MIN_AGE_HOURS) return; // L Created Date
      if (hoursSince_(row[13]) < FOLLOWUP_REMINDER_COOLDOWN_HOURS) return; // N Last Reminder Sent

      const requestId = row[0];
      const customer = findCustomerById_(row[1]); // B Customer ID
      if (!customer) return;
      const facilityName = facilityNameById_(row[2]); // C Facility ID
      const dateStr = row[4]; // E Request Date
      const slot = row[5]; // F Slot/Requested Time

      const message =
        "Hi " + customer.name + ", quick update on your " + facilityName + " booking request " +
        "(" + requestId + ") for " + dateStr + (slot ? ", " + slot : "") + " -- " +
        "we're still reviewing it and wanted to let you know it hasn't been forgotten. " +
        "We'll confirm shortly. If your plans have changed, just reply and let us know.";

      items.push({
        id: requestId, kind: "Pending request", customerName: customer.name,
        phone: customer.phone, email: customer.email, message: message
      });

      sendCustomerEmail_(customer.email, "Your L's Park booking request " + requestId + " -- still in review", message);
      reqSheet.getRange(2 + i, 14).setValue(now); // N Last Reminder Sent
    });
  }

  // -- Confirmed & not fully paid, booking date not yet passed --------
  const bkSheet = ss.getSheetByName("Bookings");
  ensureColumnHeader_(bkSheet, 16, "Last Reminder Sent"); // column P (O is reserved for Amount Charged)
  const bkLastRow = bkSheet.getLastRow();
  if (bkLastRow >= 2) {
    const bkRows = bkSheet.getRange(2, 1, bkLastRow - 1, 16).getValues();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    bkRows.forEach(function (row, i) {
      const status = String(row[10]).trim(); // K Status
      const paymentStatus = String(row[11]).trim(); // L Payment Status
      if (status !== "Confirmed") return;
      if (paymentStatus !== "Partial" && paymentStatus !== "Unpaid") return;

      const bookingDateVal = row[5]; // F Booking Date
      const bookingDate = (bookingDateVal instanceof Date) ? bookingDateVal : new Date(bookingDateVal);
      if (isNaN(bookingDate.getTime()) || bookingDate < todayStart) return; // skip past-dated bookings

      if (hoursSince_(row[15]) < FOLLOWUP_REMINDER_COOLDOWN_HOURS) return; // P Last Reminder Sent

      const bookingId = row[0];
      const customer = findCustomerById_(row[2]); // C Customer ID
      if (!customer) return;
      const facilityName = facilityNameById_(row[3]); // D Facility ID

      const message =
        "Hi " + customer.name + ", reminder about your confirmed " + facilityName + " booking " +
        "(" + bookingId + ") on " + row[5] + " -- payment is currently marked " + paymentStatus.toLowerCase() +
        ". Please arrange the remaining payment before the event. Reply here if you have questions.";

      items.push({
        id: bookingId, kind: "Unpaid balance", customerName: customer.name,
        phone: customer.phone, email: customer.email, message: message
      });

      sendCustomerEmail_(customer.email, "Reminder: payment pending for booking " + bookingId, message);
      bkSheet.getRange(2 + i, 16).setValue(now); // P Last Reminder Sent
    });
  }

  // -- Owner summary (always sent, interactive or not) -----------------
  const summaryLines = items.length
    ? items.map(function (it) { return "- [" + it.kind + "] " + it.id + " -- " + it.customerName; })
    : ["No pending requests or unpaid confirmed bookings needed a reminder this run."];
  sendOwnerNotificationEmail_(
    "Follow-up reminders run: " + items.length + " sent",
    summaryLines.join("\n")
  );

  if (!interactive) return; // trigger-driven run ends here -- no UI available

  const ui = SpreadsheetApp.getUi();
  if (!items.length) {
    ui.alert("Follow-up reminders", "Nothing needed a reminder right now.", ui.ButtonSet.OK);
    return;
  }

  const rows = items.map(function (it) {
    const waButton = it.phone
      ? '<a href="https://wa.me/91' + String(it.phone).replace(/\D/g, "") + '?text=' + encodeURIComponent(it.message) + '" target="_blank" ' +
        'style="display:inline-block;background:#25D366;color:#fff;padding:6px 14px;border-radius:16px;text-decoration:none;font-weight:bold;font-size:12px;">📱 WhatsApp</a>'
      : '<span style="color:#999;font-size:12px;">no phone on file</span>';
    return '<tr><td style="padding:6px;border-bottom:1px solid #eee;">' + it.kind + '<br><b>' + it.id + '</b></td>' +
      '<td style="padding:6px;border-bottom:1px solid #eee;">' + it.customerName + (it.email ? '<br><span style="color:#666;font-size:11px;">emailed</span>' : '<br><span style="color:#c00;font-size:11px;">no email on file</span>') + '</td>' +
      '<td style="padding:6px;border-bottom:1px solid #eee;text-align:right;">' + waButton + '</td></tr>';
  }).join("");

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:13px;">' +
    '<p>' + items.length + ' reminder(s) sent by email. Tap WhatsApp for any that need a manual nudge too:</p>' +
    '<table style="width:100%;border-collapse:collapse;">' + rows + '</table>' +
    '</div>'
  ).setWidth(480).setHeight(Math.min(500, 120 + items.length * 50));
  ui.showModalDialog(html, "Follow-up reminders sent");
}

// ── Install / remove the 12-hour automatic trigger (Admin only) ────────

function installFollowUpTrigger() {
  if (!requireRole_(["Admin"])) return;
  const ui = SpreadsheetApp.getUi();

  const existing = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === "sendFollowUpReminders";
  });
  if (!confirmAction_(
    (existing.length ? "Replacing " + existing.length + " existing reminder trigger(s). " : "") +
    "Install a trigger that runs sendFollowUpReminders automatically every 12 hours, " +
    "emailing customers with pending requests or unpaid confirmed bookings?"
  )) return;

  existing.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("sendFollowUpReminders").timeBased().everyHours(12).create();
  ui.alert("Installed. sendFollowUpReminders will now run automatically every 12 hours.");
}

function removeFollowUpTrigger() {
  if (!requireRole_(["Admin"])) return;
  const ui = SpreadsheetApp.getUi();

  const triggers = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === "sendFollowUpReminders";
  });
  triggers.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ui.alert(triggers.length
    ? "Removed " + triggers.length + " reminder trigger(s). \"Send follow-up reminder now\" still works manually."
    : "No reminder trigger was installed -- nothing to remove.");
}

// ── Status dropdown + colour coding (Admin only) ────────────────────────
//
// DESIGN CHOICE (this part was genuinely unknowable, not recovered --
// you weren't sure either what this originally did). This is a
// reasonable best guess: a dropdown + colour-coded cells for the Status
// column on both Bookings and Booking Requests, matching the exact
// status values those columns actually use elsewhere in the code. It
// only touches data validation and cell colour -- never the data itself
// -- so it's safe to run again, or to redo by hand in Format > Conditional
// formatting if this isn't what you remember.
function setupStatusManagement() {
  if (!requireRole_(["Admin"])) return;
  if (!confirmAction_(
    "Add a dropdown and colour-coding to the Status column on Bookings and Booking Requests? " +
    "This only changes formatting/validation, not your data, and can be re-run safely."
  )) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  applyStatusFormatting_(ss.getSheetByName("Bookings"), 11, [
    { value: "Confirmed", bg: "#d9ead3", fg: "#274e13" },
    { value: "Cancelled", bg: "#f4cccc", fg: "#660000" }
  ]);

  applyStatusFormatting_(ss.getSheetByName("Booking Requests"), 11, [
    { value: "New", bg: "#cfe2f3", fg: "#1c4587" },
    { value: "Under Review", bg: "#fff2cc", fg: "#7f6000" },
    { value: "Approved", bg: "#d9ead3", fg: "#274e13" },
    { value: "Rejected", bg: "#f4cccc", fg: "#660000" },
    { value: "Cancelled", bg: "#efefef", fg: "#666666" }
  ]);

  SpreadsheetApp.getUi().alert("Done. Status dropdown + colours applied to Bookings and Booking Requests.");
}

// colIndex1Based: the Status column (K = 11 on both sheets here).
function applyStatusFormatting_(sheet, colIndex1Based, statuses) {
  const range = sheet.getRange(2, colIndex1Based, Math.max(sheet.getMaxRows() - 1, 1), 1);

  // Dropdown -- warns rather than blocks, so it won't fight any status
  // value already sitting in the sheet that isn't in this list.
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statuses.map(function (s) { return s.value; }), true)
    .setAllowInvalid(true)
    .build();
  range.setDataValidation(rule);

  // Remove any previous colour rules this function added for this exact
  // range (so re-running doesn't pile up duplicate rules), then add
  // fresh ones.
  const a1 = range.getA1Notation();
  const rules = sheet.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (rg) { return rg.getA1Notation() === a1; });
  });
  statuses.forEach(function (s) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(s.value)
        .setBackground(s.bg)
        .setFontColor(s.fg)
        .setRanges([range])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);
}
