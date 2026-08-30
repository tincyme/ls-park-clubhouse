// ═══════════════════════════════════════════════════════════════════════
// ROLES: Admin / Manager access control
// ═══════════════════════════════════════════════════════════════════════
//
// SETUP STEPS:
// 1. In your Google Sheet, add a new tab named exactly: Staff
// 2. Add these exact column headers in row 1: Email, Name, Role, Active
// 3. Add one row per person, e.g.:
//      tincy@example.com | Tincy      | Admin   | Yes
//      staff@example.com | Front Desk | Manager | Yes
//    Role must be typed exactly "Admin" or "Manager". Active must be
//    typed as the plain word "Yes" — NOT a checkbox (checkboxes store
//    TRUE/FALSE, not the text "Yes", and won't match).
// 4. Paste this entire file into your Apps Script project as a file
//    named Roles, REPLACING any previous version.
// 5. IMPORTANT — if you previously ran "installAdminMenuTrigger": go to
//    the Triggers page (clock icon, left sidebar) and DELETE any
//    trigger for "buildAdminMenu_". Not used anymore.
// 6. Make sure "function onOpen()" appears EXACTLY ONCE across your
//    whole project.
// 7. Save. No manual "Run" step needed.
// 8. Reload the Sheet — EVERYONE now sees the SAME full menu, including
//    items marked "(Admin only)". This is intentional (see the comment
//    on onOpen() below for why) — actual permission enforcement happens
//    when an item is clicked, not by hiding it from view. A Manager who
//    clicks an Admin-only item gets a clear "you don't have permission"
//    message instead of it running.

// Looks up the current viewer's role from the Staff sheet. Deliberately
// uses getActiveSpreadsheet() rather than openById() — the former only
// needs permission to read the spreadsheet that's already open, which
// works inside onOpen()'s restricted execution mode with no extra
// authorization. openById() requires broader permissions that caused
// onOpen() to fail silently for anyone who hadn't separately authorized
// the project.
function getCurrentUserRole_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return null; // couldn't determine who's asking — treat as no access

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff");
  if (!sheet) return null; // Staff sheet not set up yet

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const match = rows.find(function (row) {
    return String(row[0]).trim().toLowerCase() === email.toLowerCase() &&
           String(row[3]).trim().toLowerCase() === "yes";
  });

  return match ? String(match[2]).trim() : null;
}

// Call this as the first line of any function that should be restricted.
// Returns true and lets the function continue if the current user's role
// is in allowedRoles. Returns false and shows a clear message otherwise
// — the calling function should immediately "return" when this is false.
function requireRole_(allowedRoles) {
  const role = getCurrentUserRole_();
  if (role && allowedRoles.indexOf(role) !== -1) return true;

  SpreadsheetApp.getUi().alert(
    "You don't have permission to do this.\n\n" +
    "This action requires: " + allowedRoles.join(" or ") + " access.\n" +
    "Contact the Admin if you need this."
  );
  return false;
}

// Shows a Yes/No confirmation with a clear summary of what's about to
// happen. Returns true only if the user explicitly clicks Yes. Use this
// right before any setup/config change actually takes effect — after
// gathering all the details, so the summary can be specific.
function confirmAction_(summary) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("Please confirm", summary, ui.ButtonSet.YES_NO);
  return response === ui.Button.YES;
}

// Shows the result of an action, with a clickable "Send WhatsApp
// message" button pre-filled with the customer's number and message,
// if a phone number is available. Falls back to a plain alert if not.
// True automatic (no-click) WhatsApp sending would require the paid
// WhatsApp Business API — this is the free alternative: one tap sends
// it from your own WhatsApp instead of typing it out manually.
function showResultWithWhatsApp_(title, message, customerPhone, whatsappText, emailSent) {
  const ui = SpreadsheetApp.getUi();

  if (!customerPhone) {
    ui.alert(title, message + (emailSent ? "\n\n(Customer notified by email.)" : "\n\n(No phone or email on file — customer not notified.)"), ui.ButtonSet.OK);
    return;
  }

  const digitsOnly = String(customerPhone).replace(/\D/g, "");
  const waUrl = "https://wa.me/91" + digitsOnly + "?text=" + encodeURIComponent(whatsappText);

  const emailNote = emailSent ? "Customer was also emailed automatically.<br><br>" : "";

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:6px;font-size:13px;line-height:1.5;">' +
    '<p style="white-space:pre-wrap;">' + message.replace(/</g, "&lt;") + '</p>' +
    '<p style="color:#666;">' + emailNote + '</p>' +
    '<a href="' + waUrl + '" target="_blank" ' +
    'style="display:inline-block;background:#25D366;color:#fff;padding:10px 20px;' +
    'border-radius:22px;text-decoration:none;font-weight:bold;">📱 Send WhatsApp message</a>' +
    '</div>'
  ).setWidth(420).setHeight(240);

  ui.showModalDialog(html, title);
}

// Routes an admin action (Approve, Cancel, Block, etc.) through the
// deployed Web App instead of writing to the Sheet directly under the
// clicking user's own permissions. The Web App always runs under the
// script owner's authorization, so this works correctly even if the
// operational sheets are protected to Admin-only direct editing —
// Managers' menu clicks still function normally.
//
// ADMIN_ACTION_SECRET and WEB_APP_URL are defined in phase2_booking.gs.
function callAdminAction_(action, params) {
  try {
    const payload = {
      type: "admin_action",
      secret: ADMIN_ACTION_SECRET,
      action: action,
      params: params
    };

    const response = UrlFetchApp.fetch(WEB_APP_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    return JSON.parse(response.getContentText());
  } catch (e) {
    return { success: false, error: "Could not reach the server: " + e.toString() };
  }
}

// Menu — shows the SAME full menu to everyone with edit access, rather
// than trying to hide items per role. This is deliberate: identifying
// WHO is viewing (Session.getActiveUser()) is unreliable inside onOpen()
// specifically for personal Gmail accounts not on a Workspace domain —
// Google restricts this for privacy reasons, and no code change fixes
// it. The actual security still works correctly, because every
// Admin-only function below checks requireRole_(["Admin"]) at the
// moment it's CLICKED — a different, reliable execution context (this
// is exactly how "Check my access" correctly identifies each person).
// A Manager who clicks an Admin-only item simply gets a clear
// "you don't have permission" message instead of it running.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("L's Park Tools")
    .addItem("Check my access", "checkMyAccess_")
    .addSeparator()
    .addItem("Send follow-up reminder now", "sendFollowUpReminders")
    .addSeparator()
    .addItem("Reject booking request", "rejectBookingRequest")
    .addSeparator()
    .addItem("Cancel a booking", "cancelBooking")
    .addItem("Mark payment received (confirms a request, or updates an existing booking)", "markPaymentReceived")
    .addItem("Block a slot", "blockSlot")
    .addItem("Unblock a slot", "unblockSlot")
    .addSeparator()
    .addItem("Setup status dropdown + colours (Admin only)", "setupStatusManagement")
    .addItem("Install 12-hr reminder trigger (Admin only)", "installFollowUpTrigger")
    .addItem("Remove reminder trigger (Admin only)", "removeFollowUpTrigger")
    .addSeparator()
    .addItem("Add Payment button columns to sheets (Admin only)", "setupOneClickPaymentColumns")
    .addItem("Install one-click Payment buttons (Admin only)", "installPaymentActionTrigger")
    .addItem("Remove one-click Payment buttons (Admin only)", "removePaymentActionTrigger")
    .addSeparator()
    .addItem("Open a reserved Badminton slot (Admin only)", "openReservedBadmintonSlot")
    .addItem("Close a reserved Badminton slot (Admin only)", "closeReservedBadmintonSlot")
    .addToUi();
}

// Self-diagnostic — shows EXACTLY what the script sees for you: your
// detected email, whether it's found in Staff, and what Role/Active
// values are actually stored there. Available to everyone, so anyone
// can figure out why their access isn't working without guessing.
function checkMyAccess_() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getActiveUser().getEmail();

  if (!email) {
    ui.alert(
      "Could not detect your email at all.\n\n" +
      "This can happen depending on how the Sheet is shared with you. " +
      "Ask the Admin to check the Sheet's sharing settings."
    );
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff");
  if (!sheet) {
    ui.alert("Detected email: " + email + "\n\nBut no sheet named exactly \"Staff\" was found. Ask the Admin to create it.");
    return;
  }

  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];

  const match = rows.find(function (row) {
    return String(row[0]).trim().toLowerCase() === email.toLowerCase();
  });

  if (!match) {
    ui.alert(
      "Detected email: " + email + "\n\n" +
      "This exact email was NOT found in column A of the Staff sheet.\n\n" +
      "Common causes: a typo, extra space, or you're logged into a " +
      "different Google account than the one added to Staff."
    );
    return;
  }

  const roleValue = String(match[2]).trim();
  const activeValue = String(match[3]).trim();
  const roleOk = roleValue === "Admin" || roleValue === "Manager";
  const activeOk = activeValue.toLowerCase() === "yes";

  ui.alert(
    "Detected email: " + email + "\n\n" +
    "Found in Staff sheet:\n" +
    "  Role column says: \"" + roleValue + "\"\n" +
    "  Active column says: \"" + activeValue + "\"\n\n" +
    (roleOk && activeOk
      ? "This is set up correctly — you should have access. If the menu still shows Request Access, try closing and reopening the Sheet tab completely."
      : "PROBLEM FOUND:\n" +
        (!roleOk ? "- Role must be spelled exactly \"Admin\" or \"Manager\" (capital first letter, no extra spaces/typos). Currently: \"" + roleValue + "\"\n" : "") +
        (!activeOk ? "- Active must say exactly \"Yes\" as plain text. If you used a checkbox instead, it stores TRUE/FALSE, not \"Yes\" — retype it as plain text \"Yes\". Currently: \"" + activeValue + "\"\n" : ""))
  );
}