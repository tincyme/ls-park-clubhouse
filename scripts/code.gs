// ═══════════════════════════════════════════════════════════════════════
// PHASE 2: Check Availability + Booking Requests
// ═══════════════════════════════════════════════════════════════════════
//
// SETUP STEPS — READ CAREFULLY (duplicate function names cause silent bugs):
// 1. Open Code.gs (or wherever your current doGet/doPost live).
// 2. DELETE your existing doGet function and existing doPost function
//    COMPLETELY — select from "function doGet(" to its closing "}" and
//    remove it. Same for doPost. Do not leave the old versions in the
//    file alongside the new ones below, even in a different file in the
//    same project — Apps Script treats all files as one shared scope, so
//    two functions named doPost would silently conflict.
// 3. Paste this entire file's contents into Code.gs (or as a new file —
//    doesn't matter, as long as the old doGet/doPost are gone).
// 4. Update onOpen() in status_management.gs to the version I've already
//    updated for you (adds "Approve booking request" / "Reject booking
//    request" menu items).
// 5. Save. Then Deploy > Manage deployments > pencil icon > Version:
//    New version > Deploy. (Reminder: editing code alone does NOT update
//    your live /exec URL until you do this step.)
//
// Your Booking Requests and Bookings sheets must already have these
// exact column headers (from your workbook):
//   Booking Requests: Request ID | Customer ID | Facility ID | Resource ID |
//     Request Date | Slot/Requested Time | Booking Mode | Guests |
//     Event Type | Special Request | Status | Created Date | Owner Notes
//   Bookings: Booking ID | Request ID | Customer ID | Facility ID |
//     Resource ID | Booking Date | Start Time | End Time | Booking Mode |
//     Guests | Status | Payment Status | Created Date | Notes
//     -- Payment Status (column L) is the ONLY place payment is tracked.
//     A pending request has no Bookings row yet, so it can't be "unpaid"
//     there -- instead, a New request simply can't become a Bookings row
//     at all until "Mark payment received" is run on it (see
//     confirmBookingWithPaymentCore_). That's what makes an unpaid
//     request unconfirmable, without a second Payment Status column on
//     Booking Requests.
//   Blocked Slots: Block ID | Facility ID | Resource ID | Block Date |
//     Start Time | End Time | Reason | Status | Notes
//   Customers: Customer ID | Name | Mobile | Email | Marketing Opt-In |
//     Created Date | Notes

const SPREADSHEET_ID = "1EzdjUeww3jK6xBOU0opPYYxPT1F2kE38XJvq3OgwTHs";

// Used so menu-triggered admin actions (Approve, Cancel, Block, etc.)
// write to the Sheet through the Web App — which always runs under the
// script owner's own permissions — rather than under whichever person
// clicked the menu. This lets operational sheets (Bookings, Booking
// Requests, etc.) be protected to Admin-only DIRECT editing in Sheets'
// native protection settings, while the menu tools keep working for
// Managers too, since they no longer rely on the clicking user's own
// edit rights.
//
// This secret stops a random person who only knows the public /exec
// URL (visible in the website's source) from crafting a fake admin
// action request directly. It does NOT need to be hidden from anyone
// who already has Editor access to this Apps Script project — that's
// the same trust level as being able to click the menu items anyway.
// Change this to your own random string if you want; it just needs to
// match between here and its use in roles.gs's callAdminAction_.
const ADMIN_ACTION_SECRET = "peedspark-admin-9f3a7c2e1b6d4a58";

// Same URL as ENQUIRY_API in index.html — update both together if you
// ever redeploy to a new Web App URL.
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwSGAt5wWrnMFWZ_vO8InNrLH1rd3vvmSpjS_k9evtJQfhEkeyPp5LrnREX0_866JnlCw/exec";

// Backup notification email — sent for every enquiry AND booking request,
// alongside the WhatsApp message. This exists because there's no reliable
// way to detect from the website whether a customer's WhatsApp actually
// opened (popup blockers and missing WhatsApp installs behave differently
// across browsers/devices), so email acts as a guaranteed second channel.
const OWNER_NOTIFY_EMAIL = "tincye29@gmail.com";

// ── Spam protection ──────────────────────────────────────────────────
//
// The Web App URL is public (visible in the page source), so anyone —
// not just visitors using the actual website — could send requests
// directly to it. These checks apply to BOTH enquiries and booking
// requests, before anything is saved or emailed.

const RATE_LIMIT_MAX_PER_WINDOW = 3;      // max submissions...
const RATE_LIMIT_WINDOW_SECONDS = 1800;   // ...per phone number, per 30 minutes

// Returns true if this phone number is still within its allowed rate,
// and records this attempt. Returns false if it should be blocked.
function checkRateLimit_(phone) {
  const cache = CacheService.getScriptCache();
  const key = "rl_" + String(phone).replace(/\D/g, ""); // digits only, so formatting differences don't create separate buckets
  if (!key || key === "rl_") return true; // no phone provided — let normal validation catch that instead

  const countStr = cache.get(key);
  const count = countStr ? parseInt(countStr, 10) : 0;

  if (count >= RATE_LIMIT_MAX_PER_WINDOW) {
    return false;
  }

  cache.put(key, String(count + 1), RATE_LIMIT_WINDOW_SECONDS);
  return true;
}

function sendOwnerNotificationEmail_(subject, body) {
  try {
    MailApp.sendEmail(OWNER_NOTIFY_EMAIL, subject, body);
  } catch (e) {
    // Never let an email failure break the actual enquiry/booking save.
    Logger.log("Owner notification email failed: " + e.toString());
  }
}

// Phase 2 covers Hall/Lawn (Fixed Slot) facilities only.
// Swimming Pool / Badminton hourly logic comes in Phase 4.
const FACILITY_IDS = {
  "AC Hall": "F002",
  "Non-AC Hall": "F003",
  "Lawn": "F004"
};

const SLOT_TIMES = {
  "Morning": { start: "08:00", end: "14:00" },
  "Evening": { start: "16:00", end: "22:00" }
};

// Full facility list — used for blocking slots (owner may need to block
// Pool/Badminton too, even though online booking for those isn't live yet).
const ALL_FACILITY_IDS = {
  "AC Hall": "F002",
  "Non-AC Hall": "F003",
  "Lawn": "F004",
  "Swimming Pool": "F001",
  "Badminton": "F005"
};

// ── Phase 4: hourly facilities (Swimming Pool, Badminton) ──────────────

const HOURLY_FACILITY_IDS = {
  "Swimming Pool": "F001",
  "Badminton": "F005"
};

const POOL_MAX_CAPACITY = 8;

// Badminton has 2 independent courts, tracked separately. If you add a
// 3rd court later, just add another {id, name} entry here — make sure
// the ID also exists as a row in your Resources sheet.
const BADMINTON_COURTS = [
  { id: "R001", name: "Court 1" },
  { id: "R002", name: "Court 2" }
];

// These hours are BLOCKED BY DEFAULT for Badminton (reserved for
// members) — the opposite of the normal Blocked Slots sheet, which
// blocks specific one-off times. To open one of these hours for public
// booking, the admin uses "Open a reserved Badminton slot" from the
// Sheet menu, which creates an entry in the "Badminton Unblocks" sheet.
const BADMINTON_RESERVED_WINDOWS = [
  { start: "05:00", end: "08:00" },
  { start: "17:00", end: "23:00" }
];

// Per-facility operating windows. Badminton has member-reserved slots
// as early as 5 AM and as late as 11 PM, so it needs a wider window than
// the Pool.
const FACILITY_OPERATING_HOURS = {
  "Swimming Pool": { start: 6, end: 20 },   // 6 AM – 8 PM
  "Badminton": { start: 5, end: 23 }        // 5 AM – 11 PM
};

// ── NEW: which Booking Requests statuses still hold a slot ─────────────
//
// A request that hasn't been Approved, Rejected, or Cancelled yet must
// still block the slot for everyone else — otherwise two people could
// request the same slot before the owner acts on the first one.
const PENDING_REQUEST_STATUSES = ["New", "Under Review"];

// Returns every row from "Booking Requests" whose Status is still New
// or Under Review — i.e. still unresolved and still holding its slot.
function getPendingBookingRequestRows_(ss) {
  const sheet = ss.getSheetByName("Booking Requests");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  return rows.filter(function (row) {
    return PENDING_REQUEST_STATUSES.indexOf(String(row[10]).trim()) !== -1;
  });
}

function generateHourlySlots_(facilityName) {
  const hours = FACILITY_OPERATING_HOURS[facilityName] || { start: 6, end: 20 };
  const slots = [];
  for (let h = hours.start; h < hours.end; h++) {
    const start = (h < 10 ? "0" + h : h) + ":00";
    const endHour = h + 1;
    const end = (endHour < 10 ? "0" + endHour : endHour) + ":00";
    slots.push({ start: start, end: end });
  }
  return slots;
}

function timeToMinutes_(hhmm) {
  const parts = String(hhmm).split(":");
  const hours = parseInt(parts[0], 10) || 0;
  const minutes = parseInt(parts[1], 10) || 0;
  return (hours * 60) + minutes;
}

function normalizeDateStr_(value) {
  if (!value) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  // Plain text (our guaranteed format going forward) — trust it as-is,
  // no Date parsing, no timezone conversion, no ambiguity.
  return String(value).trim();
}

function normalizeTimeStr_(value) {
  if (!value) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(value).trim();
}

// Appends a row, then force-formats specific columns (1-based indexes) as
// plain text and re-writes their values as literal strings. This prevents
// Google Sheets from silently auto-converting date/time-looking strings
// into its internal date/time type, which caused availability checks to
// mismatch due to timezone conversion differences between the Spreadsheet
// and Apps Script project settings.
function appendRowSafely_(sheet, values, textColumnIndexes) {
  sheet.appendRow(values);
  const newRow = sheet.getLastRow();
  textColumnIndexes.forEach(function (colIndex) {
    const cell = sheet.getRange(newRow, colIndex);
    cell.setNumberFormat("@");
    cell.setValue(String(values[colIndex - 1]));
  });
  return newRow;
}

// ── Availability check (used by doGet) ────────────────────────────────

function checkFixedSlotAvailability(facilityName, dateStr) {
  const facilityId = FACILITY_IDS[String(facilityName).trim()];
  if (!facilityId) {
    return { success: false, error: "Facility not supported for online availability yet. Please enquire directly." };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const bookingsSheet = ss.getSheetByName("Bookings");
  const bookingsLastRow = bookingsSheet.getLastRow();
  const bookingRows = bookingsLastRow >= 2 ? bookingsSheet.getRange(2, 1, bookingsLastRow - 1, 14).getValues() : [];

  const blocksSheet = ss.getSheetByName("Blocked Slots");
  const blocksLastRow = blocksSheet.getLastRow();
  const blockRows = blocksLastRow >= 2 ? blocksSheet.getRange(2, 1, blocksLastRow - 1, 9).getValues() : [];

  const targetDate = String(dateStr).trim();

  // NEW: unresolved requests for this facility+date, so they hold the
  // slot too — not just Confirmed bookings.
  const pendingRows = getPendingBookingRequestRows_(ss).filter(function (row) {
    return String(row[2]).trim() === facilityId && normalizeDateStr_(row[4]) === targetDate;
  });

  const results = {};

  Object.keys(SLOT_TIMES).forEach(function (slotName) {
    const slotStart = SLOT_TIMES[slotName].start;
    const slotEnd = SLOT_TIMES[slotName].end;
    const slotStartMin = timeToMinutes_(slotStart);
    const slotEndMin = timeToMinutes_(slotEnd);

    const isBooked = bookingRows.some(function (row) {
      const rFacilityId = String(row[3]).trim();          // Facility ID
      const rDate = normalizeDateStr_(row[5]);             // Booking Date
      const rStart = normalizeTimeStr_(row[6]);            // Start Time
      const rEnd = normalizeTimeStr_(row[7]);              // End Time
      const rStatus = String(row[10]).trim();              // Status

      return rFacilityId === facilityId &&
             rDate === targetDate &&
             rStart === slotStart &&
             rEnd === slotEnd &&
             rStatus === "Confirmed";
    });

    // NEW: exact slot-name match ("Morning"/"Evening"), same as how the
    // Slot column is stored for these facilities.
    const isPending = !isBooked && pendingRows.some(function (row) {
      return String(row[5]).trim() === slotName;
    });

    const isBlocked = !isBooked && !isPending && blockRows.some(function (row) {
      const rFacilityId = String(row[1]).trim();  // Facility ID
      const rDate = normalizeDateStr_(row[3]);     // Block Date
      const rStatus = String(row[7]).trim();       // Status

      if (rFacilityId !== facilityId || rDate !== targetDate || rStatus !== "Active") return false;

      // Overlap check — a block only needs to touch part of the slot to
      // make the whole slot unavailable (e.g. 2-hour maintenance block
      // inside a 6-hour Morning slot).
      const rStartMin = timeToMinutes_(normalizeTimeStr_(row[4]));
      const rEndMin = timeToMinutes_(normalizeTimeStr_(row[5]));
      return rStartMin < slotEndMin && rEndMin > slotStartMin;
    });

    results[slotName] = isBooked ? "Booked" : (isPending ? "Pending" : (isBlocked ? "Blocked" : "Available"));
  });

  return { success: true, type: "fixed", slots: results };
}

// Checks actual remaining Pool capacity for every hour spanned by
// [startTime, endTime) on the given date, using a FRESH read of the
// Bookings sheet (never cached), since this must be correct at both
// booking-request time and again at approval time — capacity can
// change between those two moments as other requests get approved.
// Returns { minRemaining, hasExclusive, worstHour }.
function checkPoolCapacityForRange_(dateStr, startTime, endTime) {
  const targetDate = String(dateStr).trim();
  const rangeStartMin = timeToMinutes_(startTime);
  const rangeEndMin = timeToMinutes_(endTime);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const bookingsSheet = ss.getSheetByName("Bookings");
  const lastRow = bookingsSheet.getLastRow();
  const bookingRows = lastRow >= 2 ? bookingsSheet.getRange(2, 1, lastRow - 1, 14).getValues() : [];

  // NEW: unresolved requests for the Pool on this date, so a pending
  // guest count/exclusive request eats into remaining capacity too.
  const pendingRows = getPendingBookingRequestRows_(ss).filter(function (row) {
    return String(row[2]).trim() === "F001" && normalizeDateStr_(row[4]) === targetDate;
  });

  const hourlySlots = generateHourlySlots_("Swimming Pool").filter(function (slot) {
    const slotStartMin = timeToMinutes_(slot.start);
    const slotEndMin = timeToMinutes_(slot.end);
    return slotStartMin < rangeEndMin && slotEndMin > rangeStartMin; // overlaps the requested range
  });

  let minRemaining = POOL_MAX_CAPACITY;
  let hasExclusive = false;
  let worstHour = "";

  hourlySlots.forEach(function (slot) {
    const slotStartMin = timeToMinutes_(slot.start);
    const slotEndMin = timeToMinutes_(slot.end);

    const overlapping = bookingRows.filter(function (row) {
      const rFacilityId = String(row[3]).trim();
      const rDate = normalizeDateStr_(row[5]);
      const rStart = normalizeTimeStr_(row[6]);
      const rEnd = normalizeTimeStr_(row[7]);
      const rStatus = String(row[10]).trim();
      if (rFacilityId !== "F001" || rDate !== targetDate || rStatus !== "Confirmed") return false;
      const rStartMin = timeToMinutes_(rStart);
      const rEndMin = timeToMinutes_(rEnd);
      return rStartMin < slotEndMin && rEndMin > slotStartMin;
    });

    // NEW: same overlap test, applied to still-pending requests.
    const overlappingPending = pendingRows.filter(function (row) {
      const match = String(row[5] || "").match(/^(\d{2}:\d{2})-(\d{2}:\d{2})$/);
      if (!match) return false;
      const rStartMin = timeToMinutes_(match[1]);
      const rEndMin = timeToMinutes_(match[2]);
      return rStartMin < slotEndMin && rEndMin > slotStartMin;
    });

    const exclusiveHere = overlapping.some(function (row) { return String(row[8]).trim() === "Exclusive"; })
      || overlappingPending.some(function (row) { return String(row[6]).trim() === "Exclusive"; });

    const sharedGuests = overlapping
      .filter(function (row) { return String(row[8]).trim() !== "Exclusive"; })
      .reduce(function (sum, row) { return sum + (Number(row[9]) || 0); }, 0)
      + overlappingPending
        .filter(function (row) { return String(row[6]).trim() !== "Exclusive"; })
        .reduce(function (sum, row) { return sum + (Number(row[7]) || 0); }, 0);

    const remainingHere = exclusiveHere ? 0 : (POOL_MAX_CAPACITY - sharedGuests);

    if (exclusiveHere) hasExclusive = true;
    if (remainingHere < minRemaining) {
      minRemaining = remainingHere;
      worstHour = slot.start + "–" + slot.end;
    }
  });

  return { minRemaining: Math.max(minRemaining, 0), hasExclusive: hasExclusive, worstHour: worstHour };
}

// ── Find or create a customer record ───────────────────────────────────

function checkHourlyAvailability(facilityName, dateStr) {
  const facilityId = HOURLY_FACILITY_IDS[String(facilityName).trim()];
  if (!facilityId) {
    return { success: false, error: "Facility not recognized for hourly booking." };
  }

  const isPool = facilityName === "Swimming Pool";
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const bookingsSheet = ss.getSheetByName("Bookings");
  const bookingsLastRow = bookingsSheet.getLastRow();
  const bookingRows = bookingsLastRow >= 2 ? bookingsSheet.getRange(2, 1, bookingsLastRow - 1, 14).getValues() : [];

  const blocksSheet = ss.getSheetByName("Blocked Slots");
  const blocksLastRow = blocksSheet.getLastRow();
  const blockRows = blocksLastRow >= 2 ? blocksSheet.getRange(2, 1, blocksLastRow - 1, 9).getValues() : [];

  // Only relevant for Badminton — admin overrides that open up a
  // normally member-reserved hour for public booking.
  let unblockRows = [];
  if (!isPool) {
    const unblockSheet = ss.getSheetByName("Badminton Unblocks");
    if (unblockSheet) {
      const unblockLastRow = unblockSheet.getLastRow();
      unblockRows = unblockLastRow >= 2 ? unblockSheet.getRange(2, 1, unblockLastRow - 1, 7).getValues() : [];
    }
  }

  const targetDate = String(dateStr).trim();

  // NEW: unresolved requests for this facility+date.
  const pendingRows = getPendingBookingRequestRows_(ss).filter(function (row) {
    return String(row[2]).trim() === facilityId && normalizeDateStr_(row[4]) === targetDate;
  });

  const hourlySlots = generateHourlySlots_(facilityName);

  const slotResults = hourlySlots.map(function (slot) {
    const slotStartMin = timeToMinutes_(slot.start);
    const slotEndMin = timeToMinutes_(slot.end);

    // Overlap match (not exact match) — a multi-hour consecutive booking
    // like 17:00–20:00 must block EVERY individual hour it spans, not
    // just an hour whose start/end happens to match exactly.
    const overlappingBookings = bookingRows.filter(function (row) {
      const rFacilityId = String(row[3]).trim();   // Facility ID
      const rDate = normalizeDateStr_(row[5]);      // Booking Date
      const rStart = normalizeTimeStr_(row[6]);     // Start Time
      const rEnd = normalizeTimeStr_(row[7]);       // End Time
      const rStatus = String(row[10]).trim();       // Status

      if (rFacilityId !== facilityId || rDate !== targetDate || rStatus !== "Confirmed") return false;

      const rStartMin = timeToMinutes_(rStart);
      const rEndMin = timeToMinutes_(rEnd);
      return rStartMin < slotEndMin && rEndMin > slotStartMin;
    });

    // NEW: same overlap test against still-pending requests.
    const overlappingPending = pendingRows.filter(function (row) {
      const match = String(row[5] || "").match(/^(\d{2}:\d{2})-(\d{2}:\d{2})$/);
      if (!match) return false;
      const rStartMin = timeToMinutes_(match[1]);
      const rEndMin = timeToMinutes_(match[2]);
      return rStartMin < slotEndMin && rEndMin > slotStartMin;
    });

    const overlappingBlocks = blockRows.filter(function (row) {
      const rFacilityId = String(row[1]).trim();
      const rDate = normalizeDateStr_(row[3]);
      const rStatus = String(row[7]).trim();
      if (rFacilityId !== facilityId || rDate !== targetDate || rStatus !== "Active") return false;

      const rStartMin = timeToMinutes_(normalizeTimeStr_(row[4]));
      const rEndMin = timeToMinutes_(normalizeTimeStr_(row[5]));
      return rStartMin < slotEndMin && rEndMin > slotStartMin;
    });

    if (isPool) {
      const isBlocked = overlappingBlocks.length > 0; // Pool has no sub-resources, any block covers the whole pool
      if (isBlocked) {
        return { start: slot.start, end: slot.end, status: "Blocked" };
      }

      const hasExclusive = overlappingBookings.some(function (row) { return String(row[8]).trim() === "Exclusive"; })
        || overlappingPending.some(function (row) { return String(row[6]).trim() === "Exclusive"; });

      const sharedGuests = overlappingBookings
        .filter(function (row) { return String(row[8]).trim() !== "Exclusive"; })
        .reduce(function (sum, row) { return sum + (Number(row[9]) || 0); }, 0)
        + overlappingPending
          .filter(function (row) { return String(row[6]).trim() !== "Exclusive"; })
          .reduce(function (sum, row) { return sum + (Number(row[7]) || 0); }, 0);

      const remaining = hasExclusive ? 0 : (POOL_MAX_CAPACITY - sharedGuests);

      // NEW: if capacity is only used up by pending (not yet Confirmed)
      // requests, show "Pending" instead of a flat "Full" so it's clear
      // this could free up again if the owner rejects the request.
      const filledOnlyByPending = remaining <= 0 && overlappingBookings.length === 0 && overlappingPending.length > 0;

      return {
        start: slot.start,
        end: slot.end,
        status: remaining > 0 ? "Available" : (filledOnlyByPending ? "Pending" : "Full"),
        remaining: Math.max(remaining, 0),
        capacity: POOL_MAX_CAPACITY
      };
    }

    // Badminton — 2 independent courts, tracked separately. Precedence:
    // 1. A confirmed booking always shows as Booked.
    // 2. NEW: an unresolved (pending) request for that specific court
    //    shows as Pending — holds the court until approved/rejected.
    // 3. An owner maintenance block always shows as Blocked.
    // 4. Otherwise, if the hour falls in a member-reserved window and
    //    hasn't been explicitly opened via Badminton Unblocks, it shows
    //    as Blocked by default.
    // 5. Otherwise, Available.
    const facilityWideBlock = overlappingBlocks.some(function (row) { return !String(row[2]).trim(); });

    const inReservedWindow = BADMINTON_RESERVED_WINDOWS.some(function (w) {
      return timeToMinutes_(w.start) < slotEndMin && timeToMinutes_(w.end) > slotStartMin;
    });

    const courts = BADMINTON_COURTS.map(function (court) {
      const courtBooked = overlappingBookings.some(function (row) { return String(row[4]).trim() === court.id; }); // Resource ID
      if (courtBooked) {
        return { id: court.id, name: court.name, status: "Booked" };
      }

      const courtPending = overlappingPending.some(function (row) { return String(row[3]).trim() === court.id; }); // Resource ID
      if (courtPending) {
        return { id: court.id, name: court.name, status: "Pending" };
      }

      if (facilityWideBlock) {
        return { id: court.id, name: court.name, status: "Blocked" };
      }
      const courtBlocked = overlappingBlocks.some(function (row) { return String(row[2]).trim() === court.id; });
      if (courtBlocked) {
        return { id: court.id, name: court.name, status: "Blocked" };
      }

      if (inReservedWindow) {
        const isUnblocked = unblockRows.some(function (row) {
          const rDate = normalizeDateStr_(row[1]);       // Date
          const rStart = normalizeTimeStr_(row[2]);       // Start Time
          const rEnd = normalizeTimeStr_(row[3]);         // End Time
          const rResourceId = String(row[4]).trim();      // Resource ID (blank = both courts)
          const rStatus = String(row[5]).trim();           // Status

          if (rDate !== targetDate || rStatus !== "Active") return false;
          if (rResourceId && rResourceId !== court.id) return false;

          const rStartMin = timeToMinutes_(rStart);
          const rEndMin = timeToMinutes_(rEnd);
          return rStartMin < slotEndMin && rEndMin > slotStartMin;
        });

        if (!isUnblocked) {
          return { id: court.id, name: court.name, status: "Blocked" };
        }
      }

      return { id: court.id, name: court.name, status: "Available" };
    });

    const anyAvailable = courts.some(function (c) { return c.status === "Available"; });
    const allBlocked = courts.every(function (c) { return c.status === "Blocked"; });
    const overallStatus = anyAvailable ? "Available" : (allBlocked ? "Blocked" : "Full");

    return { start: slot.start, end: slot.end, status: overallStatus, courts: courts };
  });

  return {
    success: true,
    type: "hourly",
    bookingModel: isPool ? "capacity" : "resource",
    slots: slotResults
  };
}

// ── Unified dispatcher used by doGet ────────────────────────────────────

function checkAvailability(facilityName, dateStr) {
  const name = String(facilityName).trim();
  if (FACILITY_IDS[name]) {
    return checkFixedSlotAvailability(name, dateStr);
  }
  if (HOURLY_FACILITY_IDS[name]) {
    return checkHourlyAvailability(name, dateStr);
  }
  return { success: false, error: "Facility not recognized." };
}

// ── Find or create a customer record ───────────────────────────────────

function findOrCreateCustomer_(name, phone, email) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Customers");
  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 7).getValues() : [];

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][2]) === String(phone)) {
      return rows[i][0]; // existing Customer ID
    }
  }

  const customerId = "C-" + new Date().getTime();
  sheet.appendRow([customerId, name, phone, email || "", "No", new Date(), ""]);
  return customerId;
}

function facilityNameById_(facilityId) {
  const entry = Object.entries(ALL_FACILITY_IDS).find(function (pair) { return pair[1] === facilityId; });
  return entry ? entry[0] : facilityId;
}

// Looks up a customer's contact details by Customer ID. Returns
// { name, phone, email } or null if not found.
function findCustomerById_(customerId) {
  if (!customerId) return null;

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Customers");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const match = rows.find(function (row) { return String(row[0]).trim() === String(customerId).trim(); });
  if (!match) return null;

  return { name: match[1], phone: String(match[2]).trim(), email: String(match[3]).trim() };
}

// Emails the customer directly, if they provided an email — silently
// does nothing otherwise (WhatsApp is the fallback channel for those
// without an email on file, handled separately via a pre-filled link
// shown to the Admin/Manager after the action).
function sendCustomerEmail_(email, subject, body) {
  if (!email) return;
  try {
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log("Customer notification email failed: " + e.toString());
  }
}

// ── Create a booking request (called from doPost) ──────────────────────

function createBookingRequest(data) {
  const facilityName = String(data.facility).trim();
  const facilityId = FACILITY_IDS[facilityName] || HOURLY_FACILITY_IDS[facilityName] || "";

  // Enforce Pool capacity BEFORE saving anything — this is what actually
  // stops overbooking, since checking status on the site is only a
  // convenience for the customer, not a guarantee.
  if (facilityId === "F001") {
    const match = String(data.slot || "").match(/^(\d{2}:\d{2})-(\d{2}:\d{2})$/);
    if (match) {
      const capacity = checkPoolCapacityForRange_(data.date, match[1], match[2]);
      const requestedGuests = Number(data.guests) || 0;
      const mode = data.bookingMode === "Exclusive" ? "Exclusive" : "Shared";

      if (mode === "Exclusive" && capacity.minRemaining < POOL_MAX_CAPACITY) {
        return { success: false, error: "Sorry, that time already has a shared booking (around " + capacity.worstHour + "), so it can't be booked Exclusive. Please choose another time." };
      }
      if (mode === "Shared" && requestedGuests > capacity.minRemaining) {
        return { success: false, error: "Sorry, only " + capacity.minRemaining + " spot(s) remaining around " + capacity.worstHour + ". Please reduce your guest count or choose another time." };
      }
    }
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Booking Requests");

  const customerId = findOrCreateCustomer_(data.name, data.phone, data.email);
  const requestId = "REQ-" + new Date().getTime();
  const requestDate = String(data.date || "").trim();

  const values = [
    requestId,
    customerId,
    facilityId,
    data.resourceId || "",              // Court ID for Badminton, blank for other facilities
    requestDate,
    data.slot || "",                    // "Morning"/"Evening" for Hall/Lawn, or "HH:mm-HH:mm" for Pool/Badminton
    data.bookingMode || "Standard",     // "Shared"/"Exclusive" for Pool, "Standard" otherwise
    data.guests || "",
    data.eventType || "",
    data.specialRequest || "",
    "New",
    new Date(),
    ""
  ];

  // Column 5 = Request Date — force plain text so it always matches
  // exactly what the website's <input type="date"> sends (yyyy-MM-dd),
  // with no auto-conversion or timezone drift.
  appendRowSafely_(sheet, values, [5]);

  const courtLabel = data.resourceId
    ? " (" + (BADMINTON_COURTS.find(function (c) { return c.id === data.resourceId; }) || {}).name + ")"
    : "";

  sendOwnerNotificationEmail_(
    "New Booking Request: " + requestId,
    "A new booking request was submitted on the website.\n\n" +
    "Request ID: " + requestId + "\n" +
    "Name: " + data.name + "\n" +
    "Phone: " + data.phone + "\n" +
    "Email: " + (data.email || "-") + "\n" +
    "Facility: " + facilityName + courtLabel + "\n" +
    "Date: " + requestDate + "\n" +
    "Slot: " + (data.slot || "-") + "\n" +
    "Booking mode: " + (data.bookingMode || "Standard") + "\n" +
    "Guests: " + (data.guests || "-") + "\n" +
    "Event type: " + (data.eventType || "-") + "\n" +
    "Special request: " + (data.specialRequest || "-") + "\n\n" +
    "Open the Booking Requests sheet to review:\n" +
    "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/edit"
  );

  return { success: true, requestId: requestId };
}

// ── doGet: availability check + default health-check text ─────────────

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "availability") {
    const result = checkAvailability(e.parameter.facility, e.parameter.date);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput("PeedsPark Enquiry API is working")
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── doPost: existing enquiry logic + new booking request branch ────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Honeypot: a hidden field real visitors never see or fill, but
    // simple bots often auto-fill every input. If it has a value,
    // silently pretend success without saving or emailing anything —
    // this avoids tipping the bot off to try a different approach.
    if (data.honeypot) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, enquiryId: "ENQ-0000", requestId: "REQ-0000" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Rate limit: max 3 submissions per phone number per 30 minutes,
    // shared across enquiries AND booking requests.
    if (data.phone && !checkRateLimit_(data.phone)) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "Too many requests from this number. Please try again in a while, or call us directly." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Validation — enforced here too (not just on the website), so
    // direct requests to this endpoint can't skip it. Applies to
    // enquiries and booking requests; admin_action payloads don't use
    // these top-level fields, so this harmlessly no-ops for those.
    if (data.type !== "admin_action") {
      if (data.phone && !/^[0-9]{10}$/.test(String(data.phone).trim())) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: "Please enter a valid 10-digit mobile number." }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      if (data.date) {
        const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (String(data.date).trim() < todayStr) {
          return ContentService
            .createTextOutput(JSON.stringify({ success: false, error: "Please select today or a future date." }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    if (data.type === "booking_request") {
      const result = createBookingRequest(data);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.type === "admin_action") {
      if (data.secret !== ADMIN_ACTION_SECRET) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: "Unauthorized." }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      let result;
      switch (data.action) {
        case "rejectBookingRequest":
          result = rejectBookingRequestCore_(data.params.requestId);
          break;
        case "cancelBooking":
          result = cancelBookingCore_(data.params.bookingId);
          break;
        case "markPaymentReceived":
          result = markPaymentReceivedCore_(data.params.id, data.params.paymentStatus);
          break;
        case "blockSlot":
          result = blockSlotCore_(data.params.facilityName, data.params.resourceId, data.params.blockDate, data.params.startTime, data.params.endTime, data.params.reason);
          break;
        case "unblockSlot":
          result = unblockSlotCore_(data.params.blockId);
          break;
        case "openReservedBadmintonSlot":
          result = openReservedBadmintonSlotCore_(data.params.date, data.params.startTime, data.params.endTime, data.params.resourceId);
          break;
        case "closeReservedBadmintonSlot":
          result = closeReservedBadmintonSlotCore_(data.params.unblockId);
          break;
        default:
          result = { success: false, error: "Unknown action: " + data.action };
      }

      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Existing enquiry-saving behaviour, unchanged.
    const sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName("Enquiries");

    const enquiryId = "ENQ-" + new Date().getTime();

    sheet.appendRow([
      new Date(),
      enquiryId,
      data.name || "",
      data.phone || "",
      data.email || "",
      data.facility || "",
      data.date || "",
      data.guests || "",
      data.message || "",
      "New",
      data.source || ""
    ]);

    sendOwnerNotificationEmail_(
      "New Enquiry: " + enquiryId,
      "A new enquiry was submitted on the website.\n\n" +
      "Enquiry ID: " + enquiryId + "\n" +
      "Name: " + (data.name || "-") + "\n" +
      "Phone: " + (data.phone || "-") + "\n" +
      "Email: " + (data.email || "-") + "\n" +
      "Facility: " + (data.facility || "-") + "\n" +
      "Preferred date: " + (data.date || "-") + "\n" +
      "Guests: " + (data.guests || "-") + "\n" +
      "Message: " + (data.message || "-") + "\n" +
      "Source: " + (data.source || "-") + "\n\n" +
      "Open the Enquiries sheet to review:\n" +
      "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/edit"
    );

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        enquiryId: enquiryId
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Owner approval workflow (menu-driven) ───────────────────────────────

// Resolves a stored slot string into {start, end}. Handles named Fixed
// slots ("Morning"/"Evening") and hourly strings ("06:00-07:00").
function resolveSlotTimes_(slotName) {
  if (SLOT_TIMES[slotName]) return SLOT_TIMES[slotName];

  const match = String(slotName).match(/^(\d{2}:\d{2})-(\d{2}:\d{2})$/);
  if (match) return { start: match[1], end: match[2] };

  return null;
}

// ── Server-side core logic — called ONLY from doPost's admin_action
// branch, which always runs under the script owner's own permissions
// regardless of who triggered the menu click. This is what lets
// Managers use these tools even if the underlying sheets are protected
// to Admin-only direct editing.

// Facilities that take a 50% advance instead of full payment upfront.
// Pool and Badminton are paid in full early, so they're deliberately
// NOT in this list -- confirmBookingWithPaymentCore_ refuses a
// "Partial" payment status for any facility not listed here.
const PARTIAL_PAYMENT_ALLOWED_FACILITY_IDS = ["F002", "F003", "F004"]; // AC Hall, Non-AC Hall, Lawn

// The ONLY way a "New" booking request becomes a Confirmed Bookings row.
// There is no separate plain "Approve" anymore, and no Payment Status
// column on Booking Requests -- payment is recorded directly on the
// Bookings row (its existing Payment Status column) at the moment it's
// created, via paymentStatus ("Paid" or "Partial"). A request simply
// cannot become a booking without this being called, which is what
// makes an unpaid request unconfirmable.
function confirmBookingWithPaymentCore_(requestId, paymentStatus) {
  const status = String(paymentStatus || "").trim();
  if (status !== "Paid" && status !== "Partial") {
    return { success: false, error: "Payment status must be \"Paid\" or \"Partial\", got: " + paymentStatus };
  }

  const requestsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Booking Requests");
  const lastRow = requestsSheet.getLastRow();
  if (lastRow < 2) return { success: false, error: "No booking requests found." };

  const rows = requestsSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const rowIndex = rows.findIndex(function (row) { return row[0] === requestId; });
  if (rowIndex === -1) return { success: false, error: "Request ID not found: " + requestId };

  const row = rows[rowIndex];
  const [ , customerId, facilityId, resourceId, requestDate, slotName, bookingMode, guests ] = row;
  const currentReqStatus = String(row[10] || "").trim();
  if (currentReqStatus !== "New") {
    return { success: false, error: "Request " + requestId + " is already \"" + currentReqStatus + "\" -- only a New request can be confirmed." };
  }

  if (status === "Partial" && PARTIAL_PAYMENT_ALLOWED_FACILITY_IDS.indexOf(facilityId) === -1) {
    return { success: false, error: facilityNameById_(facilityId) + " must be paid in FULL to confirm -- Partial is only for Hall and Lawn bookings (50% advance). Collect the remaining payment first, or mark it Paid." };
  }

  const slotTimes = resolveSlotTimes_(slotName);
  if (!slotTimes) return { success: false, error: "Could not determine start/end time for slot: " + slotName };

  if (facilityId === "F001") {
    const capacity = checkPoolCapacityForRange_(requestDate, slotTimes.start, slotTimes.end);
    const requestedGuests = Number(guests) || 0;
    const mode = bookingMode === "Exclusive" ? "Exclusive" : "Shared";

    if (mode === "Exclusive" && capacity.minRemaining < POOL_MAX_CAPACITY) {
      return { success: false, error: "Cannot confirm as Exclusive: " + capacity.worstHour + " already has another Confirmed booking. Reject this request, or contact the customer to adjust before confirming." };
    }
    if (mode === "Shared" && requestedGuests > capacity.minRemaining) {
      return { success: false, error: "Cannot confirm — only " + capacity.minRemaining + " spot(s) left around " + capacity.worstHour + ", but this request is for " + requestedGuests + ". Reject this request, or contact the customer to reduce their guest count before confirming." };
    }
  }

  const bookingsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Bookings");
  const bookingId = "B-" + new Date().getTime();

  const values = [
    bookingId, requestId, customerId, facilityId, resourceId || "",
    String(requestDate).trim(), slotTimes.start, slotTimes.end,
    bookingMode || "Standard", guests || "", "Confirmed", status, new Date(), ""
  ];
  appendRowSafely_(bookingsSheet, values, [6, 7, 8]);

  requestsSheet.getRange(rowIndex + 2, 11).setValue("Approved");

  const customer = findCustomerById_(customerId);
  const facilityName = facilityNameById_(facilityId);
  const paymentNote = status === "Partial"
    ? "\n\nWe've received your 50% advance. The remaining balance is due before/at your slot."
    : "";
  const whatsappText =
    "Hello" + (customer ? " " + customer.name : "") + ",\n\n" +
    "Your booking is CONFIRMED!\n\n" +
    "Booking ID: " + bookingId + "\n" +
    "Facility: " + facilityName + "\n" +
    "Date: " + requestDate + "\n" +
    "Time: " + slotTimes.start + "–" + slotTimes.end + paymentNote + "\n\n" +
    "Thank you for choosing PeedsPark Club House!";

  if (customer && customer.email) {
    sendCustomerEmail_(customer.email, "Your PeedsPark booking is confirmed — " + bookingId, whatsappText);
  }

  return {
    success: true,
    message: "Booking " + bookingId + " confirmed for request " + requestId + " (Payment: " + status + ").",
    customerPhone: customer ? customer.phone : "",
    customerEmailSent: !!(customer && customer.email),
    whatsappText: whatsappText
  };
}

function rejectBookingRequestCore_(requestId) {
  const requestsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Booking Requests");
  const lastRow = requestsSheet.getLastRow();
  if (lastRow < 2) return { success: false, error: "No booking requests found." };

  const rows = requestsSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const rowIndex = rows.findIndex(function (row) { return row[0] === requestId; });
  if (rowIndex === -1) return { success: false, error: "Request ID not found: " + requestId };

  requestsSheet.getRange(rowIndex + 2, 11).setValue("Rejected");
  return { success: true, message: "Request " + requestId + " marked as Rejected." };
}

function cancelBookingCore_(bookingId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Bookings");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: "No bookings found." };

  const rows = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const rowIndex = rows.findIndex(function (row) { return String(row[0]).trim() === bookingId; });
  if (rowIndex === -1) return { success: false, error: "Booking ID not found: " + bookingId };

  sheet.getRange(rowIndex + 2, 11).setValue("Cancelled");

  const bookingRow = rows[rowIndex];
  const requestId = bookingRow[1];
  const customerId = bookingRow[2];
  const facilityId = bookingRow[3];
  const bookingDate = bookingRow[5];
  const startTime = bookingRow[6];
  const endTime = bookingRow[7];

  if (requestId) {
    const reqSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Booking Requests");
    const reqLastRow = reqSheet.getLastRow();
    if (reqLastRow >= 2) {
      const reqRows = reqSheet.getRange(2, 1, reqLastRow - 1, 13).getValues();
      const reqRowIndex = reqRows.findIndex(function (row) { return String(row[0]).trim() === requestId; });
      if (reqRowIndex !== -1) {
        reqSheet.getRange(reqRowIndex + 2, 11).setValue("Cancelled");
      }
    }
  }

  const customer = findCustomerById_(customerId);
  const facilityName = facilityNameById_(facilityId);
  const whatsappText =
    "Hello" + (customer ? " " + customer.name : "") + ",\n\n" +
    "Your booking has been CANCELLED.\n\n" +
    "Booking ID: " + bookingId + "\n" +
    "Facility: " + facilityName + "\n" +
    "Date: " + bookingDate + "\n" +
    "Time: " + startTime + "–" + endTime + "\n\n" +
    "Please contact us if you have any questions.";

  if (customer && customer.email) {
    sendCustomerEmail_(customer.email, "Your PeedsPark booking was cancelled — " + bookingId, whatsappText);
  }

  return {
    success: true,
    message: "Booking " + bookingId + " cancelled. That slot is now free again.",
    customerPhone: customer ? customer.phone : "",
    customerEmailSent: !!(customer && customer.email),
    whatsappText: whatsappText
  };
}

// Single entry point for "Mark payment received." One ID box, works two
// ways depending on what's typed in:
//   - a Request ID (REQ-...): the request is still pending -- this
//     confirms it (capacity checks, creates the Bookings row with this
//     Payment Status already set, notifies the customer). This is now
//     the ONLY way a New request becomes a Confirmed booking.
//   - a Booking ID (B-...): already confirmed -- this just updates the
//     Payment Status on that existing Bookings row (e.g. a Hall booking
//     going from Partial to Paid once the balance comes in) and sends a
//     lighter "payment received" note, without re-confirming anything.
// There is deliberately only ONE Payment Status column in the whole
// system -- on Bookings -- so nothing needs to be kept in sync.
function markPaymentReceivedCore_(id, paymentStatus) {
  const cleanId = String(id || "").trim();

  if (cleanId.indexOf("REQ-") === 0) {
    return confirmBookingWithPaymentCore_(cleanId, paymentStatus);
  }

  if (cleanId.indexOf("B-") === 0) {
    return updateBookingPaymentStatusCore_(cleanId, paymentStatus);
  }

  return { success: false, error: "\"" + id + "\" doesn't look like a Request ID (starts REQ-) or a Booking ID (starts B-)." };
}

// Updates the Payment Status on an ALREADY-CONFIRMED booking (e.g. Hall
// balance paid after the 50% advance). Does not touch Booking Requests
// or Bookings status -- the booking is already Confirmed.
function updateBookingPaymentStatusCore_(bookingId, paymentStatus) {
  const status = String(paymentStatus || "").trim();
  if (status !== "Paid" && status !== "Partial") {
    return { success: false, error: "Payment status must be \"Paid\" or \"Partial\", got: " + paymentStatus };
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Bookings");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: "No bookings found." };

  const rows = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const rowIndex = rows.findIndex(function (row) { return String(row[0]).trim() === bookingId; });
  if (rowIndex === -1) return { success: false, error: "Booking ID not found: " + bookingId };

  const bookingRow = rows[rowIndex];
  const facilityId = bookingRow[3];

  if (status === "Partial" && PARTIAL_PAYMENT_ALLOWED_FACILITY_IDS.indexOf(facilityId) === -1) {
    return { success: false, error: facilityNameById_(facilityId) + " must be paid in FULL -- Partial is only for Hall and Lawn bookings (50% advance)." };
  }

  sheet.getRange(rowIndex + 2, 12).setValue(status); // Payment Status column L

  const customerId = bookingRow[2];
  const bookingDate = bookingRow[5];
  const startTime = bookingRow[6];
  const endTime = bookingRow[7];

  const customer = findCustomerById_(customerId);
  const facilityName = facilityNameById_(facilityId);
  const paymentLine = status === "Paid"
    ? "We've received your payment in full. Thank you!"
    : "We've received your 50% advance. The remaining balance is due before/at your slot.";
  const whatsappText =
    "Hello" + (customer ? " " + customer.name : "") + ",\n\n" +
    paymentLine + "\n\n" +
    "Booking ID: " + bookingId + "\n" +
    "Facility: " + facilityName + "\n" +
    "Date: " + bookingDate + "\n" +
    "Time: " + startTime + "–" + endTime + "\n\n" +
    "We look forward to seeing you at PeedsPark Club House!";

  if (customer && customer.email) {
    sendCustomerEmail_(customer.email, "Payment received — " + bookingId, whatsappText);
  }

  return {
    success: true,
    message: "Booking " + bookingId + " Payment Status set to " + status + ".",
    customerPhone: customer ? customer.phone : "",
    customerEmailSent: !!(customer && customer.email),
    whatsappText: whatsappText
  };
}

function blockSlotCore_(facilityName, resourceId, blockDate, startTime, endTime, reason) {
  const facilityId = ALL_FACILITY_IDS[facilityName];
  if (!facilityId) return { success: false, error: "Facility not recognized: " + facilityName };

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Blocked Slots");
  const blockId = "BLK-" + new Date().getTime();

  const values = [blockId, facilityId, resourceId || "", blockDate, startTime, endTime, reason, "Active", ""];
  appendRowSafely_(sheet, values, [4, 5, 6]);

  const courtNote = resourceId ? (" — " + ((BADMINTON_COURTS.find(function (c) { return c.id === resourceId; }) || {}).name || resourceId) + " only") : "";
  return { success: true, message: "Blocked: " + blockId + "\n" + facilityName + courtNote + " on " + blockDate + ", " + startTime + "–" + endTime + "." };
}

function unblockSlotCore_(blockId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Blocked Slots");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: "No blocks found." };

  const rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const rowIndex = rows.findIndex(function (row) { return String(row[0]).trim() === blockId; });
  if (rowIndex === -1) return { success: false, error: "Block ID not found: " + blockId };

  sheet.getRange(rowIndex + 2, 8).setValue("Inactive");
  return { success: true, message: "Block " + blockId + " removed. That time is available again." };
}

function openReservedBadmintonSlotCore_(date, startTime, endTime, resourceId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Badminton Unblocks");
  if (!sheet) return { success: false, error: "Couldn't find a 'Badminton Unblocks' sheet. Please create one first with headers: Unblock ID | Date | Start Time | End Time | Resource ID | Status | Notes" };

  const unblockId = "UNB-" + new Date().getTime();
  const values = [unblockId, date, startTime, endTime, resourceId || "", "Active", ""];
  appendRowSafely_(sheet, values, [2, 3, 4]);

  const courtSummary = resourceId
    ? ((BADMINTON_COURTS.find(function (c) { return c.id === resourceId; }) || {}).name || resourceId)
    : "both courts";
  return { success: true, message: "Opened: " + unblockId + "\n" + date + ", " + startTime + "–" + endTime + " — " + courtSummary + "." };
}

function closeReservedBadmintonSlotCore_(unblockId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Badminton Unblocks");
  if (!sheet) return { success: false, error: "Couldn't find a 'Badminton Unblocks' sheet." };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: "No overrides found." };

  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const rowIndex = rows.findIndex(function (row) { return String(row[0]).trim() === unblockId; });
  if (rowIndex === -1) return { success: false, error: "Unblock ID not found: " + unblockId };

  sheet.getRange(rowIndex + 2, 6).setValue("Inactive");
  return { success: true, message: "Closed " + unblockId + ". That time is reserved for members again." };
}

// ── Menu-facing wrappers — gather input, check role, confirm, then
// route the actual write through callAdminAction_ (defined in roles.gs)
// instead of writing to the Sheet directly under the clicking user's
// own permissions.

function rejectBookingRequest() {
  if (!requireRole_(["Admin", "Manager"])) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Reject Booking Request", "Enter the Request ID to reject:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const requestId = response.getResponseText().trim();
  const result = callAdminAction_("rejectBookingRequest", { requestId: requestId });
  ui.alert(result.success ? result.message : ("Could not reject: " + result.error));
}

// ── Cancel an approved booking (frees the slot again) ───────────────────

function cancelBooking() {
  if (!requireRole_(["Admin", "Manager"])) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Cancel Booking", "Enter the Booking ID to cancel (e.g. B-1723...):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const bookingId = response.getResponseText().trim();
  const result = callAdminAction_("cancelBooking", { bookingId: bookingId });

  if (!result.success) {
    ui.alert("Could not cancel: " + result.error);
    return;
  }

  showResultWithWhatsApp_("Booking Cancelled", result.message, result.customerPhone, result.whatsappText, result.customerEmailSent);
}

// ── Mark a booking's payment as received ────────────────────────────────

// This is the ONLY way a "New" booking request becomes a Confirmed
// booking (there's no separate plain "Approve" anymore) -- run it here
// once the customer has paid, and it records the payment, confirms the
// booking, and notifies the customer in one go. Also works on an
// already-confirmed Booking ID, to update its Payment Status later
// (e.g. a Hall booking's Partial advance becoming Paid in full).
function markPaymentReceived() {
  if (!requireRole_(["Admin", "Manager"])) return;

  const ui = SpreadsheetApp.getUi();
  const idResp = ui.prompt(
    "Mark Payment Received",
    "Enter the Request ID (e.g. REQ-1723...) to confirm a pending request, " +
    "OR the Booking ID (e.g. B-1723...) to update payment on an already-confirmed booking:",
    ui.ButtonSet.OK_CANCEL
  );
  if (idResp.getSelectedButton() !== ui.Button.OK) return;
  const id = idResp.getResponseText().trim();
  if (!id) return;

  const statusResp = ui.alert(
    "Payment amount",
    "Click YES if paid in FULL.\n\n" +
    "Click NO if this is a PARTIAL payment (Hall's 50% advance only — Pool/Badminton must be paid in full).\n\n" +
    "Click CANCEL to abort.",
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (statusResp === ui.Button.CANCEL) return;
  const paymentStatus = statusResp === ui.Button.YES ? "Paid" : "Partial";

  const result = callAdminAction_("markPaymentReceived", { id: id, paymentStatus: paymentStatus });

  if (!result.success) {
    ui.alert("Could not update: " + result.error);
    return;
  }

  showResultWithWhatsApp_("Payment Recorded", result.message, result.customerPhone, result.whatsappText, result.customerEmailSent);
}

// ── Block / unblock a slot (maintenance, private use, etc.) ────────────

function blockSlot() {
  if (!requireRole_(["Admin", "Manager"])) return;

  const ui = SpreadsheetApp.getUi();

  const facilityResp = ui.prompt("Block a Slot (1 of 6)", "Facility name — AC Hall, Non-AC Hall, Lawn, Swimming Pool, or Badminton:", ui.ButtonSet.OK_CANCEL);
  if (facilityResp.getSelectedButton() !== ui.Button.OK) return;
  const facilityName = facilityResp.getResponseText().trim();
  if (!ALL_FACILITY_IDS[facilityName]) {
    ui.alert("Facility not recognized: " + facilityName + "\n\nPlease type it exactly as: AC Hall, Non-AC Hall, Lawn, Swimming Pool, or Badminton.");
    return;
  }

  let resourceId = "";
  if (facilityName === "Badminton") {
    const courtResp = ui.prompt(
      "Block a Slot (2 of 6)",
      "Which court? Type \"Court 1\", \"Court 2\", or leave blank to block BOTH courts:",
      ui.ButtonSet.OK_CANCEL
    );
    if (courtResp.getSelectedButton() !== ui.Button.OK) return;
    const courtText = courtResp.getResponseText().trim();
    if (courtText) {
      const match = BADMINTON_COURTS.find(function (c) { return c.name.toLowerCase() === courtText.toLowerCase(); });
      if (!match) {
        ui.alert("Court not recognized: " + courtText + "\n\nPlease type exactly \"Court 1\" or \"Court 2\", or leave blank for both.");
        return;
      }
      resourceId = match.id;
    }
  }

  const dateResp = ui.prompt("Block a Slot (3 of 6)", "Date (yyyy-mm-dd, e.g. 2026-08-25):", ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  const blockDate = dateResp.getResponseText().trim();

  const startResp = ui.prompt("Block a Slot (4 of 6)", "Start time, 24-hour HH:mm (e.g. 09:00):", ui.ButtonSet.OK_CANCEL);
  if (startResp.getSelectedButton() !== ui.Button.OK) return;
  const startTime = startResp.getResponseText().trim();

  const endResp = ui.prompt("Block a Slot (5 of 6)", "End time, 24-hour HH:mm (e.g. 11:00):", ui.ButtonSet.OK_CANCEL);
  if (endResp.getSelectedButton() !== ui.Button.OK) return;
  const endTime = endResp.getResponseText().trim();

  const reasonResp = ui.prompt("Block a Slot (6 of 6)", "Reason (e.g. Maintenance, Private event):", ui.ButtonSet.OK_CANCEL);
  if (reasonResp.getSelectedButton() !== ui.Button.OK) return;
  const reason = reasonResp.getResponseText().trim();

  const result = callAdminAction_("blockSlot", {
    facilityName: facilityName,
    resourceId: resourceId,
    blockDate: blockDate,
    startTime: startTime,
    endTime: endTime,
    reason: reason
  });
  ui.alert(result.success ? result.message : ("Could not block: " + result.error));
}

function unblockSlot() {
  if (!requireRole_(["Admin", "Manager"])) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Unblock a Slot", "Enter the Block ID to remove (e.g. BLK-1723...):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const blockId = response.getResponseText().trim();
  const result = callAdminAction_("unblockSlot", { blockId: blockId });
  ui.alert(result.success ? result.message : ("Could not unblock: " + result.error));
}

// ── Open / close a member-reserved Badminton slot for public booking ───
//
// Badminton hours 5–8 AM and 5–11 PM are reserved for members by
// default (see BADMINTON_RESERVED_WINDOWS). These tools let the admin
// open a specific date/time (optionally just one court) for public
// booking, and close it again later.
//
// REQUIRES a "Badminton Unblocks" sheet with these exact headers:
//   Unblock ID | Date | Start Time | End Time | Resource ID | Status | Notes

function openReservedBadmintonSlot() {
  if (!requireRole_(["Admin"])) return;

  const ui = SpreadsheetApp.getUi();

  const dateResp = ui.prompt("Open a Reserved Slot (1 of 4)", "Date (yyyy-mm-dd, e.g. 2026-08-25):", ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  const date = dateResp.getResponseText().trim();

  const startResp = ui.prompt("Open a Reserved Slot (2 of 4)", "Start time, 24-hour HH:mm (e.g. 05:00):", ui.ButtonSet.OK_CANCEL);
  if (startResp.getSelectedButton() !== ui.Button.OK) return;
  const startTime = startResp.getResponseText().trim();

  const endResp = ui.prompt("Open a Reserved Slot (3 of 4)", "End time, 24-hour HH:mm (e.g. 08:00):", ui.ButtonSet.OK_CANCEL);
  if (endResp.getSelectedButton() !== ui.Button.OK) return;
  const endTime = endResp.getResponseText().trim();

  const courtResp = ui.prompt("Open a Reserved Slot (4 of 4)", "Which court? Type \"Court 1\", \"Court 2\", or leave blank for BOTH:", ui.ButtonSet.OK_CANCEL);
  if (courtResp.getSelectedButton() !== ui.Button.OK) return;
  const courtText = courtResp.getResponseText().trim();

  let resourceId = "";
  let courtSummary = "Both courts";
  if (courtText) {
    const match = BADMINTON_COURTS.find(function (c) { return c.name.toLowerCase() === courtText.toLowerCase(); });
    if (!match) {
      ui.alert("Court not recognized: " + courtText + "\n\nPlease type exactly \"Court 1\" or \"Court 2\", or leave blank for both.");
      return;
    }
    resourceId = match.id;
    courtSummary = match.name;
  }

  if (!confirmAction_(
    "You are about to OPEN this member-reserved time for PUBLIC booking:\n\n" +
    "Date: " + date + "\n" +
    "Time: " + startTime + "–" + endTime + "\n" +
    "Court: " + courtSummary + "\n\n" +
    "Members may lose priority access to this slot. Continue?"
  )) {
    ui.alert("Cancelled — no changes were made.");
    return;
  }

  const result = callAdminAction_("openReservedBadmintonSlot", {
    date: date,
    startTime: startTime,
    endTime: endTime,
    resourceId: resourceId
  });
  ui.alert(result.success ? result.message : ("Could not open slot: " + result.error));
}

function closeReservedBadmintonSlot() {
  if (!requireRole_(["Admin"])) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Close a Reserved Slot", "Enter the Unblock ID to remove (e.g. UNB-1723...):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const unblockId = response.getResponseText().trim();

  // Look up details first (a READ, works fine regardless of edit
  // protection) so the confirmation can show what's actually being closed.
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Badminton Unblocks");
  if (!sheet) {
    ui.alert("Couldn't find a 'Badminton Unblocks' sheet.");
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert("No overrides found."); return; }

  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const rowIndex = rows.findIndex(function (row) { return String(row[0]).trim() === unblockId; });
  if (rowIndex === -1) {
    ui.alert("Unblock ID not found: " + unblockId);
    return;
  }

  const overrideRow = rows[rowIndex];
  const courtLabel = String(overrideRow[4]).trim()
    ? (BADMINTON_COURTS.find(function (c) { return c.id === String(overrideRow[4]).trim(); }) || {}).name || overrideRow[4]
    : "Both courts";

  if (!confirmAction_(
    "You are about to RE-LOCK this slot back to members-only:\n\n" +
    "Date: " + overrideRow[1] + "\n" +
    "Time: " + overrideRow[2] + "–" + overrideRow[3] + "\n" +
    "Court: " + courtLabel + "\n\n" +
    "Continue?"
  )) {
    ui.alert("Cancelled — no changes were made.");
    return;
  }

  const result = callAdminAction_("closeReservedBadmintonSlot", { unblockId: unblockId });
  ui.alert(result.success ? result.message : ("Could not close slot: " + result.error));
}

function setupStatusManagement() {
  if (!requireRole_(["Admin"])) return;

  if (!confirmAction_(
    "This will add a dropdown (New / Contacted / Follow-up / Converted / Lost) " +
    "and colour-coding to the Status column in Enquiries.\n\n" +
    "Existing data will not be changed.\n\nContinue?"
  )) return;

  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName("Enquiries");

  const statusValues = ["New", "Contacted", "Follow-up", "Converted", "Lost"];

  // Status is column J (10th column) based on your current headers:
  // Timestamp | Enquiry ID | Name | Mobile | Email | Facility | Date | Guests | Message | Status
  const statusColumn = 10;
  const lastRow = Math.max(sheet.getLastRow(), 1000); // covers existing + future rows

  const range = sheet.getRange(2, statusColumn, lastRow - 1, 1);

  // 1. Dropdown validation
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusValues, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);

  // 2. Colour-coded conditional formatting -- drop any rule this function
  // added on a previous run (same range) first, so re-running doesn't pile
  // up duplicate rules.
  const rangeA1 = range.getA1Notation();
  const rules = sheet.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (rg) { return rg.getA1Notation() === rangeA1; });
  });

  const colours = {
    "New": "#FFF2CC",       // pale yellow — needs attention
    "Contacted": "#D9E8FB", // pale blue — in progress
    "Follow-up": "#FCE4D6", // pale orange — needs action
    "Converted": "#D9EAD3", // pale green — success
    "Lost": "#F4CCCC"       // pale red — closed, no sale
  };

  Object.keys(colours).forEach(status => {
    const cfRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(colours[status])
      .setRanges([range])
      .build();
    rules.push(cfRule);
  });

  sheet.setConditionalFormatRules(rules);

  const message =
    "Done! Status column now has a dropdown (New / Contacted / Follow-up / " +
    "Converted / Lost) with colour coding. Existing 'New' rows are unaffected.";

  try {
    // Works only when triggered from the Sheet's UI (e.g. the custom menu).
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    // Running directly from the Apps Script editor has no UI context —
    // log instead so it doesn't throw.
    Logger.log(message);
  }
}

const REMINDER_THRESHOLD_HOURS = 12;
const OPEN_STATUSES = ["New", "Contacted", "Follow-up"]; // not Converted/Lost

function sendFollowUpReminders() {
  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName("Enquiries");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // no data rows yet

  // Columns: A Timestamp | B EnquiryID | C Name | D Mobile | E Email |
  //          F Facility | G Date | H Guests | I Message | J Status | K Source
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const now = new Date();
  const overdue = [];

  data.forEach(row => {
    const [timestamp, enquiryId, name, mobile, , facility, , , , status] = row;
    if (!timestamp || !status) return;
    if (OPEN_STATUSES.indexOf(status) === -1) return;

    const ageHours = (now - new Date(timestamp)) / (1000 * 60 * 60);
    if (ageHours >= REMINDER_THRESHOLD_HOURS) {
      overdue.push({
        enquiryId: enquiryId,
        name: name,
        mobile: mobile,
        facility: facility,
        status: status,
        ageHours: Math.floor(ageHours)
      });
    }
  });

  if (overdue.length === 0) return; // nothing to report — no email sent

  const lines = overdue.map(o =>
    `${o.enquiryId} — ${o.name} (${o.mobile}) — ${o.facility} — ` +
    `Status: ${o.status} — Waiting ${o.ageHours}h`
  );

  const subject = `L's Park: ${overdue.length} enquir${overdue.length === 1 ? "y needs" : "ies need"} follow-up`;
  const body =
    `The following enquiries have been open ${REMINDER_THRESHOLD_HOURS}+ hours ` +
    `without moving to Converted or Lost:\n\n` +
    lines.join("\n") +
    `\n\nOpen the Enquiries sheet to update their status:\n` +
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;

  MailApp.sendEmail(OWNER_NOTIFY_EMAIL, subject, body);
}

function installFollowUpTrigger() {
  if (!requireRole_(["Admin"])) return;

  if (!confirmAction_(
    "This will turn ON automatic follow-up reminder emails, checked every " +
    "12 hours, sent to " + OWNER_NOTIFY_EMAIL + ".\n\nContinue?"
  )) return;

  removeFollowUpTriggerSilently_(); // avoid duplicate triggers if run twice — no confirmation dialog, we already confirmed above

  ScriptApp.newTrigger("sendFollowUpReminders")
    .timeBased()
    .everyHours(12)
    .create();

  const message = "Reminder trigger installed — checking every 12 hours for enquiries open 12+ hours.";
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

// Menu-facing version — checks role and confirms before turning triggers off.
function removeFollowUpTrigger() {
  if (!requireRole_(["Admin"])) return;

  if (!confirmAction_(
    "This will turn OFF automatic follow-up reminder emails.\n\n" +
    "You can turn it back on anytime with \"Install 12-hr reminder trigger.\"\n\nContinue?"
  )) return;

  removeFollowUpTriggerSilently_();
  SpreadsheetApp.getUi().alert("Reminder trigger removed. No more automatic emails until reinstalled.");
}
// Internal helper — the actual deletion logic, no role check or
// confirmation. Called by both installFollowUpTrigger (to clear old
// triggers first) and removeFollowUpTrigger (after its own confirmation).
function removeFollowUpTriggerSilently_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendFollowUpReminders") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
