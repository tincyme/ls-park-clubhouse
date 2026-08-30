#!/usr/bin/env python3
"""
PeedsPark Club House -- Automated Test Suite
==============================================

Runs the sanity check + functional tests against the site's HTML files
directly (no live server needed), with the Google Apps Script backend
mocked out -- so this NEVER touches your real Google Sheet, sends a
real WhatsApp message, or needs internet access. Safe to run as often
as you like, including right before a checkin.

WHAT THIS DOES NOT COVER (by design):
  - Section D (owner tools: Approve/Reject/Cancel/Block/etc.) lives in
    the Google Sheet menu, not the website -- there's nothing here for
    Playwright to click. Test those by hand in the Sheet.
  - Rate limiting (max 3 submissions / 30 min) and the Sheet
    permissions check need the real backend -- also manual.
  - This checks what the WEBSITE does with whatever the backend tells
    it. It does not verify the backend's own logic (e.g. that Apps
    Script actually computes Pool capacity correctly) -- that needs
    real bookings in the real Sheet to check by hand occasionally.

Full manual checklist (including the above) is at the published
Artifact -- ask Claude for the link if you don't have it.

------------------------------------------------------------------
SETUP (one-time, in Command Prompt):

    cd F:\\Github\\ls-park-clubhouse
    pip install playwright
    playwright install chromium firefox webkit

RUN (any time after that) -- runs against Chromium (covers Chrome/Edge),
Firefox, AND WebKit (covers Safari) by default, one after another, so
one run tells you if a booking goes through the same way across all
three real browser engines:

    cd F:\\Github\\ls-park-clubhouse
    python tests\\run_tests.py

To test just one engine (faster while iterating):
    python tests\\run_tests.py --browsers chromium

Exits with code 0 if EVERY browser's run passed, 1 if anything failed
in any browser -- so it also works as a pre-checkin gate in a script
if you want:
    python tests\\run_tests.py && git push ...
------------------------------------------------------------------
"""

import argparse
import json
import os
import sys
import datetime

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright isn't installed yet. Run:\n")
    print("    pip install playwright")
    print("    playwright install chromium firefox webkit\n")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Test bookkeeping
# ---------------------------------------------------------------------------

RESULTS = []  # list of dicts: browser, section, id, desc, status, detail
CURRENT_BROWSER = "chromium"  # updated by main() as it loops over --browsers


def record(section, tid, desc, status, detail=""):
    RESULTS.append({"browser": CURRENT_BROWSER, "section": section, "id": tid, "desc": desc, "status": status, "detail": detail})
    mark = {"PASS": "[PASS]", "FAIL": "[FAIL]", "SKIP": "[SKIP]"}[status]
    line = f"  {mark} [{CURRENT_BROWSER}] {tid:<5} {desc}"
    print(line)
    if status == "FAIL" and detail:
        print(f"           -> {detail}")


def check(section, tid, desc, condition, detail_if_fail=""):
    if condition:
        record(section, tid, desc, "PASS")
    else:
        record(section, tid, desc, "FAIL", detail_if_fail)


def section_header(title):
    print()
    print(title)
    print("-" * len(title))


# ---------------------------------------------------------------------------
# Playwright helpers
# ---------------------------------------------------------------------------

def collect_console_errors(page):
    """Tracks real JS exceptions (pageerror) -- the kind of bug that broke
    the enquiry form once before (an unresolved merge conflict shipped to
    production). Resource-load failures (e.g. a font blocked by a firewall,
    or Google Fonts unreachable offline) are NOT counted as failures --
    they're a network/environment detail, not a website bug -- but they're
    still returned separately so you can see them if you're curious."""
    errors = []
    resource_warnings = []
    page.on("pageerror", lambda e: errors.append(str(e)))

    def on_console(m):
        if m.type != "error":
            return
        if m.text.startswith("Failed to load resource"):
            resource_warnings.append(m.text)
        else:
            errors.append(m.text)

    page.on("console", on_console)
    return errors


def mock_availability(page, payload):
    def handler(route):
        if "action=availability" in route.request.url:
            route.fulfill(status=200, content_type="application/json", body=json.dumps(payload))
        else:
            route.continue_()
    page.route("**/macros/s/**", handler)


def mock_post(page, payload, capture=None):
    """Mocks any POST to the backend with the given JSON payload.
    If capture is a list, the parsed request body is appended to it."""
    def handler(route):
        if route.request.method == "POST":
            if capture is not None:
                try:
                    capture.append(json.loads(route.request.post_data or "{}"))
                except Exception:
                    capture.append(None)
            route.fulfill(status=200, content_type="application/json", body=json.dumps(payload))
        else:
            route.continue_()
    page.route("**/macros/s/**", handler)


def capture_next_alert(page):
    """Registers a one-shot dialog handler; returns a list that will
    contain the alert's message text after it fires (auto-accepted)."""
    box = []
    def on_dialog(dialog):
        box.append(dialog.message)
        dialog.accept()
    page.on("dialog", on_dialog)
    return box


def today_str():
    return datetime.date.today().strftime("%Y-%m-%d")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

FACILITY_PAGES = ["ac-hall.html", "non-ac-hall.html", "lawn.html", "pool.html", "badminton.html"]
ALL_PAGES = ["index.html"] + FACILITY_PAGES + ["privacy.html"]


def main():
    parser = argparse.ArgumentParser(description="PeedsPark automated test suite")
    parser.add_argument("--dir", default=None,
                         help="Folder containing index.html etc. Defaults to the repo root "
                              "(the parent of this script's folder).")
    parser.add_argument("--headed", action="store_true", help="Show the browser window while testing.")
    parser.add_argument("--browsers", default="chromium,firefox,webkit",
                         help="Comma-separated list of engines to run against: chromium (Chrome/Edge), "
                              "firefox, webkit (Safari). Default runs all three. Example: --browsers chromium")
    args = parser.parse_args()

    requested_browsers = [b.strip().lower() for b in args.browsers.split(",") if b.strip()]
    valid_browsers = {"chromium", "firefox", "webkit"}
    unknown = [b for b in requested_browsers if b not in valid_browsers]
    if unknown:
        print(f"Unknown browser(s): {', '.join(unknown)} -- choose from chromium, firefox, webkit")
        sys.exit(1)
    if not requested_browsers:
        print("No browsers specified.")
        sys.exit(1)

    site_dir = args.dir or os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    site_dir = os.path.abspath(site_dir)

    if not os.path.isfile(os.path.join(site_dir, "index.html")):
        print(f"Couldn't find index.html in: {site_dir}")
        print("Pass the folder that contains your site's HTML files with --dir, e.g.")
        print(r'    python tests\run_tests.py --dir F:\Github\ls-park-clubhouse')
        sys.exit(1)

    base_url = "file:///" + site_dir.replace("\\", "/").lstrip("/")

    print("PeedsPark Club House -- Automated Test Suite")
    print(f"Testing site files in: {site_dir}")
    print(f"Run at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")

    global CURRENT_BROWSER

    print(f"Browsers this run: {', '.join(requested_browsers)}")

    with sync_playwright() as p:
        engines = {"chromium": p.chromium, "firefox": p.firefox, "webkit": p.webkit}

        for browser_name in requested_browsers:
            CURRENT_BROWSER = browser_name
            launch_kwargs = {"headless": not args.headed}

            # Only Chromium supports pointing at a specific local executable
            # (e.g. a portable Chrome install) -- Firefox/WebKit always use
            # Playwright's own installed browser.
            if browser_name == "chromium":
                exe = os.environ.get("PLAYWRIGHT_CHROMIUM_PATH")
                if exe:
                    launch_kwargs["executable_path"] = exe

            print()
            print("=" * 60)
            print(f"RUNNING UNDER: {browser_name.upper()}")
            print("=" * 60)

            try:
                browser = engines[browser_name].launch(**launch_kwargs)
            except Exception as e:
                print(f"Could not launch {browser_name}: {e}")
                print(f"If this is the first time using {browser_name}, run: playwright install {browser_name}")
                record("SETUP", f"{browser_name}-launch", f"Launch {browser_name}", "FAIL", str(e))
                continue

            run_sanity(browser, base_url)
            run_enquiry_tests(browser, base_url)
            run_availability_tests(browser, base_url)
            run_booking_tests(browser, base_url)
            run_regression_tests(browser, base_url)

            browser.close()

    print_summary()


# ---------------------------------------------------------------------------
# SANITY -- mirrors the 8-step manual sanity check
# ---------------------------------------------------------------------------

def run_sanity(browser, base_url):
    section_header("SANITY CHECK (run this block after every change)")
    S = "SANITY"

    # S1/S2/S3 -- every page loads, has a title, no console errors
    for f in ALL_PAGES:
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        errors = collect_console_errors(page)
        try:
            page.goto(f"{base_url}/{f}")
            page.wait_for_timeout(300)
            title_ok = "peedspark" in page.title().lower()
            check(S, f"S1:{f}", f"{f} loads with a PeedsPark title", title_ok,
                  f"Page title was: {page.title()!r}")
            check(S, f"S3:{f}", f"{f} has no console errors on load", len(errors) == 0,
                  "; ".join(errors[:3]))
        except Exception as e:
            record(S, f"S1:{f}", f"{f} loads", "FAIL", str(e))
        page.close()

    # S4 -- images load (no broken <img>)
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    page.goto(f"{base_url}/index.html")
    page.wait_for_timeout(400)
    broken = page.eval_on_selector_all(
        "img", "els => els.filter(e => e.complete && e.naturalWidth === 0).map(e => e.src)"
    )
    check(S, "S4", "Homepage images all load (no broken <img>)", len(broken) == 0,
          f"Broken: {broken}")
    page.close()

    # S5 -- past-date guard present on every date field
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    page.goto(f"{base_url}/index.html")
    page.wait_for_timeout(200)
    home_min = page.get_attribute("#date", "min")
    check(S, "S5:home", "Homepage 'Preferred date' rejects past dates (min attr = today)",
          home_min == today_str(), f"min was {home_min!r}, expected {today_str()!r}")
    page.close()

    for f in FACILITY_PAGES:
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        page.goto(f"{base_url}/{f}")
        page.wait_for_timeout(200)
        avail_min = page.get_attribute("#availDate", "min")
        check(S, f"S5:{f}", f"{f} 'Check Availability' date rejects past dates (min attr = today)",
              avail_min == today_str(), f"min was {avail_min!r}, expected {today_str()!r}")
        page.close()

    # S6 -- availability check renders a result
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    mock_availability(page, {"success": True, "type": "fixed", "slots": {"Morning": "Available", "Evening": "Booked"}})
    page.goto(f"{base_url}/ac-hall.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    results_visible = page.is_visible("#availResults")
    check(S, "S6", "Check Availability returns a rendered result", results_visible)
    page.close()

    # S7 -- quick enquiry mocked submit shows confirmation
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    mock_post(page, {"success": True, "enquiryId": "ENQ-SANITY-1"})
    page.goto(f"{base_url}/index.html")
    page.fill("#name", "Sanity Test")
    page.fill("#phone", "9999999999")
    page.select_option("#facility", index=1)
    page.wait_for_timeout(3200)  # anti-bot minimum-fill-time
    page.click("#enquiryForm button[type=submit]")
    page.wait_for_timeout(500)
    check(S, "S7", "Quick Enquiry submit shows the confirmation panel",
          page.is_visible("#confirmationPanel"))
    page.close()

    # S8 -- mobile viewport, no horizontal overflow
    page = browser.new_page(viewport={"width": 375, "height": 800})
    page.goto(f"{base_url}/index.html")
    page.wait_for_timeout(300)
    overflow = page.evaluate("document.documentElement.scrollWidth - window.innerWidth")
    check(S, "S8", "Homepage has no horizontal overflow at 375px wide", overflow <= 2,
          f"scrollWidth exceeds viewport by {overflow}px")
    page.close()


# ---------------------------------------------------------------------------
# A. Enquiry system
# ---------------------------------------------------------------------------

def run_enquiry_tests(browser, base_url):
    section_header("A. ENQUIRY SYSTEM (homepage Quick Enquiry)")
    S = "A"

    # A1 -- valid submission
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    sent = []
    mock_post(page, {"success": True, "enquiryId": "ENQ-A1"}, capture=sent)
    page.goto(f"{base_url}/index.html")
    page.fill("#name", "A1 Tester")
    page.fill("#phone", "9876543210")
    page.select_option("#facility", index=1)
    page.wait_for_timeout(3200)
    page.click("#enquiryForm button[type=submit]")
    page.wait_for_timeout(500)
    conf_ok = page.is_visible("#confirmationPanel")
    payload_ok = bool(sent) and sent[0] and sent[0].get("name") == "A1 Tester" and sent[0].get("phone") == "9876543210"
    check(S, "A1", "Valid enquiry: confirmation shown, correct data sent to backend",
          conf_ok and payload_ok, f"confirmation={conf_ok}, payload={sent}")
    page.close()

    # A2 -- backend rejects an invalid phone; UI surfaces the error, no confirmation
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    alert_box = capture_next_alert(page)
    mock_post(page, {"success": False, "error": "Please enter a valid 10-digit mobile number."})
    page.goto(f"{base_url}/index.html")
    page.fill("#name", "A2 Tester")
    page.fill("#phone", "12345")
    page.select_option("#facility", index=1)
    page.wait_for_timeout(3200)
    page.click("#enquiryForm button[type=submit]")
    page.wait_for_timeout(500)
    error_shown = any("10-digit" in m for m in alert_box)
    check(S, "A2", "Backend phone-validation error is shown to the user, no confirmation",
          error_shown and not page.is_visible("#confirmationPanel"),
          f"alert(s) seen: {alert_box}")
    page.close()


# ---------------------------------------------------------------------------
# B. Availability check
# ---------------------------------------------------------------------------

def run_availability_tests(browser, base_url):
    section_header("B. AVAILABILITY CHECK")
    S = "B"

    # B1/B2 -- fixed-slot facility: Available shows a button, Booked doesn't
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    mock_availability(page, {"success": True, "type": "fixed", "slots": {"Morning": "Available", "Evening": "Booked"}})
    page.goto(f"{base_url}/ac-hall.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    slots_text = page.text_content("#availSlots") or ""
    buttons = page.locator("#availSlots button:has-text('Request to Book')").count()
    check(S, "B1", "Available slot shows a 'Request to Book' button", buttons == 1,
          f"found {buttons} buttons, text was: {slots_text}")
    check(S, "B2", "Booked slot has no button and shows 'Booked'", "Booked" in slots_text,
          f"text was: {slots_text}")
    page.close()

    # B3 -- Pending status renders, no button for that slot
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    mock_availability(page, {"success": True, "type": "fixed", "slots": {"Morning": "Pending", "Evening": "Available"}})
    page.goto(f"{base_url}/ac-hall.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    slots_text = page.text_content("#availSlots") or ""
    buttons = page.locator("#availSlots button:has-text('Request to Book')").count()
    check(S, "B3", "Pending slot shows 'Pending' with no button (Evening still bookable)",
          "Pending" in slots_text and buttons == 1,
          f"buttons={buttons}, text was: {slots_text}")
    page.close()

    # B4 -- Blocked status renders, no button
    page = browser.new_page(viewport={"width": 1400, "height": 1000})
    mock_availability(page, {"success": True, "type": "fixed", "slots": {"Morning": "Blocked", "Evening": "Available"}})
    page.goto(f"{base_url}/lawn.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    slots_text = page.text_content("#availSlots") or ""
    check(S, "B4", "Blocked slot shows 'Blocked'", "Blocked" in slots_text, f"text was: {slots_text}")
    page.close()

    # B5 -- date-change auto-refresh, and no stale-date booking
    page = browser.new_page(viewport={"width": 1400, "height": 1200})

    def by_date_handler(route):
        url = route.request.url
        if "action=availability" in url:
            if "date=2026-09-28" in url:
                payload = {"success": True, "type": "fixed", "slots": {"Morning": "Available", "Evening": "Booked"}}
            elif "date=2026-09-29" in url:
                payload = {"success": True, "type": "fixed", "slots": {"Morning": "Booked", "Evening": "Available"}}
            else:
                payload = {"success": True, "type": "fixed", "slots": {"Morning": "Available", "Evening": "Available"}}
            route.fulfill(status=200, content_type="application/json", body=json.dumps(payload))
        else:
            route.continue_()
    page.route("**/macros/s/**", by_date_handler)

    page.goto(f"{base_url}/ac-hall.html")
    page.fill("#availDate", "2026-09-28")
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    page.fill("#availDate", "2026-09-29")
    page.wait_for_timeout(500)  # auto-refresh listener, no button click
    try:
        page.click("#availSlots button:has-text('Request to Book')", timeout=2000)
        slot_label = page.text_content("#bookingSlotLabel") or ""
        b5_ok = "2026-09-29" in slot_label and "Evening" in slot_label
        detail = f"booking form opened for: {slot_label}"
    except Exception as e:
        b5_ok = False
        detail = str(e)
    check(S, "B5", "Changing date auto-refreshes; booking uses the NEW date, not a stale one",
          b5_ok, detail)
    page.close()

    # B6 -- Badminton: two courts behave independently
    page = browser.new_page(viewport={"width": 1400, "height": 1200})
    mock_availability(page, {
        "success": True, "type": "hourly", "bookingModel": "resource",
        "slots": [{"start": "18:00", "end": "19:00", "courts": [
            {"id": "R001", "name": "Court 1", "status": "Available"},
            {"id": "R002", "name": "Court 2", "status": "Booked"},
        ]}],
    })
    page.goto(f"{base_url}/badminton.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    slots_text = page.text_content("#availSlots") or ""
    buttons = page.locator("#availSlots button:has-text('Request to Book')").count()
    check(S, "B6", "Badminton courts tracked independently (Court 1 bookable, Court 2 not)",
          "Court 1" in slots_text and "Court 2" in slots_text and buttons == 1,
          f"buttons={buttons}, text was: {slots_text}")
    page.close()


# ---------------------------------------------------------------------------
# C. Booking request flow
# ---------------------------------------------------------------------------

def run_booking_tests(browser, base_url):
    section_header("C. BOOKING REQUEST FLOW")
    S = "C"

    # C1/C2 -- form opens correct, valid submit hides the button afterward
    #
    # Note: filling the date field itself triggers an availability refresh
    # (the B5 fix), same as clicking Check Availability does -- so the mock
    # below keys off whether the booking has actually been SUBMITTED yet,
    # not off how many times availability happened to be queried.
    page = browser.new_page(viewport={"width": 1400, "height": 1200})
    state = {"submitted": False}

    def handler(route):
        if route.request.method == "POST":
            state["submitted"] = True
            route.fulfill(status=200, content_type="application/json",
                           body=json.dumps({"success": True, "requestId": "REQ-C2"}))
            return
        if "action=availability" in route.request.url:
            if not state["submitted"]:
                # Only Morning is open, so the click below is unambiguous.
                payload = {"success": True, "type": "fixed", "slots": {"Morning": "Available", "Evening": "Booked"}}
            else:
                # Post-submit refresh: Morning is now Pending (just requested),
                # Evening is still open -- exactly one button should remain.
                payload = {"success": True, "type": "fixed", "slots": {"Morning": "Pending", "Evening": "Available"}}
            route.fulfill(status=200, content_type="application/json", body=json.dumps(payload))
        else:
            route.continue_()
    page.route("**/macros/s/**", handler)

    page.goto(f"{base_url}/ac-hall.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(300)
    page.click("#availSlots button:has-text('Request to Book')")
    page.wait_for_timeout(200)
    slot_label = page.text_content("#bookingSlotLabel") or ""
    check(S, "C1", "Booking form opens for the currently-shown facility/date/slot",
          "Morning" in slot_label, f"label was: {slot_label}")

    page.fill("#bkName", "C2 Tester")
    page.fill("#bkPhone", "9000090000")
    page.wait_for_timeout(3200)
    page.click("#bookingSubmitBtn")
    page.wait_for_timeout(600)
    confirm_ok = page.is_visible("#bookingConfirmationPanel")
    buttons_left = page.locator("#availSlots button:has-text('Request to Book')").count()
    check(S, "C2", "Successful submit shows confirmation and hides that slot's button",
          confirm_ok and buttons_left == 1,  # only Evening's button should remain
          f"confirm={confirm_ok}, buttons remaining={buttons_left}")
    page.close()

    # C3 -- Pool: requesting more guests than remain is blocked BEFORE
    # submission (the form disables Send Booking Request live, as you
    # type -- it never even reaches the server for this case).
    page = browser.new_page(viewport={"width": 1400, "height": 1200})
    mock_availability(page, {
        "success": True, "type": "hourly", "bookingModel": "capacity",
        "slots": [{"start": "10:00", "end": "11:00", "status": "Available", "remaining": 3, "capacity": 8}],
    })
    page.goto(f"{base_url}/pool.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    page.click("#availSlots button:has-text('Request to Book')")
    page.wait_for_timeout(200)
    page.fill("#bkGuests", "6")  # only 3 remain
    page.wait_for_timeout(150)
    submit_disabled = page.is_disabled("#bookingSubmitBtn")
    note_text = page.text_content("#bkDurationNote") or ""
    check(S, "C3", "Pool: over-capacity guest count disables Send Booking Request with a clear reason",
          submit_disabled and "3 spot" in note_text,
          f"disabled={submit_disabled}, note={note_text!r}")
    page.close()

    # C5 (folded) -- hourly facilities show a duration selector, fixed-slot facilities don't
    page = browser.new_page(viewport={"width": 1400, "height": 1200})
    mock_availability(page, {
        "success": True, "type": "hourly", "bookingModel": "capacity",
        "slots": [{"start": "10:00", "end": "11:00", "status": "Available", "remaining": 5, "capacity": 8}],
    })
    page.goto(f"{base_url}/pool.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    page.click("#availSlots button:has-text('Request to Book')")
    page.wait_for_timeout(200)
    pool_duration_visible = page.is_visible("#bkDurationWrap")
    page.close()

    page = browser.new_page(viewport={"width": 1400, "height": 1200})
    mock_availability(page, {"success": True, "type": "fixed", "slots": {"Morning": "Available", "Evening": "Available"}})
    page.goto(f"{base_url}/ac-hall.html")
    page.fill("#availDate", today_str())
    page.click("#checkAvailBtn")
    page.wait_for_timeout(400)
    page.click("#availSlots button:has-text('Request to Book')")
    page.wait_for_timeout(200)
    hall_duration_visible = page.is_visible("#bkDurationWrap")
    check(S, "C5", "Duration selector shows for hourly facilities (Pool), hidden for fixed-slot (AC Hall)",
          pool_duration_visible and not hall_duration_visible,
          f"pool={pool_duration_visible}, ac-hall={hall_duration_visible}")
    page.close()


# ---------------------------------------------------------------------------
# F. Cross-page regression
# ---------------------------------------------------------------------------

def run_regression_tests(browser, base_url):
    section_header("F. CROSS-PAGE REGRESSION")
    S = "F"

    active_ok = True
    active_detail = {}
    for f in FACILITY_PAGES:
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        page.goto(f"{base_url}/{f}")
        page.wait_for_timeout(200)
        active = page.eval_on_selector_all("nav a.active", "els => els.map(e => e.textContent.trim())")
        active_detail[f] = active
        if len(active) != 1:
            active_ok = False
        page.close()
    check(S, "F1", "Exactly one nav tab is active on each facility page", active_ok, str(active_detail))

    cta_texts = {}
    cta_ok = True
    for f in ALL_PAGES:
        if f == "privacy.html":
            continue
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        page.goto(f"{base_url}/{f}")
        page.wait_for_timeout(200)
        cta = page.text_content(".nav-cta")
        cta_texts[f] = (cta or "").strip()
        if (cta or "").strip() != "Enquire Now":
            cta_ok = False
        page.close()
    check(S, "F2", "Header CTA reads 'Enquire Now' consistently on every page", cta_ok, str(cta_texts))

    footer_ok = True
    footer_detail = {}
    for f in ALL_PAGES:
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        page.goto(f"{base_url}/{f}")
        page.wait_for_timeout(200)
        has_link = page.locator("a[href='privacy.html']").count() > 0
        footer_detail[f] = has_link
        if not has_link:
            footer_ok = False
        page.close()
    check(S, "F3", "Every page's footer links to privacy.html", footer_ok, str(footer_detail))

    overflow_ok = True
    overflow_detail = {}
    for f in ALL_PAGES:
        page = browser.new_page(viewport={"width": 375, "height": 800})
        page.goto(f"{base_url}/{f}")
        page.wait_for_timeout(300)
        overflow = page.evaluate("document.documentElement.scrollWidth - window.innerWidth")
        overflow_detail[f] = overflow
        if overflow > 2:
            overflow_ok = False
        page.close()
    check(S, "F4", "No horizontal overflow on any page at 375px wide", overflow_ok, str(overflow_detail))

    conflict_ok = True
    conflict_detail = {}
    for f in ALL_PAGES:
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        page.goto(f"{base_url}/{f}")
        content = page.content()
        markers = [m for m in ("<<<<<<<", "=======", ">>>>>>>") if m in content]
        conflict_detail[f] = markers
        if markers:
            conflict_ok = False
        page.close()
    check(S, "F5", "No unresolved Git merge-conflict markers in any page's source", conflict_ok, str(conflict_detail))


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

def print_summary():
    passed = sum(1 for r in RESULTS if r["status"] == "PASS")
    failed = [r for r in RESULTS if r["status"] == "FAIL"]
    total = len(RESULTS)
    browsers_run = sorted(set(r["browser"] for r in RESULTS))

    print()
    print("=" * 60)
    print(f"OVERALL RESULT: {passed}/{total} passed across {len(browsers_run)} browser(s): {', '.join(browsers_run)}")

    if len(browsers_run) > 1:
        print()
        print("Per-browser breakdown:")
        for b in browsers_run:
            b_results = [r for r in RESULTS if r["browser"] == b]
            b_passed = sum(1 for r in b_results if r["status"] == "PASS")
            b_failed = [r for r in b_results if r["status"] == "FAIL"]
            status_word = "ALL PASSED" if not b_failed else f"{len(b_failed)} FAILED"
            print(f"  {b:<10} {b_passed}/{len(b_results)} passed -- {status_word}")

    if failed:
        print(f"\n{len(failed)} FAILED (across all browsers):")
        for r in failed:
            print(f"  - [{r['browser']}] [{r['section']}] {r['id']}: {r['desc']}")
            if r["detail"]:
                print(f"      {r['detail']}")
        # Flag anything that failed in one browser but not another -- the
        # interesting cross-browser signal ("booking works in Chrome but
        # not Safari") rather than a bug that's broken everywhere.
        by_test = {}
        for r in RESULTS:
            key = (r["section"], r["id"])
            by_test.setdefault(key, {})[r["browser"]] = r["status"]
        inconsistent = {k: v for k, v in by_test.items() if len(set(v.values())) > 1}
        if inconsistent:
            print(f"\n{len(inconsistent)} test(s) behaved DIFFERENTLY across browsers (the real cross-browser signal):")
            for (section, tid), per_browser in inconsistent.items():
                detail = ", ".join(f"{b}={status}" for b, status in sorted(per_browser.items()))
                print(f"  - [{section}] {tid}: {detail}")
    print()
    print("Not covered by this script (needs manual checking):")
    print("  D. Owner tools (Approve/Reject/Cancel/Block) -- Google Sheet menu")
    print("  A4. Rate limiting (3 submissions / 30 min)   -- needs the real backend")
    print("  E3. Google Sheet sharing permissions          -- check in Google Drive")
    print("  Whether the Apps Script Web App itself is redeployed with the latest code")
    print("=" * 60)

    sys.exit(0 if not failed else 1)


if __name__ == "__main__":
    main()
