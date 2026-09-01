# PeedsPark — Staging Pipeline Setup

**Risk note, since this is your first time running this kind of pipeline:** the whole point of staging is to stop a half-tested change from touching real customer bookings or the live site. The single biggest way this protection breaks is a copy-paste mistake between the two backends below (pointing staging at the *production* Apps Script/Sheet by accident, or vice versa). Every step that involves a URL is flagged **warning** - go slowly on those.

## Architecture

```
PRODUCTION (unchanged, real customers)
  peedspark.com  ->  GitHub Pages (main branch)  ->  PROD Apps Script  ->  PROD Google Sheet

STAGING (new)
  <name>.netlify.app  ->  Netlify (any branch/PR)  ->  STAGING Apps Script  ->  STAGING Google Sheet (a copy)
```

Nothing about production changes. Netlify is a second, independent home for test deploys -- GitHub Pages keeps serving `main` to peedspark.com exactly as it does today.

## One-time setup

### 1. Duplicate the Google Sheet
In Google Sheets, open the live `LS_Park_Availability_Booking_v1` workbook -> **File -> Make a copy** -> name it something obviously not-production, e.g. `PeedsPark_STAGING_Availability_Booking`. This becomes the database that staging tests write to, so a bad test booking can never show up in your real Bookings sheet.

### 2. Deploy a separate Apps Script Web App for the copy
1. In the **copy**, open **Extensions -> Apps Script**.
2. Paste in the same `scripts/code.gs` and `scripts/roles.gs` content you use in production.
3. **WARNING - critical:** in `code.gs`, find the line:
   ```
   const WEB_APP_URL = "https://script.google.com/macros/s/.../exec";
   ```
   Leave it as-is for now -- you'll come back and paste the *staging* deployment's own URL into this exact line after step 4, once you have it. Admin actions (Confirm/Cancel/Block/Mark Paid) call this URL to reach themselves, so if this ever points at the *production* URL, staging's admin actions would silently write to real bookings. Do not skip step 5.
4. **Deploy -> New deployment -> Web app.** Execute as **Me**, Who has access: **Anyone**. Deploy, and copy the `/exec` URL it gives you -- this is your **staging Apps Script URL**.
5. Go back into the script, paste that same URL into the `WEB_APP_URL` constant from step 3, save, and **Deploy -> Manage deployments -> pencil icon -> Version: New version -> Deploy** again (editing code alone never updates the live URL -- this project has hit that exact gotcha before).

### 3. Point the website's config at the staging backend
In this repo, open `js/config.js` and replace the placeholder:
```js
var STAGING_API_URL = "PASTE_STAGING_APPS_SCRIPT_WEB_APP_URL_HERE";
```
with the `/exec` URL from step 2.4. Commit this change. (The `PROD_API_URL` line above it is already correct and doesn't need touching -- leave it alone.)

How the switch works: `js/config.js` checks the domain the page is loaded from. `peedspark.com`, `www.peedspark.com`, and `tincyme.github.io` use the production backend; every other domain (Netlify preview URLs, `file://` when you open a page locally, `localhost`) uses staging. You never have to remember to flip anything by hand -- the same HTML/JS files behave correctly on both.

### 4. Connect the repo to Netlify
1. Sign up at netlify.com (free tier is enough), "Sign up with GitHub."
2. **Add new site -> Import an existing project -> GitHub -> `tincyme/ls-park-clubhouse`.**
3. Build settings: leave the build command **blank**, publish directory `.`. (There's a `netlify.toml` already in the repo with these same settings, so Netlify should pick them up automatically.)
4. Deploy. Netlify gives you a URL like `random-name-123.netlify.app` -- that's your permanent staging URL. Optional: **Site settings -> Domain management -> Options -> Edit site name** to rename it to something memorable, e.g. `peedspark-staging.netlify.app`.
5. **Site settings -> Build & deploy -> Deploy contexts**: confirm "Deploy previews" is on for pull requests (it's on by default). This means every PR you open also gets its own one-off preview URL, separate from the main staging URL, without any extra setup.

That's the whole one-time setup. Nothing here touches GitHub Pages, the CNAME, or the production Apps Script/Sheet.

## Day-to-day workflow

1. **Branch as you already do** -- a feature branch off `main` for whatever you're building.
2. **Open the PR against a `staging` branch** (create this branch once, from current `main`, if it doesn't exist yet) rather than straight into `main`. Netlify posts a deploy-preview link on the PR automatically.
3. **Test on that preview URL**, not just locally:
   - Run the relevant section(s) of the manual Test Plan (`claude/test-plan-link.md`) against the preview URL.
   - Run the automated script (`tests/run_tests.py`) -- its mocked-backend run stays your fast first gate.
   - For anything that writes to the Sheet (a real booking, a status change), click through it for real on the preview URL at least once -- this hits the **staging** Sheet only, so it's safe, and it's the only way to catch real-Google-Sheets quirks a mocked test can't see (this project has already hit two of those: Sheets silently converting date/time text, and a permissions-mode bug in `onOpen()` -- both were real-backend-only bugs).
   - Per your standing 27 Aug practice: decide whether this feature needs a new Test Plan case or automated check *before* merging, not after.
4. **Merge the feature branch into `staging`** once it's green. This updates the shared staging URL for anyone else checking it, but never touches `main` or peedspark.com.
5. **When you're confident in what's sitting on `staging`, merge `staging` into `main`.** That's the one action that goes live to real customers -- GitHub Pages picks it up automatically, same as it always has.

## What this does not cover

- **Netlify account creation, the Apps Script deploy, and the Google Sheet copy are manual, account-bound steps** -- no session can do these on your behalf, they need your own Google/Netlify logins.
- This staging setup is for the **live GitHub Pages / Apps Script / Google Sheets site** only. The separate Supabase rewrite in progress (`chota1business/PeedsParkClubhouse`) has its own environment story (Supabase database branching) and isn't part of this.
- A pre-existing item worth flagging while we're in this code: `scripts/code.gs` is committed into this **public** GitHub repository (required for free GitHub Pages) with `ADMIN_ACTION_SECRET` written in plain text -- meaning that secret is technically visible to anyone who looks at the repo, not just people with Editor access to the Apps Script project. That's a separate, pre-existing risk from before this session and outside today's task, but you may want to revisit it (e.g. moving secrets out of the committed file) when you have time.
