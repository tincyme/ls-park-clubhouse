// PeedsPark environment config
// ------------------------------------------------------------------
// Picks which backend (Google Apps Script Web App + Google Sheet) the
// site talks to, based on the hostname the page is loaded from. This
// is what lets the exact same HTML/JS files run safely against a
// staging Sheet on a staging URL, and only ever touch the real
// customer data when served from the real production domains.
//
// Edit the two URLs below whenever you deploy a new Apps Script
// version (Deploy > Manage deployments) for either environment.
// See STAGING_SETUP.md for how to set up the staging Sheet + Script.
(function () {
  var PROD_HOSTS = ["peedspark.com", "www.peedspark.com", "tincyme.github.io"];

  // Same Apps Script Web App URL that was hardcoded in every page before.
  var PROD_API_URL =
    "https://script.google.com/macros/s/AKfycbwSGAt5wWrnMFWZ_vO8InNrLH1rd3vvmSpjS_k9evtJQfhEkeyPp5LrnREX0_866JnlCw/exec";

  // Fill this in once you've made a staging copy of the Google Sheet
  // and deployed a separate Apps Script Web App against it (Step 1-2
  // of STAGING_SETUP.md). Until then, staging/local pages will call
  // this placeholder and every request will simply fail loudly
  // instead of silently touching production data.
  var STAGING_API_URL = "PASTE_STAGING_APPS_SCRIPT_WEB_APP_URL_HERE";

  var isProd = PROD_HOSTS.indexOf(window.location.hostname) !== -1;

  window.PEEDSPARK_CONFIG = {
    ENQUIRY_API: isProd ? PROD_API_URL : STAGING_API_URL,
    ENV: isProd ? "production" : "staging"
  };
})();
