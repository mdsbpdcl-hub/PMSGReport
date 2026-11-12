/*
 * Settings for the PM Surya Ghar Report Maker web page.
 * Edit this file on GitHub (pencil icon), change the link, then "Commit changes".
 */
window.PMSG_CONFIG = {
  // The dashboard's web app link, between the quotes. Copy it from the Apps Script project:
  // Deploy > Manage deployments > Web app. It starts with https://script.google.com/ and ends with /exec.
  // (This is NOT the Google Sheet link.) Leave it empty to only make the Excel report.
  dashboardLink: "",

  // Change this number whenever you upload a new pmsg_report_app.py, so browsers fetch the new one.
  engineVersion: "3.0",
};
