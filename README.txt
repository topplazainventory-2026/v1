TOP PLAZA STOCK MANAGEMENT SYSTEM
GitHub source + Google Apps Script hosting + Google Sheets database

IMPORTANT ARCHITECTURE
- GitHub stores your source-code backup and change history.
- Google Apps Script runs the web application and serves index.html.
- Your existing Google Sheets remains the database.
- No GitHub Pages, Cloudflare Worker, API bridge, or database migration is used.

FILES
- Code.gs: Google Apps Script backend. Keep this in the Apps Script project as Code.gs.
- index.html: Frontend HTML. In the Apps Script editor, create an HTML file named exactly: index
  because Code.gs calls createHtmlOutputFromFile('index').

SETUP / SAFE WORKFLOW
1. Open your existing Google Apps Script project connected to your Google Sheet.
2. Back up the current Apps Script project and spreadsheet before making changes.
3. In Apps Script, confirm the HTML file is named exactly "index" (lowercase).
4. These are the original files you uploaded. They have not been rewritten for an external API.
5. If restoring these files, replace the contents of the corresponding Apps Script files with
   these source files. Do not create a second database or change spreadsheet IDs/sheet names.
6. Save the Apps Script project.
7. Deploy > Manage deployments > Edit the existing Web app deployment, or create a new Web app
   deployment if needed. Choose execution and access settings appropriate for your users.
8. Open the Apps Script Web app URL to use the system. The live page is served by Apps Script,
   not GitHub Pages.
9. In GitHub, create/open your repository and upload Code.gs, index.html, and README.txt.
   GitHub is your version-controlled source copy. When code changes in GitHub, copy the updated
   file contents into the matching Apps Script file, save, and update the Web app deployment.

IMPORTANT NOTES
- Do not publish this app as a GitHub Pages website: the page uses google.script.run, which is
  provided by the Apps Script HTML-service runtime and is not available on ordinary GitHub Pages.
- Do not upload passwords, private keys, OAuth tokens, or spreadsheet credentials to GitHub.
- The app relies on the existing spreadsheet and sheet names configured in Code.gs.
- The provided Code.gs calls createHtmlOutputFromFile('index'), so the Apps Script HTML file
  must be named "index" in lowercase.
- The supplied Code.gs does not define getImageFromCell(), although the HTML references it.
  This is an existing source-code mismatch, not a GitHub hosting issue. The logo feature may need
  its own follow-up fix; this note does not change the rest of your code.
