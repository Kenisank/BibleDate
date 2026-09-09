# Connecting the registration site to a Google Sheet

This takes about 5 minutes and needs no coding beyond copy-paste.

## 1. Create the Sheet
1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet.
2. Name it something like "Bible Date Registrations".

## 2. Add the backend script
1. In the Sheet, click **Extensions → Apps Script**.
2. Delete anything in the editor and paste in the full contents of `apps-script-backend.gs` (included alongside this guide).
3. Near the top of the script, edit these lines with your real event details:
   ```
   const EVENT_VENUE = "PASTE THE VENUE ADDRESS HERE";
   ```
   (EVENT_NAME and EVENT_DATE are already filled in — change them too if anything changes.)
4. Click the **Save** icon (or Ctrl/Cmd+S).

## 3. Deploy it as a web app
1. Click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy**.
5. Google will ask you to authorize the script — click through and allow it (it's your own script, so this is safe). It will ask for permission to send email on your behalf (from your Gmail address) — this is expected, since that's how registrants get their emails.
6. Copy the **Web app URL** it gives you (looks like `https://script.google.com/macros/s/XXXXXXX/exec`).

## 4. Connect the site
1. Open `bible-date-registration.html` in a text editor.
2. Find this line near the top of the `<script>` section:
   ```
   const SHEET_API_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
   ```
3. Replace the placeholder text with the URL you copied.
4. Save the file.

## 5. Deploy the site
Upload `bible-date-registration.html` to your host the same way you'd deploy any static site (e.g. Vercel — drag-and-drop or connect a repo).

## How the emails work
- **On registration:** the person immediately gets a "pending verification" email.
- **On verification:** open the "Organizer view" on the site (passcode-protected), find the person's row, and tick the **Verified** checkbox. This generates an access code and emails the person their access code, venue, date and time. It's one-way — once ticked it locks, so it won't re-send.
- **On food purchase:** tick **Item purchased from Crave & Crumb** for that person once their order is sorted. This does two things at once: it opens a pre-filled WhatsApp message to Amanda (Crave & Crumb) with the guest's name, items, and price so you can send the order straight to her — and it emails the guest a confirmation of what was arranged. Also one-way and locks after ticking. (Your browser may ask permission to open a pop-up the first time — allow it.)
- Every registration and its status is visible directly in the Google Sheet at any time, in addition to the site's Organizer view.

## One thing to know
Sending email from Apps Script (`MailApp`) uses your Google account's daily sending quota — around 100 emails a day on a personal Gmail account, higher on Google Workspace. For an event this size that should be more than enough, but worth knowing if registrations are very high.

## If you redeploy
If you ever choose **New deployment** again (rather than editing the existing one via "Manage deployments"), you'll get a new URL and need to update `SHEET_API_URL` in the HTML file again.
