# Student Grade Portal

A dynamic student grade portal that uses Google Sheets as the database, Google Apps Script as the API, and a static `index.html` frontend for the student login/dashboard.

## Architecture

- Google Sheet: each sheet tab is one subject, for example `Programming in Python (Section 1)`.
- Google Apps Script: generates per-student passwords, emails unsent students, validates `Subject + PRN/Application ID + Password`, detects columns dynamically, and returns the matching row as JSON.
- Frontend: loads subject names and PRN/Application ID options from sheet tabs, logs in with subject plus PRN or Application ID and password, stores the session in `sessionStorage`, and generates dashboard cards from whatever columns the sheet returns.

## Setup

1. Open `Code.gs` in this repo and copy it into your Apps Script project.
2. Confirm `SHEET_ID` points to the correct Google Sheet.
3. Deploy Apps Script as a Web App:
   - Execute as: `Me`
   - Who has access: `Anyone`
4. Copy the Web App URL into `SCRIPT_URL` in `index.html`.
5. Run `sendStudentPasswords()` in Apps Script.
6. Approve the Google permissions for Sheets and Gmail/Mail when prompted.
7. Open the frontend.

## Sending Student Passwords

Run this Apps Script function manually:

```js
sendStudentPasswords()
```

It will process every sheet tab as a subject:

- Adds `Password` just after the PRN column if it is missing.
- Adds `Email_send` after `Password` if it is missing.
- Generates a password only when the password cell is empty.
- Sends email only when `Email_send` is not `Send`.
- Treats rows with a blank `Password` as unsent, even if `Email_send` was accidentally copied as `Send`.
- Logs each sent row, PRN, and recipient email in Apps Script execution output.
- Marks successful emails as `Send`.
- Leaves missing/failed emails as `unsend`, so rerunning only retries those rows.

## Live Preview

https://007arjungangwar.github.io/StudentGradePortal/

## GitHub Pages

To use the official GitHub Pages URL, open repository `Settings` -> `Pages`, set `Source` to `Deploy from a branch`, select `gh-pages` and `/root`, then save.

Expected URL after publishing:

https://007arjungangwar.github.io/StudentGradePortal/

## Sheet Rules

- The first row must contain column headers.
- At least one header must contain `PRN` or an application identifier.
- PRN headers can be named `PRN`, `Prn Number`, or `Student PRN`.
- Application ID headers can be named `Application`, `Application ID`, `Application No`, or `Application Number`.
- The login page automatically loads PRN/Application ID choices after a subject is selected.
- One header must contain `Email`, for example `Email`, `Email ID`, or `Student Email`.
- Every sheet tab name becomes a subject option on the login page.
- Teachers can add, remove, or rename any other columns without frontend code changes.
- Empty cells are returned as `Pending`.
- The dashboard does not show internal `Password` or `Email_send` columns.
