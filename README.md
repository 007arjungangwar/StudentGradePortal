# Student Grade Portal

A dynamic student grade portal that uses Google Sheets as the database, Google Apps Script as the API, and a static `index.html` frontend for the student login/dashboard.

## Architecture

- Google Sheet: stores student rows and any teacher-managed columns.
- Google Apps Script: validates PRN + `MASTER_PASSWORD`, detects columns dynamically, and returns the matching row as JSON.
- Frontend: logs in with PRN/password, stores the session in `sessionStorage`, and generates dashboard cards from whatever columns the sheet returns.

## Setup

1. Open `Code.gs` in this repo and copy it into your Apps Script project.
2. Set `MASTER_PASSWORD` in `Code.gs`.
3. Confirm `SHEET_ID` points to the correct Google Sheet.
4. Deploy Apps Script as a Web App:
   - Execute as: `Me`
   - Who has access: `Anyone`
5. Copy the Web App URL into `SCRIPT_URL` in `index.html`.
6. Open the frontend.

## Live Preview

https://raw.githack.com/007arjungangwar/StudentGradePortal/main/index.html

## GitHub Pages

To use the official GitHub Pages URL, open repository `Settings` -> `Pages`, set `Source` to `Deploy from a branch`, select `gh-pages` and `/root`, then save.

Expected URL after publishing:

https://007arjungangwar.github.io/StudentGradePortal/

## Sheet Rules

- The first row must contain column headers.
- One header must contain `PRN`, for example `PRN`, `Prn Number`, or `Student PRN`.
- Teachers can add, remove, or rename any other columns without frontend code changes.
- Empty cells are returned as `Pending`.
