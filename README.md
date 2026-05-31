# Student Grade Portal

A static student grade portal connected to the Google Apps Script web app for live result lookup.

## Live Preview

Open the portal here:

https://raw.githack.com/007arjungangwar/StudentGradePortal/main/index.html

## Official GitHub Pages Setup

To use the official GitHub Pages URL, enable Pages in the repository settings:

1. Open `Settings` in this repository.
2. Open `Pages`.
3. Set `Source` to `Deploy from a branch`.
4. Select branch `gh-pages` and folder `/root`.
5. Save.

After GitHub finishes publishing, the official URL should be:

https://007arjungangwar.github.io/StudentGradePortal/

## Apps Script Note

The portal sends login data as `text/plain` JSON to avoid Apps Script CORS preflight issues. The Apps Script backend should read `e.postData.contents` and return JSON like:

```json
{
  "success": true,
  "data": {
    "name": "Student Name",
    "prn": "123",
    "assignment": "18",
    "midsem": "24",
    "EndSem": "45",
    "total": "87",
    "attendance %": "92"
  }
}
```
