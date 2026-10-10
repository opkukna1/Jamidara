# Cover Letter Generator (GitHub Pages, no Firebase)

A static site. No login, no database. Your data stays in your browser.

## Files
- `index.html`, `app.js`, `style.css`, `js/` : the app
- `template/covering-letter-template.docx` : Word template for the cover letter
- `mapping.json` : Hindi / English names of every District, Panchayat Samiti and Gram Panchayat (you create it once, see below)

## Deploy
1. Create a GitHub repo and upload all files (keep the folder structure).
2. Settings -> Pages -> Deploy from branch -> `main` / root.
3. Open `https://USERNAME.github.io/REPOSITORY/`.

## mapping.json (one time)
1. Open the site -> menu -> **Data Files** -> section 2, choose your mapping Excel
   (columns `GP_EN, GP_HI, PS_EN, PS_HI, DIST_EN, DIST_HI`).
2. Press **Download mapping.json**.
3. In GitHub: Add file -> Upload files -> put `mapping.json` next to `index.html` -> Commit.
Repeat only when new Panchayats are added.

## Every day
1. Menu -> **Data Files** -> choose today's Dispatch .xlsx (the Year is detected from "Audit Party No").
2. Press **Use this file**. It is stored in this browser; adding a file for the same Year replaces the old one.
3. Menu -> **Reports** -> pick Cover Letter or AMS Report, then Download or Share on WhatsApp.

## AMS Report
Tick Panchayat Samitis, press **Save Watch List** (saved in this browser), then **Generate AMS Report**.

## Notes
- Each phone / browser needs the Dispatch file added once (it is not shared between devices).
- Clearing browser data removes the stored Dispatch file; just add it again.
- Word placeholders used by the cover letter template: `{{YEAR}} {{DISPATCH_NO}} {{DATE}} {{PS_NAME_HI}} {{DISTRICT_HI}} {{GP_NAME_HI}} {{PARA_COUNT}} {{PARA_BREAKUP}}`.
- `main.py` / `requirements.txt` are an optional old backend and are not needed for the site.
