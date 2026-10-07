# Covering Letter Generator (GitHub Pages + Firebase)

## Setup
1. **Create Firebase project:** console.firebase.google.com → Add project.
2. **Create Firestore:** Build → Firestore Database → Create database (production mode, nearest region).
3. **Register web app:** Project settings → Your apps → `</>` Web → register.
4. **Add config:** paste the `firebaseConfig` values into `firebase-config.js` (this is the only place).
5. **Collections:** nothing to create by hand. Both `dispatches` and `mappings` are filled from the app: Drawer → Upload Dispatch File → login → upload Dispatch.xlsx, then upload the Mapping.xlsx (columns `GP_EN, GP_HI, PS_EN, PS_HI, DIST_EN, DIST_HI`).
6. **Rules** (Firestore → Rules), replace ADMIN_EMAIL:
```
rules_version = '2';
service cloud.firestore { match /databases/{db}/documents {
  match /mappings/{id}   { allow read: if true;
    allow create, update: if request.auth != null && request.auth.token.email == 'ADMIN_EMAIL';
    allow delete: if false; }
  match /dispatches/{id} { allow read: if true;
    allow create, update: if request.auth != null && request.auth.token.email == 'ADMIN_EMAIL';
    allow delete: if false; }
}}
```
7. **Admin user:** Authentication → Get started → Email/Password → Users → Add user. Use that email in the rules above.
   Authentication → Settings → Authorized domains: add `USERNAME.github.io`.
8. **Word template:** save your official template as `template/covering-letter-template.docx` using `{{YEAR}} {{DISTRICT}} {{PS}} {{GP}} {{DISPATCH_NO}} {{DISPATCH_DATE}} {{PARA_NO}}`. Type each placeholder in one go (Word can split text into pieces if you edit mid-word; retype if you get a template error).
9. **GitHub:** create repo, upload all files (keep folder structure).
10. **Pages:** Settings → Pages → Deploy from branch → `main` / root.
11. **Test:** open `https://USERNAME.github.io/REPOSITORY/`, choose Year/District/PS/GP, generate.
12. **New Dispatch.xlsx:** Drawer → Upload Dispatch File → login → choose the year → select file → Preview → Upload / Update.
13. **Safe updates:** documents get ID `Year_UnitID` and are merged (never deleted), so re-uploading updates in place. Only the admin email can write.

## Notes
- Defaults: `{{DISPATCH_NO}}` = *Audit Party No*, `{{DISPATCH_DATE}}` = *Planned Start Date*, `{{PARA_NO}}` = *Converted to Para*. Change in the `F` object at the top of `app.js`.
- "All Gram Panchayats" renders the template once per GP and joins them with page breaks.
- Test locally with a web server (e.g. VS Code Live Server); ES modules don't run from `file://`.
