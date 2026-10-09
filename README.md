# React Contact Manager

A React app for managing contacts in the browser. Contacts are stored in IndexedDB, so they remain available in the same browser without downloading an updated spreadsheet after each change.

## Features

- Add, edit, and delete contacts with changes saved to the browser database.
- Load contacts automatically when the app opens again in the same browser.
- Import an Excel workbook once, then save its contacts to the browser database.
- Search contacts by name, email, or phone number.

The database is local to each browser and device. It does not sync between users or devices. The Excel import accepts a workbook with a `Contacts` sheet and `id`, `name`, `email`, and `phone` columns. If the sheet is absent, the first worksheet is used.

## Run locally

```bash
npm install
npm run dev
```

## Deploy

Push to the `main` branch. GitHub Actions builds the app and publishes the `dist` folder to GitHub Pages.
