# React + JavaScript + Spreadsheet Demo

This is a simple React app that reads and writes an Excel spreadsheet **in the browser**, without a Node.js backend.

## Important limitation
Browsers cannot silently open and edit an arbitrary local `.xlsx` file. This demo uses the browser's File System Access API where supported (Chrome/Edge):
1. Click **Connect spreadsheet** and select an existing `.xlsx` file, or create/select a spreadsheet file.
2. The app reads the `Contacts` sheet.
3. Add/edit/delete contacts in the UI.
4. Click **Save to spreadsheet** to write the updated rows back to the selected file.

For browsers without File System Access API support, the app can import an Excel file and download an updated copy, but it cannot overwrite the original automatically.

## Run it
Requires Node.js only to run the Vite development server; there is **no Node.js backend**.

```bash
npm install
npm run dev
```

Open the local URL shown by Vite, usually http://localhost:5173.

## Spreadsheet format
Use an Excel workbook with a sheet named `Contacts` and columns:
- `id` (optional; generated if missing)
- `name`
- `email`
- `phone`

If the sheet is empty or absent, the app starts with sample contacts and can save them to the workbook.

## Files
- `src/App.jsx` — UI and spreadsheet operations
- `src/main.jsx` — React entry point
- `src/style.css` — styling
- `index.html` — app HTML
