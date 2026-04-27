# BPils

Internal staffing tool for generating and updating monthly work plans from a guest list source.

## What it does

1. Connects to a Google Sheet or loads an Excel source workbook
2. Detects month tabs such as `apr26`, `mai26`, `sep26`
3. Generates 3 Excel plan types:
   - `Viesnīca`
   - `Restorāns`
   - `Virtuve`
4. Lets the manager fill in the downloaded Excel plan
5. Writes staffing updates back into the original Google Sheet or Excel workbook

## Stack

- React
- Vite
- ExcelJS
- Google Sheets API

## Environment

Create the Vercel environment variable:

- `VITE_GOOGLE_CLIENT_ID`

This must be a Google OAuth Web Client ID with your deployed site added to `Authorized JavaScript origins`.
