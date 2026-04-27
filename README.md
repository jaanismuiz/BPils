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

## Deploy to Vercel

1. Open Vercel and import the GitHub repository `jaanismuiz/BPils`
2. Keep the default Vite build settings
3. Add the environment variable `VITE_GOOGLE_CLIENT_ID`
4. Deploy the project

## Google OAuth setup

In Google Cloud:

1. Open the OAuth Web Client used by this app
2. Add your deployed Vercel URL to `Authorized JavaScript origins`
3. If you later add a custom domain, add that domain there too

Example:

- `https://your-project.vercel.app`

Without this step, Google Sheets login will not work on the deployed site.
