# Kids Budget Google Web App

This is a Google Apps Script web app with two user roles:

- **Kid**: can log in and see current budget plus operation history.
- **Parent**: can log in, see any kid account, and add incoming/outgoing operations.

The app uses **Google Sheets** as the database.

## Google Sheet setup

Use a spreadsheet with two tabs:

### 1) `Users`
Header row:

`username | password | role | childId | displayName`

Example rows:

- `kid_anna | 1234 | kid | KID001 | Anna`
- `dad_joe | 1234 | parent |  | Joe`

### 2) `Operations`
Header row:

`timestamp | childId | amount | type | description | createdBy`

Notes:
- Incoming operations use a positive amount.
- Outgoing operations use a negative amount.

## Files

- `Code.gs` — backend logic (auth, sessions, sheet reads/writes).
- `Index.html` — responsive login + dashboard UI.
- `appsscript.json` — Apps Script manifest.

## Deploy

1. Create a new Apps Script project attached to your Google Sheet.
2. Copy files from this repo into the Apps Script project.
3. Deploy as web app.
4. Share/deploy URL to users.

## Security note

This starter keeps passwords as plain text in the `Users` sheet for simplicity. For production, store password hashes and improve auth/session hardening.
