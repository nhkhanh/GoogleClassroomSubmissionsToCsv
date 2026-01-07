## Setup

See https://developers.google.com/classroom/quickstart/nodejs to create credentials.json.

You also need to enable the following APIs in Google Cloud Console:
- Google Classroom API
- Google Sheets API

## Usage

### Export Google Classroom submissions to CSV

```bash
npm start
```

This will list your courses and let you select one to export all coursework submissions as CSV files.

### Import CSV files to Google Sheets

```bash
node import-to-sheets.js <csv-directory>
```

Example:
```bash
node import-to-sheets.js ./export/2510-AWAD-22KTPM2
```

This will create a new Google Sheets spreadsheet with each CSV file as a separate sheet.
