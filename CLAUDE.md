# Claude Code Notes

## Project Overview

This project exports Google Classroom submissions to CSV files and can import them into Google Sheets.

## Scripts

- `index.js` - Main script to export Google Classroom submissions to CSV
- `import-to-sheets.js` - Import CSV files to Google Sheets

## Authentication

Uses OAuth2 with separate tokens:
- `token.json` - For Google Classroom API
- `token-sheets.json` - For Google Sheets API

Both require `credentials.json` from Google Cloud Console.

## Required APIs

Enable in Google Cloud Console:
- Google Classroom API
- Google Sheets API
