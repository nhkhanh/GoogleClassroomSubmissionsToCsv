const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { google } = require('googleapis');
const { authenticate } = require('@google-cloud/local-auth');

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
];

const TOKEN_PATH = path.join(process.cwd(), 'token-sheets.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
const EXPORT_DIR = path.join(process.cwd(), 'export');

async function selectFolder() {
  if (process.argv[2]) {
    const dir = process.argv[2];
    if (!fs.existsSync(dir)) {
      console.error(`Error: Directory not found: ${dir}`);
      process.exit(1);
    }
    return dir;
  }

  if (!fs.existsSync(EXPORT_DIR)) {
    console.error(`Error: Export directory not found: ${EXPORT_DIR}`);
    process.exit(1);
  }

  const folders = fs.readdirSync(EXPORT_DIR, { withFileTypes: true })
    .filter(d => d.isDirectory())
    .map(d => d.name)
    .sort();

  if (folders.length === 0) {
    console.error('No subfolders found in export directory.');
    process.exit(1);
  }

  console.log('Select a folder to import:\n');
  folders.forEach((folder, i) => {
    console.log(`  ${i + 1}. ${folder}`);
  });
  console.log();

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const answer = await new Promise(resolve => {
    rl.question(`Enter number (1-${folders.length}): `, resolve);
  });
  rl.close();

  const index = parseInt(answer, 10) - 1;
  if (isNaN(index) || index < 0 || index >= folders.length) {
    console.error('Invalid selection.');
    process.exit(1);
  }

  return path.join(EXPORT_DIR, folders[index]);
}

let CSV_DIRECTORY;

async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.promises.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

async function deleteToken() {
  try {
    await fs.promises.unlink(TOKEN_PATH);
    console.log('Deleted invalid token, re-authenticating...');
  } catch (err) {
    // Token file doesn't exist, ignore
  }
}

async function saveCredentials(client) {
  const content = await fs.promises.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;

  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.promises.writeFile(TOKEN_PATH, payload);
}

async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    try {
      await client.getAccessToken();
      return client;
    } catch (err) {
      if (err.response?.data?.error === 'invalid_grant') {
        await deleteToken();
      } else {
        throw err;
      }
    }
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
    additionalOptions: {
      access_type: 'offline',
      prompt: 'consent',
    },
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

function parseCSV(content) {
  const lines = content.split('\n').filter(line => line.trim());
  if (lines.length === 0) return [];

  const result = [];
  let currentLine = '';
  let inQuotes = false;

  for (const line of lines) {
    currentLine += (currentLine ? '\n' : '') + line;
    const quoteCount = (currentLine.match(/"/g) || []).length;
    inQuotes = quoteCount % 2 !== 0;

    if (!inQuotes) {
      const row = [];
      let cell = '';
      let insideQuotes = false;

      for (let i = 0; i < currentLine.length; i++) {
        const char = currentLine[i];
        if (char === '"') {
          if (insideQuotes && currentLine[i + 1] === '"') {
            cell += '"';
            i++;
          } else {
            insideQuotes = !insideQuotes;
          }
        } else if (char === ',' && !insideQuotes) {
          row.push(cell);
          cell = '';
        } else {
          cell += char;
        }
      }
      row.push(cell);
      result.push(row);
      currentLine = '';
    }
  }

  return result;
}

async function importCsvsToSheets(auth) {
  const sheets = google.sheets({ version: 'v4', auth });

  // Get all CSV files
  const files = fs.readdirSync(CSV_DIRECTORY)
    .filter(file => file.endsWith('.csv'))
    .sort();

  if (files.length === 0) {
    console.log('No CSV files found in', CSV_DIRECTORY);
    return;
  }

  console.log(`Found ${files.length} CSV files to import`);

  // Create spreadsheet with sheets for each CSV
  const sheetRequests = files.map((file, index) => ({
    properties: {
      sheetId: index,
      title: path.basename(file, '.csv').substring(0, 100), // Sheet name max 100 chars
      index: index,
    },
  }));

  const createResponse = await sheets.spreadsheets.create({
    requestBody: {
      properties: {
        title: `${path.basename(CSV_DIRECTORY)} Submissions`,
      },
      sheets: sheetRequests,
    },
  });

  const spreadsheetId = createResponse.data.spreadsheetId;
  console.log(`Created spreadsheet: https://docs.google.com/spreadsheets/d/${spreadsheetId}`);

  // Import data to each sheet
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const sheetName = path.basename(file, '.csv').substring(0, 100);
    const filePath = path.join(CSV_DIRECTORY, file);

    console.log(`Importing: ${file}`);

    const content = fs.readFileSync(filePath, 'utf-8');
    const data = parseCSV(content);

    if (data.length === 0) {
      console.log(`  Skipped (empty file)`);
      continue;
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${sheetName}'!A1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: data,
      },
    });

    console.log(`  Imported ${data.length} rows`);
  }

  console.log('\nDone! Spreadsheet URL:');
  console.log(`https://docs.google.com/spreadsheets/d/${spreadsheetId}`);
}

selectFolder()
  .then(folder => {
    CSV_DIRECTORY = folder;
    console.log(`\nImporting from: ${CSV_DIRECTORY}\n`);
    return authorize();
  })
  .then(importCsvsToSheets)
  .catch(console.error);
