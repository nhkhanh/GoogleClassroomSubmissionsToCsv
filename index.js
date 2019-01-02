/**
 * @license
 * Copyright Google Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
  // [START classroom_quickstart]
const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const Json2csvParser = require('json2csv').Parser;

const COURSE_ID = 16576321386;
const EXPORT_DIRECTORY = 'export';
// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/classroom.courses.readonly',
  'https://www.googleapis.com/auth/classroom.coursework.me.readonly',
  'https://www.googleapis.com/auth/classroom.coursework.students.readonly',
  'https://www.googleapis.com/auth/classroom.coursework.students',
  'https://www.googleapis.com/auth/classroom.coursework.me',
];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Google Classroom API.
  authorize(JSON.parse(content), exportToCsvs);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error retrieving access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

/**
 * Lists the first 10 courses the user has access to.
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function exportToCsvs(auth) {
  if (!fs.existsSync(EXPORT_DIRECTORY)){
    fs.mkdirSync(EXPORT_DIRECTORY);
  }
  const classroom = google.classroom({ version: 'v1', auth });
  classroom.courses.courseWork.list({ courseId: COURSE_ID }, (err, res) => {
    if (err) {
      console.log(err);
      return;
    }
    res.data.courseWork.forEach((courseWork => {
      classroom.courses.courseWork.studentSubmissions.list({
          courseId: COURSE_ID,
          courseWorkId: courseWork.id,
        },
        (err, res) => {
          // console.log(res.data.studentSubmissions);
          const prune = res.data.studentSubmissions
            .filter(s => s.assignmentSubmission.attachments)
            .map(submission => {
              return {
                updateTime: submission.updateTime, state: submission.state, late: submission.late,
                files: submission.assignmentSubmission.attachments && submission.assignmentSubmission.attachments.reduce(
                  (acc, current) => acc += (current.driveFile ? (current.driveFile.title + ',') : ''),
                  ''),
              };
            });
          const parser = new Json2csvParser();
          const csv = parser.parse(prune);
          const fileName = `./${EXPORT_DIRECTORY}/${courseWork.title}.csv`;

          fs.writeFile(fileName, csv, err => {
            if (err)
              console.log(err);
            else
              console.log(`${fileName} created!`);
          });
        });
    }));
  });
}
