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
async function exportToCsvs(auth) {
  if (!fs.existsSync(EXPORT_DIRECTORY)) {
    fs.mkdirSync(EXPORT_DIRECTORY);
  }
  const classroom = google.classroom({ version: 'v1', auth });
  classroom.courses.list((err, res) => {
    console.log('Courses list:');
    console.table(res.data.courses.map(course => ({ id: course.id, name: course.name })));
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question('Enter course id: ', (courseId) => {
      rl.close();
      classroom.courses.courseWork.list({ courseId }, (err, res) => {
        if (err) {
          console.log(err);
          return;
        }
        res.data.courseWork.forEach((courseWork => {
          const deadline = new Date(courseWork.dueDate.year,
          courseWork.dueDate.month - 1,
          courseWork.dueDate.day,
          courseWork.dueTime.hours - new Date().getTimezoneOffset() / 60,
          courseWork.dueTime.minutes || 0);
          classroom.courses.courseWork.studentSubmissions.list({
            courseId,
            courseWorkId: courseWork.id,
          },
          (err, res) => {
            if (err) {
              console.log(err);
              return;
            }
            if (!res)
              return;
            let rowIndex = 1;
            const prune = res.data.studentSubmissions
            .filter(s => s.assignmentSubmission.attachments)
            .map(submission => {
              const updateTime = new Date(submission.updateTime);
              const files = submission.assignmentSubmission.attachments && submission.assignmentSubmission.attachments.reduce(
              (acc, current) => acc += (current.driveFile ? (current.driveFile.title + ',') : ''),
              '');
              const matchStudentId = files && files.match(/\d{6,10}/g);
              const studentId = matchStudentId && matchStudentId[0];
              rowIndex++;
              return {
                updateTime: toExcelDate(updateTime),
                state: submission.state,
                late: submission.late,
                files,
                deadline: toExcelDate(deadline),
                lateHours: `=(A${rowIndex} - E${rowIndex}) * 60`,
                studentId,
                penalty: `=IF(F${rowIndex} <= 0, 1, IF(F${rowIndex} <= 2, 0.8, IF(F${rowIndex} <= 24, 0.5, 0)))`,
                finalGrade: `=H${rowIndex} * J${rowIndex}`,
                actualGrade: '',
                studentId2: matchStudentId && matchStudentId[1],
                studentId3: matchStudentId && matchStudentId[2],
              };
            });
            if (!prune || !prune.length) {
              console.log('Ignored no submission course work: ', courseWork.title, '');
              return;
            }
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
    });
  });
}

function addHoursToDate(date, hours) {
  return new Date(new Date(date).setHours(date.getHours() + hours));
}


function toExcelDate(date) {
  return addHoursToDate(date, -new Date().getTimezoneOffset() / 60)
  .toISOString()
  .replace('T', ' ')
  .replace('Z', '');
}
