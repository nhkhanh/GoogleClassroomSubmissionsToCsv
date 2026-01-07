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
const path = require('path');
const { google } = require('googleapis');
const {authenticate} = require('@google-cloud/local-auth');
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
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.promises.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * Deletes the saved token file.
 *
 * @return {Promise<void>}
 */
async function deleteToken() {
  try {
    await fs.promises.unlink(TOKEN_PATH);
    console.log('Deleted invalid token.json, re-authenticating...');
  } catch (err) {
    // Token file doesn't exist, ignore
  }
}

/**
 * Serializes credentials to a file compatible with GoogleAuth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
  const content = await fs.promises.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;

  // Unable to get refresh token? Visit https://stackoverflow.com/a/10857806/2353894
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.promises.writeFile(TOKEN_PATH, payload);
}

/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    // Test if the token is still valid
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
    if (err) {
      console.error('Error listing courses:', err.message);
      return;
    }

    console.log('Courses list:');
    console.table(res.data.courses.map(course => ({ id: course.id, name: course.name })));
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question('Enter the index that you want to export course\'s submission list: ', (index) => {
      rl.close();
      const {id: courseId, name: courseName} = res.data.courses[index];
      const courseFolder = `${EXPORT_DIRECTORY}/${courseName}`;
      // Create course folder
      if (!fs.existsSync(courseFolder))
        fs.mkdirSync(courseFolder);
      classroom.courses.courseWork.list({ courseId }, (err, res) => {
        if (err) {
          console.log(err);
          return;
        }
        res.data.courseWork.forEach((courseWork => {
          const deadline = courseWork.dueDate && new Date(courseWork.dueDate.year,
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
            .filter(s => s.assignmentSubmission && s.assignmentSubmission.attachments)
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
                deadline: deadline && toExcelDate(deadline),
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
            const fileName = `${courseFolder}/${courseWork.title}.csv`;

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

authorize().then(exportToCsvs).catch(console.error);
