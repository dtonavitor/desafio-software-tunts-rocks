const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
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
        const content = await fs.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        return google.auth.fromJSON(credentials);
    } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
    });
    await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
    return client;
    }
        client = await authenticate({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
    });
    if (client.credentials) {
        await saveCredentials(client);
    }
    return client;
}

/**
 * Get information from the spreadsheet, such as "Faltas", "P1", "P2" and "P3"
 * @see https://docs.google.com/spreadsheets/d/1NNPU9egmEDJnytlAscxKQsUt4Q0kq4YtoNWIvkWF8rQ/edit?usp=sharing
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
async function readSheetInformation(auth) {
    const sheets = google.sheets({version: 'v4', auth});
    const res = await sheets.spreadsheets.values.get({
        spreadsheetId: '1NNPU9egmEDJnytlAscxKQsUt4Q0kq4YtoNWIvkWF8rQ',
        range: 'engenharia_de_software',
    });
    const rows = res.data.values;

    if (!rows || rows.length === 0) {
        console.log('No data found.');
        return;
    }

    let statusValues = [];
    let finalApprovalGrades = [];

    // This gives a dynamic number of classes, so we can change the number of classes in the spreadsheet and it will still work 
    let numberOfClasses = rows[1][0].split(": ")[1]

    console.log('Faltas, P1, P2, P3:');
    rows.forEach((row, index) => {
        // These rows are the headers, so we skip them.
        if (index === 0 || index === 1 || index === 2 ) {
            return;
        }
        // Print columns C to F, which correspond to indices 2 to 5.
        console.log(`${row[2]}, ${row[3]}, ${row[4]}, ${row[5]}`);

        let statusValue = []
        let finalApprovalGrade = []

        if (Math.ceil((row[2] * 100) / numberOfClasses) > 25) {
            statusValue.push("Reprovado por Falta")
        } else {
            // diving by 30 because the max grade is 10 (make the grades in the 0-10 interval) and there are 3 grades
            let meanGrade = Math.ceil((parseInt(row[3]) + parseInt(row[4]) + parseInt(row[5])) / 30)
            if (meanGrade < 5) {
                statusValue.push("Reprovado por Nota")
                finalApprovalGrade.push(0)
            } else if (meanGrade >= 5 && meanGrade < 7) {
                statusValue.push("Exame Final")

                let grade = 10 - meanGrade
                finalApprovalGrade.push(grade)
            } else {
                statusValue.push("Aprovado")
                finalApprovalGrade.push(0)
            }
        }

        statusValues.push(statusValue)
        finalApprovalGrades.push(finalApprovalGrade)
    });

    await updateStatus(auth, statusValues);
    await updateFinalApproval(auth, finalApprovalGrades);
}

/**
 * Completes the "Situação" column in the spreadsheet
 * @see https://docs.google.com/spreadsheets/d/1NNPU9egmEDJnytlAscxKQsUt4Q0kq4YtoNWIvkWF8rQ/edit?usp=sharing
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 * @param {Array} statusValues The values to be inserted for the "Situação" column
 */
async function updateStatus(auth, statusValues) {
    const sheets = google.sheets({version: 'v4', auth});

    // The range is dynamic, so we can add more students to the spreadsheet and it will still work
    const res = await sheets.spreadsheets.values.update({
        spreadsheetId: '1NNPU9egmEDJnytlAscxKQsUt4Q0kq4YtoNWIvkWF8rQ',
        range: `G4:G${statusValues.length + 3}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
            values: statusValues,
        },
    });
    console.log(res.data);
}

/**
 * Completes the "Nota para Aprovação Final" column in the spreadsheet
 * @see https://docs.google.com/spreadsheets/d/1NNPU9egmEDJnytlAscxKQsUt4Q0kq4YtoNWIvkWF8rQ/edit?usp=sharing
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 * @param {Array} finalGrades The values to be inserted for the "Nota para Aprovação Final" column
 */
async function updateFinalApproval(auth, finalGrades) {
    const sheets = google.sheets({version: 'v4', auth});

    // The range is dynamic, so we can add more students to the spreadsheet and it will still work
    const res = await sheets.spreadsheets.values.update({
      spreadsheetId: '1NNPU9egmEDJnytlAscxKQsUt4Q0kq4YtoNWIvkWF8rQ',
      range: 'H4:H' + (finalGrades.length + 3),
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: finalGrades,
      },
    });
    console.log(res.data);
  }

authorize().then(readSheetInformation).catch(console.error);