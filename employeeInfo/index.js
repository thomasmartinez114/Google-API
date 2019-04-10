const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Sheets API.
    authorize(JSON.parse(content), listMajors);
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
            if (err) return console.error('Error while trying to retrieve access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https: //docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
function listMajors(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    sheets.spreadsheets.values.get({
        spreadsheetId: '17_dcw6jUCjZqT-K3WwRwU6EcohDVIhCJ0uheMdu_blk',
        range: 'employeeData!A2:E',
    }, (err, res) => {
        if (err) return console.log('The API returned an error: ' + err);
        const rows = res.data.values;


        if (rows.length) {

            // console.log('--------------------');
            // console.log('The rows data');
            // console.log(rows[0][0], rows[0][1]);
            // console.log('--------------------');
            // console.log('First Name, Last Name, Employee ID, Department, Salary');
            console.log('Generating HTML pages for each employee . . . .');


            // A [0]: First Name    // B [1]: Last Name // C [2]: Employee ID   // D [3]: Department    // E [4]: Salary

            rows.map((row) => {

                let employeeData = `
                <!DOCTYPE html>
                    <html>
                    <head>
                    <meta charset='utf-8'>
                    <title>${row[0]}_${row[2]}</title>
                    <link rel='stylesheet' href='styles.css'>
                    </head>
                        <body>
                        <div class='title'>
                            <h1>Employee Information</h1>
                        </div>
                        
                    <div class='wrapper'>
                            <div class='wrapper-content'>
                            
                            <div class='employee-photo'>
                                <img src='../images/${row[0]}.png'>
                            </div>
                            
                                <div class='employee-name'>
                                    <p>Name: ${row[0]} ${row[1]}</p>
                                </div>
                                
                                <div class='employee-id'>
                                    <p>ID Number: ${row[2]}</p>
                                </div>
                                
                                <div class='employee-dept'>
                                    <p>Department: ${row[3]}</p>
                                </div>
                                
                                <div class='employee-salary'>
                                    <p>Salary: $${row[4]}</p>
                                </div>
                            </div>
                    </div>
                        </body>
                    </html>`;



                fs.writeFile(`./html/${row[0]}_${row[2]}` + ".html", employeeData, (err) => {
                    if (err) throw err;

                })
            });
        }
        else {
            console.log('No data found.');
        }


    });
}
