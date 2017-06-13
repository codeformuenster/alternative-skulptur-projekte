
var fs = require('fs');
var readline = require('readline');
var google = require('googleapis');
var googleAuth = require('google-auth-library');

// If modifying these scopes, delete your previously saved credentials
// at ~/.credentials/sheets.googleapis.com-nodejs-quickstart.json
var SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
var TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH ||
    process.env.USERPROFILE) + '/.credentials/';
var TOKEN_PATH = TOKEN_DIR + 'sheets.googleapis.com-nodejs-quickstart.json';

// Load client secrets from a local file.
fs.readFile('client_secret.json', function processClientSecrets(err, content) {
  if (err) {
    console.log('Error loading client secret file: ' + err);
    return;
  }
  // Authorize a client with the loaded credentials, then call the
  // Google Sheets API.
  authorize(JSON.parse(content), listMajors);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  var clientSecret = credentials.installed.client_secret;
  var clientId = credentials.installed.client_id;
  var redirectUrl = credentials.installed.redirect_uris[0];
  var auth = new googleAuth();
  var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, function(err, token) {
    if (err) {
      getNewToken(oauth2Client, callback);
    } else {
      oauth2Client.credentials = JSON.parse(token);
      callback(oauth2Client);
    }
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
  var authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES
  });
  console.log('Authorize this app by visiting this url: ', authUrl);
  var rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  rl.question('Enter the code from that page here: ', function(code) {
    rl.close();
    oauth2Client.getToken(code, function(err, token) {
      if (err) {
        console.log('Error while trying to retrieve access token', err);
        return;
      }
      oauth2Client.credentials = token;
      storeToken(token);
      callback(oauth2Client);
    });
  });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
  try {
    fs.mkdirSync(TOKEN_DIR);
  } catch (err) {
    if (err.code != 'EEXIST') {
      throw err;
    }
  }
  fs.writeFile(TOKEN_PATH, JSON.stringify(token));
  console.log('Token stored to ' + TOKEN_PATH);
}

/**
 * Print the names and majors of students in a sample spreadsheet:
 * https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 */
function listMajors(auth) {
  var sheets = google.sheets('v4');
  sheets.spreadsheets.values.get({
    auth: auth,
    spreadsheetId: '1I9h_U4kWcQ9OaKEhPlK-I5u0ToI_QR9nZAHrMTSBIIQ',
    range: 'Formularantworten 1!A1:V',
  }, function(err, response) {
    if (err) {
      console.log('The API returned an error: ' + err);
      return;
    }
    var rows = response.values;
    if (rows.length == 0) {
      console.log('No data found.');
    } else {
      console.log('var SchattenSkulpturEvents = [');
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        // Print columns A and E, which correspond to indices 0 and 4
        /*.
        ["03.05.2017 12:06:30",
          "Weyand, Georgios",
          "Der Sonnenwagen",
          "http://www.facebook.com/DerSonnenwagen/",
          "Kinetische Gemeinschafts-Skulptur mit medialer Installation",
          "","24h ohne mediale Installation an der Dammstr. 19 - an ausgewählten Tagen und Orten mit medialer Installation",
          "Dammstr. 19 / verschiedene Orte in Münster",
          "Sonstiges",
          [51.951001, 7.632366],
          103
        ],

        [ '12.06.2017 19:24:13',
          'Kubeja, Jochen',
          'Graffiti-Werke - Urban Street Art',
          'www.wohnzimmer-ev.de',
          'Am Hawerkamp finden sich überall außergewöhnliche Graffitis, geschweisste Skulpturen und vielfältige Urban Street Art.\n"Ein besonderes Gebiet in Münster zum Ausgehen ist der Hawerkamp [... Hier] hat sich eine außergewöhnliche künstlerische und musikalische Szene etabliert, die jährlich über 120.000 Besucher anzieht. Neben den zahlreichen Musikevents in den einzelnen Clubs ist diese Besucherzahl auch auf das jährlich stattfindende Hawerkamp-Festival und den offenen Ateliers der über 50 Künstler zurückzuführen. Bei der Bewerbung Münsters zur Kulturhauptstadt Europas 2010 war der Hawerkamp einer der kulturellen Schwerpunkte neben Prinzipalmarkt und Picasso-Museum. In einem bundesweit einzigartigen Projekt wird seit dem 1. Januar 2006 das alternativ geprägte Gelände durch den Verein "Erhaltet den Hawerkamp" eigenverantwortlich verwaltet."\nQuelle: http://wiki.muenster.org/index.php/Hawerkamp',
          'Rund um die Uhr',
          'Am Hawerkamp 31, Münster' ]
  */

        if (!row[0].match(/^\d\d\./)) {
          continue;
        }

        var wantedcols = [0,1,3,6,9,11,12,13,15,17,18,19,20];
        var newrow = [];
        for (var j = 0; j<wantedcols.length;j++) {
          newrow.push( row[wantedcols[j]] );
        }

        if (newrow[12]) {
          newrow[12] = newrow[12].replace(/"|\s/g,"").split(",");
        }
//        newrow[0] = new Date(newrow[0]).toISOString();
        console.log(newrow);
        console.log(",");
      }
      console.log("[]];");
    }
  });
}
