var { google } = require('googleapis');
var sheets = google.sheets('v4');
var authorization = require("./authorization.js");

var exports = module.exports = {};


exports.insert = function (data, SHEET_ID, SHEET_NAME) {
    return new Promise((done, error) => {
        authorization.prepare((auth) => {
            var promise = saveFormData(auth, data, SHEET_ID, SHEET_NAME);
            promise.then(() => done(), (err) => error(err));
        });
    });
}

function saveFormData(auth, data, SHEET_ID, SHEET_NAME) {
    return new Promise((done, error) => {
        data.timestamp = new Date().toLocaleString();

        // nacti hlavicky v listu
        getSheetHeader(auth, SHEET_ID, SHEET_NAME).then((response) => {
            // zkontroluj je - tato sync funkce vrati spravne poradi hlavicek a hodnot s ohledem na to,
            // jak ma uzivatel serazene sloupce v sheetu
            var result = checkHeader(response, data, SHEET_ID, SHEET_NAME);

            // vloz data do sheetu
            insertData(auth, result.orderedValues, SHEET_ID, SHEET_NAME).then(() => {
                // pokud je potreba, zmen hlavicky.
                if (result.anyChangesToKeys)
                    updateHeader(auth, result.orderedKeys, SHEET_ID, SHEET_NAME).then(
                        () => done(),
                        (err) => error(err));
                else
                    done();
            }, (err) => error(err));
        }, (err) => error(err));
    });
}

function getSheetHeader(auth, SHEET_ID, SHEET_NAME) {
    return new Promise(function (done, error) {
        // GET REQUEST
        sheets.spreadsheets.values.get({
            auth: auth,
            spreadsheetId: SHEET_ID,
            range: SHEET_NAME + '!A1:Z1',
        }, function (err, response) {
            if (err) {
                console.log('The API returned an error: ' + err);
                error(err);
                return;
            }

            done(response);
        });
    });
}

function checkHeader(response, data) {
    var orderedValues = [];
    var orderedKeys = [];

    var rows = response.data.values;
    var j = 0;
    if (rows && rows.length != 0) {
        var row = rows[0];
        for (j = 0; j < row.length; j++) {
            var col = row[j];
            orderedKeys[j] = col;
            if (data[col])
                orderedValues[j] = data[col];
        }
    }

    var anyChangesToKeys = false;
    for (var key in data) {
        if (!orderedKeys.some(function (k) { return k == key })) {
            orderedKeys[j] = key;
            orderedValues[j] = data[key];
            anyChangesToKeys = true;
            j++;
        }
    }

    return {
        'orderedKeys': orderedKeys,
        'orderedValues': orderedValues,
        'anyChangesToKeys': anyChangesToKeys
    };
}

function updateHeader(auth, orderedKeys, SHEET_ID, SHEET_NAME) {
    return new Promise(function (done, error) {
        sheets.spreadsheets.values.update({
            auth: auth,
            spreadsheetId: SHEET_ID,
            range: SHEET_NAME + '!A1:Z1',
            valueInputOption: 'USER_ENTERED',
            resource: {
                'values': [orderedKeys]
            }
        }, function (err, response) {
            if (err) {
                error(err);
                return;
            }

            if (response && response.data && response.data && response.data.updatedRows == 1)
                error();
            else
                done();
        });
    });
}

function insertData(auth, orderedValues, SHEET_ID, SHEET_NAME) {
    return new Promise(function (done, error) {
        sheets.spreadsheets.values.append({
            auth: auth,
            spreadsheetId: SHEET_ID,
            range: SHEET_NAME + '!A2:Z2', // dvojka je strasne dulezita. To aby se nemohla stat, pri prazdne tabulce, ze bude zaznam vlozen na prvni radek
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: {
                'values': [orderedValues]
            }
        }, function (err, response) {
            if (err) {
                error(err);
                return;
            }

            if (response && response.data && response.data.updates && response.data.updates.updatedRows == 1)
                done();
            else
                error('bad update');
        });
    });
}