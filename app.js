const { google } = require("googleapis");
const fs = require("fs");
const xlsx = require("xlsx");
var workbook = xlsx.readFile("excel.xlsx");
var sheetNames = workbook.SheetNames;

function GetKeys() {
  return new Promise((resolve, reject) => {
    fs.readFile("keys.json", "utf8", function readFileCallback(err, data) {
      if (err) {
        reject(err);
      } else {
        obj = JSON.parse(data);
        resolve(obj);
      }
    });
  });
}

function GetConfig() {
  return new Promise((resolve, reject) => {
    fs.readFile("config.json", "utf8", function readFileCallback(err, data) {
      if (err) {
        reject(err);
      } else {
        obj = JSON.parse(data);
        resolve(obj);
      }
    });
  });
}

function delay() {
  return new Promise(resolve => setTimeout(resolve, 15000));
}

function GetSheetArray(name) {
  var sheet = workbook.Sheets[name];
  var result = [];
  var row;
  var rowNum;
  var colNum;
  var range = xlsx.utils.decode_range(sheet["!ref"]);
  for (rowNum = range.s.r; rowNum < 146; rowNum++) {
    row = [];
    for (colNum = range.s.c; colNum < 21; colNum++) {
      var nextCell = sheet[xlsx.utils.encode_cell({ r: rowNum, c: colNum })];
      if (typeof nextCell === "undefined") {
        row.push("");
      } else row.push(nextCell.w);
    }
    result.push(row);
  }
  return result;
}

async function CopyTemplate(token, tempGID, spreadsheetId, name, array) {
  await delay();
  var sheets = google.sheets("v4");

  sheets.spreadsheets.sheets
    .copyTo({
      oauth_token: token,
      spreadsheetId: spreadsheetId,
      sheetId: tempGID,
      requestBody: {
        destinationSpreadsheetId: spreadsheetId
      }
    })
    .then(
      async data => {
        console.log(name);
        var newSheetId = data.data.sheetId;
        sheets.spreadsheets.batchUpdate({
          oauth_token: token,
          spreadsheetId: spreadsheetId,
          requestBody: {
            requests: [
              {
                updateSheetProperties: {
                  properties: {
                    sheetId: newSheetId,
                    title: name,
                    tabColor: {
                      red: Math.random() * (255 - 0) + 0,
                      green: Math.random() * (255 - 0) + 0,
                      blue: Math.random() * (255 - 0) + 0
                    }
                  },
                  fields: "title"
                }
              }
            ]
          }
        });
        await delay();
        console.log("Ubaci za " + name);
        google
          .sheets("v4")
          .spreadsheets.values.batchUpdate({
            oauth_token: token,
            spreadsheetId: spreadsheetId,

            resource: {
              valueInputOption: "RAW",
              data: [
                {
                  range: `${name}!A1:U147`,
                  values: array
                }
              ]
            }
          })
          .then(data => {})
          .catch(error => {
            console.log(error);
          });
      },
      error => {
        console.log(error);
      }
    );
}

//MAIN
async function main() {
  //get keys
  var keys = await GetKeys();
  //create client
  const client = new google.auth.JWT(
    keys.client_email,
    null,
    keys.private_key,
    [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/spreadsheets.readonly"
    ]
  );
  //authorize client, get access_token
  client.authorize(async (err, tokens) => {
    if (err) {
      console.log(err);
      return;
    } else {
      //get Google Sheet information
      const token = tokens.access_token;
      var config = await GetConfig();
      var spreadsheetId = config.spreadsheetId;
      tempGID = config.tempGID;

      for (const name of sheetNames) {
        //get data from Excel
        var array = await GetSheetArray(name);
        //create sheet, rename sheet, insert data
        await CopyTemplate(token, tempGID, spreadsheetId, name, array);
      }
    }
  });
}

main();
