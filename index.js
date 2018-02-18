// https://pinboard.in/api

// TODO: More code comments
// TODO: More logs
// TODO: Perf
// TODO: Disclaimer
// TODO: Backup
// TODO: Clearer UX with "status" column
// TODO: Guide
// TODO: More standard eslint

// Config
// ===
var PINBOARD_FIELDS = [
  'url',
  'title', // called "description" in API due to historical reasons
  'shared',
  'toread',
  'tags',
  'extended' // longer "description"
];

var LOGIC_FIELDS = [
  'reference', // original url
  'status'
];

var MERGED_FIELDS = Array.concat(LOGIC_FIELDS, PINBOARD_FIELDS);

// Hooks
// ===
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pinboard')
    .addItem('Load', 'load')
    .addItem('Update', 'update')
    .addToUi();
}


function onEdit(ev) {
  var range = ev.range;
  var note = range.getNote();
  var columnNumber = range.getColumn();
  var rowNumber = range.getRow();

  if (
    columnNumber > 2 && // Ignore "reference" and "status"
    rowNumber > 1 // Ignore header
  ) {
    var row = range.getRow();
    var statusRange = range.getSheet().getRange(row, 2);
    statusRange.setValue('UPDATE');

    if (!note) {
      range.setNote('Original: ' + ev.oldValue);
    }
  }
}



function load() {
  // Ask for token
  // ===
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Pinboard API token',
    'https://pinboard.in/settings/password',
    ui.ButtonSet.OK_CANCEL
  );

  var button = result.getSelectedButton();
  var token = result.getResponseText();
  if (button === ui.Button.OK) {
    // Load
    // ===
    var url = 'https://api.pinboard.in/v1/posts/all';
    url += '?auth_token=' + token;
    url += '&format=json';

    Logger.log('Fetch: ' + url);

    // TODO: Handle errors
    // 429 Too Many Requests

    // Parse
    var res = JSON.parse(UrlFetchApp.fetch(url).getContentText());

    Logger.log('Response: ' + res);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.insertSheet(generateSheetName() + ' | TOKEN:' + token);
    ss.setActiveSheet(sheet);

    // Insert into sheet
    // ===
    sheet.appendRow(MERGED_FIELDS);
    res.posts.forEach(function (post) {
      var row = PINBOARD_FIELDS.map(function (x) {
        return post[x];
      });
      // LOGIC_FIELDS initial values
      sheet.appendRow([post.url, ''].concat(row));
    });

    // Protect & format
    // ===

    // Layout
    formatActiveSheet();

    // Reference
    var rangeReference = sheet.getRange('A2:A');
    rangeReference.protect().setDescription('This is used for scripting purposes, do not change');
    rangeReference.setBackground('F7F7F7');

    // Format
    // ===
  } else {
    ui.alert('👋 Bye... If you need help, contact @straaajk on Twitter');
  }
}

function update() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var statuses = sheet.getRange('B2:B').getValues();
  statuses.forEach(function (status, i) {
    var rowNumber = i + 2;
    // (row, column, numRows, numColumns)
    var row = sheet.getRange(rowNumber, 1, 1, MERGED_FIELDS.length).getValues()[0];
    var model = {};
    row.forEach(function (val, j) {
      model[MERGED_FIELDS[j]] = val;
    });

    Logger.log(model);

    var statusRange = sheet.getRange(rowNumber, 2);

    if (model.status === 'UPDATE') {
      var replace = 'yes'; // initial

      if (model.reference !== model.url) {
        var deleteUrl = 'https://api.pinboard.in/v1/posts/delete';
        deleteUrl += '?auth_token=' + getTokenFromActiveSheet();
        deleteUrl += '&format=json';
        deleteUrl += '&url=' + encodeURIComponent(model.reference);

        var deleteRes = JSON.parse(UrlFetchApp.fetch(deleteUrl).getContentText());
        if (deleteRes.result_code !== 'done') {
          replace = 'no';
          // TODO: Handle item not found
          throw new Error('Deleting failed, cancelling...');
        }
      }

      if (model.url) { // If empty url, means DELETE
        var url = 'https://api.pinboard.in/v1/posts/add';
        url += '?auth_token=' + getTokenFromActiveSheet();
        url += '&format=json';
        url += '&url=' + encodeURIComponent(model.url);
        url += '&description=' + encodeURIComponent(model.title);
        url += '&extended=' + encodeURIComponent(model.extended);
        url += '&tags=' + encodeURIComponent(model.tags);
        url += '&shared=' + model.shared;
        url += '&toread=' + model.toread;
        url += '&replace=' + replace;

        var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        var text = response.getContentText();
        var json = JSON.parse(text);

        Logger.log(JSON.stringify(json, null, 2));
        if (json.result_code !== 'done') {
          statusRange.setValue('ERROR');
          statusRange.setNote(JSON.stringify(json, null, 2));
          return // Do not continue
          // TODO: Better error
        }
      }

      statusRange.setValue('DONE');
    } else {
      // do nothing
    }
  });
}

// Utils
// ===

function generateSheetName() {
  var datetime = new Date().toISOString();
  return datetime.slice(0, datetime.indexOf('.')).replace('T', ' ');
}

function formatActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 50);
  sheet.hideColumns(1);
  sheet.setColumnWidth(2, 80); // status
  sheet.setColumnWidth(3, 300); // url
  sheet.setColumnWidth(4, 300); // description
  sheet.setColumnWidth(5, 50); // shared
  sheet.setColumnWidth(6, 50); // shared
  sheet.setColumnWidth(7, 250); // tags
}

function getTokenFromActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var name = sheet.getName();
  var token = name.split('TOKEN:')[1]

  if (typeof token !== 'string') {
    throw new Error("Failed parsing token from Sheet name");
  }

  return token;
}


// Testing
// ===

function testOnEdit() {
  onEdit({
    user: Session.getActiveUser().getEmail(),
    source: SpreadsheetApp.getActiveSpreadsheet(),
    range: SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
    value: SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
    oldValue: SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
    authMode: 'LIMITED'
  });
}