// Global vars
var ss = SpreadsheetApp.getActiveSpreadsheet()
  , sheet = ss.getActiveSheet()
  , rangeData = sheet.getDataRange();

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  createMenu();
}

/**
 * Create our custom UI menu with some basic ops
 */
function createMenu() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Tools')
    .addItem('Fix Range Error', 'fixMe')
    .addSeparator()
    .addItem('About', 'about')
    .addToUi();
}

/**
 * Take an input of the current date and attempt to determine which column
 * to match given that date. This will be our working column for 24 hours.
 */
function findWorkingColumn(today) {
  var firstColumn = 2
    , lastColumn = rangeData.getLastColumn() - 1
    , dateSearchRange = sheet.getRange(firstColumn, 2, 1, lastColumn)
    , rangeValues = dateSearchRange.getValues();

  for (i = 0; i < lastColumn; i++) {
    var date = Utilities.formatDate(rangeValues[0][i], "GMT", "yyyy-MM-dd");

    if (date === today) {
      return i + 2;
    }
  }
}

/**
 * Take an input of our current hour of the day and attempt to 
 * determine which row that correlates to given a 24 hour time range
 */
function findWorkingRow(time) {
  var firstRow = 3
    , lastRow = 24
    , hourSearchRange = sheet.getRange(firstRow, 1, lastRow, 1)
    , rangeValues = hourSearchRange.getValues();

  for (i = 0; i < lastRow; i++) {
    // Some weird conversion of these hour formatted cells here... Gotta offset by 6..idk why.
    var hour = Utilities.formatDate(rangeValues[i][0], "GMT-6", "HH");

    if (hour === time) {
      return i + 3;
    }
  }
}

/**
 * Generate the query param for our REST call to SFCC to get orders within
 * a given timeframe.
 */
function generateQueryParams(row, column) {
  var targetDate = sheet.getRange(2, column).getDisplayValue()
    , targetFromHour = sheet.getRange(row, 1).getDisplayValue()
    , targetToHour = targetFromHour.substring(0, targetFromHour.indexOf(":"))
    , formattedFromDateTime = targetDate + 'T' + targetFromHour + '.000Z'
    , formattedToDateTime = targetDate + 'T' + targetToHour + ':59:59.999Z'
    , queryParam = '?from=' + formattedFromDateTime + '&to=' + formattedToDateTime;

  return queryParam;
}

/**
 * Build an array of sheets withint the spreadsheet leaving out the 
 * Rollup sheet, as this is intended to be a summation sheet
 */
function getSheetList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
    , sheetsArray = [];

  ss.getSheets().forEach(function (sheet) {
    sheetsArray.push(sheet.getName());
  });

  // Filter out the rollup tab of course, it's not a country site, just our totals,
  // but also filter out the sites that haven't been built and put into production yet.
  // They return a 403 presently.
  var filteredArray = sheetsArray.filter(function (sheet) {
    return sheet !== 'Rollup'
      && sheet != 'brand-1'
      && sheet != 'brand-2'
      && sheet != 'brand-3';
  });

  return filteredArray;
}

/**
 * Basic init function..possibly will remove/consolidate with fixMe() method.
 */
function init() {
  var currentDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd")
    , date = new Date()
    , currentHour = date.getUTCHours()
    , currentHour = ("0" + currentHour).slice(-2)
    , currentMin = date.getMinutes()
    , workingColumn = findWorkingColumn(currentDate)
    , workingRow = findWorkingRow(currentHour)
    , queryParam = generateQueryParams(workingRow, workingColumn)
    , sheetList = getSheetList();


  // Google Apps Script hourly trigger guarantees +/- 15 minutes...this won't work for us.
  // We need more consistency, once an hour. We'll trigger every minute and take action only if it's at the
  // end of the hour, minute 59 specifically. Otherwise we'd be calling OCAPI every minute,
  // and there are limits. Not to mention this would be terribly inefficienct.
  if (currentMin === 59) {
    for (var s in sheetList) {
      var storeId = ss.getSheetByName(sheetList[s])
        , url = 'https://host.apigee.net/commerce/v1/' + storeId.getName() + '/orders' + queryParam
        , options = {
          'headers': {
            'x-apikey': '${APIKEY}'  // If your API doesn't require this you can safely remove the options object
          }
        };

      try {
        var response = UrlFetchApp.fetch(url, options)
          , json = response.getContentText()
          , data = JSON.parse(json);

        // Write our order count to the cell and wait to do the next site
        // so we don't hit any OCAPI API thresholds
        storeId.getRange(workingRow, workingColumn).setValue(data.order_count);
        Utilities.sleep(5000);

      }
      catch (e) {
        // If the service call fails, we'll log it and set the cell color to red so we know
        // what we need to make a new service call for. We'll send an email to myself as well.
        Logger.log('Error calling Apigee: ' + e);
        MailApp.sendEmail('ryan@nighthauk.com', 'Peak Order Rate Error: ' + storeId.getName(), e);
        storeId.getRange(workingRow, workingColumn).setBackground('#991A00');
      }
    }
  }
}

/**
 * We need to account for failures on our service call to Apigee/OCAPI. We'll set the cell to red if it fails.
 * This method will be used to fix an empty, errored cell, and reset the color if corrected.
 */
function fixMe() {
  var ui = SpreadsheetApp.getUi()
    , targetCell = sheet.getSelection().getCurrentCell()
    , dateCol = targetCell.getColumn()
    , hourRow = targetCell.getRow()
    , queryParam = generateQueryParams(hourRow, dateCol)
    , sheetList = getSheetList();

  for (var s in sheetList) {
    var storeId = ss.getSheetByName(sheetList[s])
      , url = 'https://host.apigee.net/commerce/v1/' + storeId.getName() + '/orders' + queryParam
      , options = {
        'headers': {
          'x-apikey': '${APIKEY}'  // If your API doesn't require this you can safely remove the options object
        }
      };


    try {
      var response = UrlFetchApp.fetch(url, options)
        , json = response.getContentText()
        , data = JSON.parse(json);

      // Write our order count to the cell and wait to do the next site
      // so we don't hit any OCAPI API thresholds
      storeId.getRange(hourRow, dateCol).setValue(data.order_count);
      storeId.getRange(hourRow, dateCol).setBackground(null);
      Utilities.sleep(2500);
    }
    catch (e) {
      // If the service call fails, we'll log it and set the cell color to red so we know
      // what we need to make a new service call for. We'll send an email to myself as well.
      Logger.log('Error calling Apigee: ' + e);
      MailApp.sendEmail('ryan@nighthauk.com', 'Order Rate Error: ' + storeId.getName(), e);
      storeId.getRange(hourRow, dateCol).setBackground('#991A00');
    }
  }
}

/**
 * Pointless fun
 */
function about() {
  SpreadsheetApp.getUi()
    .alert('Created by Ryan Hauk. Eliminating monotonous tasks and human airor with automation.');
}