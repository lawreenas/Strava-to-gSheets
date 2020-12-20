// Taken from: https://www.benlcollins.com/spreadsheets/strava-api-with-google-sheets/
// Strava API: https://developers.strava.com/docs/reference/#api-Activities-getLoggedInAthleteActivities
// OAuth Library: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
// Version 1.1

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Strava App')
    .addItem('Get data', 'getStravaActivityData')
    .addToUi();
}

var sheet = null;

function getStravaActivityData() {
  var stravaData = callStravaAPI();
  
  if (stravaData) {
    printActivities(stravaData);
  }
}

function printActivities(stravaData) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.getActiveSheet();

    var totals = {
      duration: 0,
      distance: 0,
      elevation: 0
    };
    
    var byDate = groupStravaActivitiesByDay(stravaData);
    var row = sheet.getActiveCell().getRow();
    var col = sheet.getActiveCell().getColumn();

    for(colIdx = col; colIdx < 20; colIdx++) {
      var dateToGet = sheet.getRange(row - 1, colIdx).getValue();
      var currentCellValue = sheet.getRange(row, colIdx).getValue();
      
      if (!dateToGet) break;
      var activities = byDate[dateToGet];
      
      if (activities) {
        var data = "";
        activities.forEach(
          function(a){
            if (isRunActivity(a)) {
              totals = updateTotals(totals, a.distance, a.elapsed_time, a.total_elevation_gain);
            }
            data = printActivityData(a, currentCellValue) + "\n" + data;
          })
        sheet.getRange(row, colIdx).setValue(data);  
      } else {
        sheet.getRange(row, colIdx).setValue("-");
      }
    }
    
    writeTotals(row, totals);
}

function isRunActivity(a) {
  return a.type == "Run";
}
function printActivityData(a, currentCellValue) {
  if (isRunActivity(a)) {
    var laps = "";
    if (a.workout_type == 3 || currentCellValue.includes("reniruote")) {
      laps = printLaps(a.id);
    }
    return printRun(a) + laps; 
  }
  if (a.type == "Workout") {
    return printWorkout(a);
  }
  return printRun(a);
}

function printRun(a) {
  return a.name + " \n" +
    "🩴" + getDistance(a.distance) + " km " + getPace(a.average_speed) + "/km \n" + 
    "❤️" + getHr(a.average_heartrate) + " bpm \n" +
    "⛰️" + a.total_elevation_gain + " m+ \n" + 
    "⏱" + getDuration(a.elapsed_time)+ " \n\n"; 
}

function printLaps(activityId) {
  var laps = fetchRunLaps(activityId);
  var lapData = "Laps:\n";
  
  laps.forEach(
    function(lap){
      lapData = lapData + printLap(lap);
    }
  );
  return lapData;
}

function printLap(lap) {
  return "- " +
    getDistance(lap.distance) + " km " +
    getPace(lap.average_speed) + "/km " +
    "❤️" + getHr(lap.average_heartrate) + "/" + getHr(lap.max_heartrate) + "\n";
}

function printWorkout(a) {
    return a.name + " \n" +
      "⏱" + getDuration(a.elapsed_time)+ " \n\n"; 
}
    
function groupStravaActivitiesByDay(stravaData) {
  var byDate = {};
  stravaData.map(function(a) {
    var date = new Date(a.start_date).getDate();
    var currentDateActivities = byDate[date];
    if (currentDateActivities) {
      byDate[date] = [...currentDateActivities, a];
    } else {
      byDate[date] = [a];
    }
  });
  return byDate;
}

function updateTotals(totals, distance, time, elev) {
  return {
    duration: totals.duration + time,
    distance: totals.distance + distance,
    elevation: totals.elevation + elev
  };
}

function writeTotals(row, totals) {
  sheet.getRange(row, 9).setValue(getDistance(totals.distance));  
  sheet.getRange(row, 10).setValue(getDuration(totals.duration));  
  sheet.getRange(row, 11).setValue(totals.elevation);  
}

// Convert min/s -> min/km
function getPace(metersPerSec) {
  var secondsPerKm = parseInt(1/(metersPerSec/1000));
  var paceMin = Math.floor(secondsPerKm/60);
  var paceSec = Math.floor(secondsPerKm-paceMin*60);
  
  return paceMin + ":" + (paceSec < 10 ? "0"+paceSec : paceSec);
}


function getDistance(stravaDistance)  {
 return Number.parseFloat(stravaDistance / 1000).toFixed(2);
}

function getDuration(seconds) {
    var sec_num = parseInt(seconds, 10);
    var hours   = Math.floor(sec_num / 3600);
    var minutes = Math.floor((sec_num - (hours * 3600)) / 60);
    var seconds = sec_num - (hours * 3600) - (minutes * 60);

    if (minutes < 10) { minutes = "0" + minutes;}
    if (seconds < 10) { seconds = "0" + seconds;}
    return hours+'h '+ minutes + 'm ' + seconds + "s";
}

function getHr(hr) {
  return hr ? Math.round(hr) : "--"; 
}


// ******
// STRAVA
// ******

// call the Strava API
function callStravaAPI() {
  
  // set up the service
  var service = getStravaService();
  
  if (service.hasAccess()) {
    Logger.log('App has access.');
    
    var till = Math.fround(Date.now()/1000);
    var from = till - 1209600; //-14 days
    var max = 30;
  
    var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
    var params = '?before=' + till + '&after=' + from + '&per_page=30';

    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
    
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
    
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
    
    return response;  
  }
  else {
    Logger.log("App has no access yet.");
    
    var authorizationUrl = service.getAuthorizationUrl();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet(); 
    
    sheet.getActiveCell().setValue(authorizationUrl);
    
    Logger.log("Open the following URL and re-run the script: %s", authorizationUrl);
  }
}

function fetchRunLaps(id) {
   var service = getStravaService();
  
    var endpoint = 'https://www.strava.com/api/v3/activities/' + id + '/laps';
    var params = '';

    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
    
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
    
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
    
    return response;  
}
  
// configure the service
function getStravaService() {
  var id = '55641';
  var secret = '456f50520af93dd69e8053ac91ef81b9b547a8b0';

  return OAuth2.createService('Strava')
    .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
    .setTokenUrl('https://www.strava.com/oauth/token')
    .setClientId(id)
    .setClientSecret(secret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('activity:read_all');
}

// handle the callback
function authCallback(request) {
  var stravaService = getStravaService();
  var isAuthorized = stravaService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}