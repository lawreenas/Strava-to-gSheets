// Taken from: https://www.benlcollins.com/spreadsheets/strava-api-with-google-sheets/
// Strava API: https://developers.strava.com/docs/reference/#api-Activities-getLoggedInAthleteActivities
// 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
// custom menu

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Strava App')
    .addItem('Get data', 'getStravaActivityData')
    .addToUi();
}

var sheet = null;

// Get athlete activity data
function getStravaActivityData() {
 
  // get the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getActiveSheet(); //ss.getSheetByName('Sheet1');

  var stravaData = callStravaAPI();
  
  if (stravaData) {
    var totals = {
      duration: 0,
      distance: 0,
      elevation: 0
    };
    
    var byDate = groupStravaActivitiesByDay(stravaData);
    var row = sheet.getActiveCell().getRow();
    var col = sheet.getActiveCell().getColumn();
    
    for(colIdx = col; colIdx < 20; colIdx++) {
      var dateToGet = sheet.getRange(row - 2, colIdx).getValue();
      
      if (!dateToGet) break;
      var activities = byDate[dateToGet];
      
      if (activities) {
        var data = "";
        activities.forEach(
          function(a){
            totals = updateTotals(totals, a.distance, a.elapsed_time, a.total_elevation_gain);
            data = data + printActivityData(a);
          })
        sheet.getRange(row, colIdx).setValue(data);  
      } else {
        sheet.getRange(row, colIdx).setValue("-");
      }
    }
    
    writeTotals(row, totals);
  }
}

function printActivityData(a) {
  if (a.type == "Run") {
    return printRun(a); 
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
 return Number.parseFloat(stravaDistance / 1000).toPrecision(3);
}

function getDuration(seconds) {
    var sec_num = parseInt(seconds, 10); // don't forget the second param
    var hours   = Math.floor(sec_num / 3600);
    var minutes = Math.floor((sec_num - (hours * 3600)) / 60);
    var seconds = sec_num - (hours * 3600) - (minutes * 60);

    if (minutes < 10) {minutes = "0"+minutes;}
    if (seconds < 10) {seconds = "0"+seconds;}
    return hours+'h '+minutes+'m '+seconds+"s";
}

function getHr(hr) {
 return Math.round(hr); 
}


// STRAVA



// call the Strava API
function callStravaAPI() {
  
  // set up the service
  var service = getStravaService();
  
  if (service.hasAccess()) {
    Logger.log('App has access.');
    
    var epochNow = Math.fround(Date.now()/1000);
    var epochWeekAgo = epochNow - 691200; //-8 days
  
    var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
    var params = '?before=' + epochNow + '&after=' + epochWeekAgo + '&per_page=30';

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

  
// configure the service
function getStravaService() {
  
   var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName('Strava');

  var id = String(sheet.getRange(1,1).getValue()); //'55641'; //
  var secret = sheet.getRange(1,2).getValue(); // '456f50520af93dd69e8053ac91ef81b9b547a8b0'; //
 
  return OAuth2.createService('Strava10')
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