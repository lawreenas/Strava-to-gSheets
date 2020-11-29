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
    
    for(idx = col; idx > 0; idx--) {
      var dateToGet = sheet.getRange(row - 2, idx).getValue();
      
      if (!dateToGet) break;
      var activities = byDate[dateToGet];
      
      if (activities) {
        var data = "";
        activities.forEach(
          function(a){
            totals = updateTotals(totals, a.distance, a.elapsed_time, a.total_elevation_gain);
            data = data + printActivityData(a);
          })
        sheet.getRange(row, idx).setValue(data);  
      } else {
        sheet.getRange(row, idx).setValue("-");
      }
    }
    
    writeTotals(row, totals);
  }
}

function printActivityData(a) {
  return a.name + " \n" +
    "🩴" + getDistance(a.distance) + " km \n" + 
    "❤️" + a.average_heartrate + " bpm \n" +
    "⛰️" + a.total_elevation_gain + " m+ \n" + 
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

function getPace(metersPerSec) {
  return metersPerSec;
}


function getDistance(stravaDistance)  {
 return Number.parseFloat(stravaDistance / 1000).toPrecision(4);
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

// call the Strava API
function callStravaAPI() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName('Strava');

  var CLIENT_ID = String(sheet.getRange(1,1).getValue()); //'55641';
  var CLIENT_SECRET = sheet.getRange(2,1).getValue(); // '456f50520af93dd69e8053ac91ef81b9b547a8b0';
  
  
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