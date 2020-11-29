var CLIENT_ID = '55641';
var CLIENT_SECRET = '456f50520af93dd69e8053ac91ef81b9b547a8b0';


// configure the service
function getStravaService() {
return OAuth2.createService('Strava2')
  .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
  .setTokenUrl('https://www.strava.com/oauth/token')
  .setClientId(CLIENT_ID)
  .setClientSecret(CLIENT_SECRET)
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