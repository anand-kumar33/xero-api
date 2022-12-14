var ss = SpreadsheetApp.getActiveSpreadsheet();
var sh = ss.getSheetByName('data');

var CLIENT_ID = 'client-id';
var CLIENT_SECRET = 'client-secret';

function getBankTransactions() {
  var service = getService();
  if (service.hasAccess()) {
    // Retrieve the tenantId from storage.
    var tenantId = service.getStorage().getValue('tenantId');
    // Make a request to retrieve user information.

    // var previousday = new Date();
    // previousday.setDate(previousday.getDate()-1);

    // y=Utilities.formatDate(previousday, 'GMT+8', 'yyyy')
    // m=Utilities.formatDate(previousday, 'GMT+8', 'M')
    // d=Utilities.formatDate(previousday, 'GMT+8', 'd')

    var url = 'https://api.xero.com/api.xro/2.0/BankTransactions?order=Date DESC' //?where=Date==DateTime(' + y + ', '+m+', '+d+')';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Xero-tenant-id': tenantId
      },
    });

    //Logger.log(response);

    if(response.getResponseCode() === 200){
      var result = JSON.parse(response).BankTransactions;
      let header = ["BankTransactionID", "AccountID", "Code", "Name", "Type", "Reference", "IsReconciled", "HasAttachments", "ContactID", "UserName", "DateString", "Status", "LineAmountTypes", "SubTotal", "TotalTax", "Total", "UpdatedDateUTC", "CurrencyCode"];

      let finalData = [];

      for(let data of result){
        let BankTransactionID = data.BankTransactionID;
        let { AccountID, Code, Name } = data.BankAccount;
        let Type = data.Type;
        let Reference = data.Reference;
        let IsReconciled = data.IsReconciled;
        let HasAttachments = data.HasAttachments;
        let { ContactID, Name: UserName} = data.Contact || '';
        let DateString = Utilities.formatDate(new Date(data.DateString), 'GMT', 'dd MMM yy');
        let Status = data.Status;
        let LineAmountTypes = data.LineAmountTypes;
        let SubTotal = data.SubTotal;
        let TotalTax = data.TotalTax;
        let Total = data.Total;
        let UpdatedDateUTC = data.UpdatedDateUTC;
        let CurrencyCode = data.CurrencyCode;

        finalData.push([BankTransactionID, AccountID, Code, Name, Type, Reference, IsReconciled, HasAttachments, ContactID, UserName, DateString,
        Status, LineAmountTypes, SubTotal, TotalTax, Total, UpdatedDateUTC, CurrencyCode]);
      }

      if(finalData.length > 0){
        //sh.getRange('A2:D').clearContent();
        finalData.unshift(header);
        sh.getRange(sh.getLastRow()+1, 1, finalData.length, finalData[0].length).setValues(finalData);
      }
    }

  
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',authorizationUrl);
    openUrl( authorizationUrl );
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService().reset();
}

/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('Xero')
    // Set the endpoint URLs.
    .setAuthorizationBaseUrl(
        'https://login.xero.com/identity/connect/authorize')
    .setTokenUrl('https://identity.xero.com/connect/token')

    // Set the client ID and secret.
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)

    // Set the name of the callback function that should be invoked to
    // complete the OAuth flow.
    .setCallbackFunction('authCallback')

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getScriptProperties())

    // Set the scopes to request from the user. The scope "offline_access" is
    // required to refresh the token. The full list of scopes is available here:
    // https://developer.xero.com/documentation/oauth2/scopes
    .setScope('accounting.settings.read accounting.transactions offline_access');
};

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    // Retrieve the connected tenants.
    var response = UrlFetchApp.fetch('https://api.xero.com/connections', {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      },
    });
    var connections = JSON.parse(response.getContentText());
    // Store the first tenant ID in the service's storage. If you want to
    // support multiple tenants, store the full list and then let the user
    // select which one to operate against.
    service.getStorage().setValue('tenantId', connections[0].tenantId);
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redict URI to register in the Dropbox application settings.
 */
function logRedirectUri() {
  Logger.log(OAuth2.getRedirectUri());
}


function openUrl( url ){
  var html = HtmlService.createHtmlOutput('<!DOCTYPE html><html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically.  Click below:<br/><a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(55);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}
