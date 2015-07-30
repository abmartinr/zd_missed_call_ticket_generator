var spreadsheet;
var page;
var range;
var values;
var subdomain;
var user;
var token;
var lastchecked;

function initialize() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  page = ss.getActiveSheet();
  range = page.getDataRange();
  values = range.getValues();
  var scriptProperties = PropertiesService.getScriptProperties();
  user = scriptProperties.getProperty('USER');
  token = scriptProperties.getProperty('TOKEN');
  subdomain = scriptProperties.getProperty('SUBDOMAIN');
  last_checked = scriptProperties.getProperty('LAST_CHECKED');
}

function parseSheet() {
  //Let's initialize the app
  initialize();
  
  var toremove = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][1] != "" || values[i][7] == "Completed" ||  values[i][3] == "Unknown") {
      //I'll delete the line if it has a ticket ID, if it's shown as completed on the call Status or if the caller is unknown
      Logger.log("Adding line " + (i + 1) + " for removal");
      toremove.push(i + 1);
    } else {
      Logger.log("Valid ticket to be created found on line " + (i + 1));
    }
  }
  // Now that we have all the lines to be removed, let's remove 
  for (var i = 0; i < toremove.length; i++) {
    var line = parseInt(toremove[i]);
    page.deleteRow(line - i);
  }
}

function generateTickets() {
  initialize();
  var from_list = [];
  var scriptProperties = PropertiesService.getScriptProperties();  
  
  for (var i = 1; i<values.length; i++) {
    last_checked = scriptProperties.getProperty('LAST_CHECKED');
    var date = new Date(values[1][2]);
    var pfrom = values[i][3].replace("(", "").replace(")", "").replace("-", "").replace(/\s+/g, '');
    var date = new Date(values[i][2]);
    Logger.log("Checking id: " +values[i][0]+". Last checked was " + last_checked);
    if (from_list.indexOf(pfrom) <= -1 && parseInt(values[i][0])>last_checked ){
      Logger.log("Creating ticket for id: " + values[i][0]);
      from_list.push(pfrom);
      var pto = values[i][4].replace(/\s+/g, '');;
      var data = {
        ticket: {
          description: 'Missed Call From ' + pfrom,
          via_id: 45,
          voice_comment: {
          call_duration: 0,
          from: pfrom,
          location: 'Dublin, Ireland',
          started_at: date,
          to: pto
        }
      }
      };
    var requester = getRequester(pfrom);
    
    if (!requester) {
      data.ticket.requester = {
        "name": "+14804283978",
        "phone": "+14804283978"
      };
    } else {
      data.ticket.requester_id = requester;
    }
    createTicket(JSON.stringify(data));
    scriptProperties.setProperty('LAST_CHECKED', values[i][0]);
  }else{
    Logger.log("Not creating ticket for ID: " + values[i][0]);
  }
  
} 

}

function doAllTheFun() {
  if(verifyConfig()){
    sortFormResponses();
    parseSheet();
    generateTickets();
  }
}
function verifyConfig(){
  // Check if config has 0 values
  initialize();
  var res = true;
  var error = "";
  if(user == ''){
    error += "<p>ERROR: Please add the username on the configuration page.</p>";
    res = false;
  }else if(subdomain == ''){
    error += "<p>ERROR: Please add the subdomain on the configuration page.</p>";
    res = false;
  }else if(token == ''){
    error += "<p>ERROR: Please add the token on the configuration page.</p>";
    res = false;
  }else if(!doTestAPICall()){
    error += "<p>ERROR :The API details are incorrect, please review them before continuing.</p>";
    res = false;
  }
  
  if(!res){
   doShowErrors(error);
  }
  
  return res;
}

function doShowErrors(error){
  
  var html = HtmlService.createHtmlOutput()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('My custom sidebar')
      .setWidth(300);
  html.clear();
  html.append(error);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

function doTestAPICall() {
  var headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(user + "/token:" + token)
  };
  var url = "https://" + subdomain + ".zendesk.com/api/v2/users/me.json";
  var options = {
    'contentType': 'application/json',
    'method': 'get',
    'headers': headers
  };
  var response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  Logger.log(data);
  if (data.user.email != user) {
    //User was not authenticated
    return false;
  } else {
    //User was correctly authenticated
    return true;
  }
}

function getRequester(phonenumber) {
  var headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(user + "/token:" + token)
  };
  var url = "https://" + subdomain + ".zendesk.com/api/v2/search.json?query=type:user+ " + phonenumber;
  var options = {
    'contentType': 'application/json',
    'method': 'get',
    'headers': headers
  };
  var response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  if (data.results.length == 0) {
    //Requester does not exist, will return false
    return false;
  } else {
    //Requester will be the first result of the search
    return data.results[0].id;
  }
}

function createTicket(data) {
  var headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(user + "/token:" + token)
  };
  var url = "https://" + subdomain + ".zendesk.com/api/v2/channels/voice/tickets.json";
  var options = {
    'contentType': 'application/json',
    'method': 'POST',
    'headers': headers,
    'payload': data
  };
  
  var response = UrlFetchApp.fetch(url, options);
  
}

function onOpen() {
  var scriptProperties = PropertiesService.getScriptProperties();
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .createMenu('Zendesk Ticket Creator')
  .addItem('Configuration', 'openConfig')
  .addItem('Run!', 'doAllTheFun')
  .addToUi();
  
  // Let's detect if there are properties
  scriptProperties = PropertiesService.getScriptProperties();
  prop = scriptProperties.getProperties();
  if (isEmpty(prop)) {
    //initialize the properties as empty
    scriptProperties.setProperties({
      'SUBDOMAIN': '',
      'USER': '',
      'TOKEN': '',
      'LAST_CHECKED': '0'
    });
  }
}

function openConfig() {
  SpreadsheetApp.getUi().showModalDialog(doBuildConfigHtml().setWidth(400).setHeight(450), 'Zendesk Ticket Generator Config');
}

function onInstall(){
  onOpen();
}

function isEmpty(obj) {
  return Object.keys(obj).length === 0;
}

function doBuildConfigHtml() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var t = HtmlService.createTemplateFromFile('configuration');
  Logger.log(scriptProperties.getProperties());
  t.data = scriptProperties.getProperties();
  return t.evaluate().setSandboxMode(HtmlService.SandboxMode.NATIVE);
}

function saveConfig(parameters) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    'SUBDOMAIN': parameters.inputSubdomain,
    'USER': parameters.inputUser,
    'TOKEN': parameters.inputToken,
    'LAST_CHECKED': parameters.inputID
  });
  google.script.host.close();
}

function sortFormResponses() {
  // change name of sheet to your sheet name
  initialize();
  var lastCol = page.getLastColumn();
  var lastRow = page.getLastRow();
  // assumes headers in row 1
  var r = page.getRange(2, 1, lastRow - 1, lastCol);
  // Note the use of an array
  r.sort([{ column: 1, ascending: true }]);

}