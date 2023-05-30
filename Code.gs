/*
@OnlyCurrentDoc
*/

var mr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
var Roster = mr.getRange("A:AI");

function onOpen(){
  var ui = SpreadsheetApp.getUi();
    
  initMenu(ui);    
  runChecks();
}

function everyMinutes() {
  Roster.sort({column: 1, ascending: false});   
}

function onEdit() {
  Roster.sort({column: 1, ascending: false});   
}

function runChecks() {
  checkStrikes();
}

function sortStrikes() {
  var strike_data_ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('strikeData');
  var strike_data_ws_range = strike_data_ws.getRange("A2:I");

  strike_data_ws_range.sort({column: 9, ascending: false});
}

function getCurrentSelectedValues() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeRange = activeSheet.getActiveRange().getValues();
  
  return activeRange;
}

function checkStrikes() {  
  var strike_data_ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('strikeData');
  var totalStrikeCount = strike_data_ws.getLastRow();
  var currentDate = new Date();

  for (var strikeRow = 2; strikeRow <= totalStrikeCount; strikeRow = strikeRow + 1) {

    var range = strike_data_ws.getRangeList(['E'+strikeRow, 'H'+strikeRow]);
    var rangeValue = range.getRanges().map(range => [range.getValue()]);;  
    var expirationDate = new Date(rangeValue[0]);        

    if((expirationDate <= currentDate) && (rangeValue[1]=="false")) {
      var strikeActive = strike_data_ws.getRange('I'+strikeRow);   
      strikeActive.setValue("FALSE");
    }
  }  
}

// add roster row at end of page, if eligble regardless of strikes: true, otherwise false
// if eligiblity is true, remove strikes
// optomised diffrence: no need to get: Activity
// Rank row (A), or any of the data; only need steamid
// per loop instead of 6 data cells only need 1

function qoataStrikeCheck() {
// instead of taking data through each row, take all rows / cells before hand, and then just check an premade data set (keep track of Position through strikerow);


//  instead of looking at each strike
// go through each roster row, and each time its eligible and there is a strike
// delete any strike that matches
  var strike_data_ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('strikeData');
  var totalStrikeCount = strike_data_ws.getLastRow();
  var roster_ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');

  var totalRosterCount = roster_ws.getLastRow();

  // iterate through strike rows
 
      // run for each active qoata strike once through the roster 
      for (var rosterRow = 2; rosterRow <= totalRosterCount; rosterRow = rosterRow + 1) {
        // check if qoata strike steamid equals to the steamid on the roster
        // 0=STEAMID, 1=eligibility
        var rosterRangePart = roster_ws.getRangeList(['AJ'+rosterRow, 'E'+rosterRow]);
        var rosterRangePartValue = rosterRangePart.getRanges().map(range => [range.getValue()]);;  
        //  Logger.log(rosterRangePartValue[0] + " steamid:" + rosterRangePartValue[1]);
        if(rosterRangePartValue[0] == "Eligible.id") {
           for (var strikeRow = 2; strikeRow <= totalStrikeCount; strikeRow = strikeRow + 1) {
            // get steamid, qoata strike
            var range = strike_data_ws.getRangeList(['B'+strikeRow, 'H'+strikeRow]);
            var rangeValue = range.getRanges().map(range => [range.getValue()]);;    
            var row_clear = strike_data_ws.getRange('A'+strikeRow+':'+'I'+strikeRow);   

          if ((rangeValue[1].toString() == "true") && rangeValue[0].toString() == rosterRangePartValue[1]) {  
            // Logger.log(rosterRangePartValue[0] + " steamid:" + rosterRangePartValue[1]);
            row_clear.clear();
          }
        }       
      }
  }
  showAlert("All qoata strikes of those who met their qoata have been removed")
  sortStrikes();
}


function strikeMember(strikeData){
  var formattedDate = Utilities.formatDate(new Date(), "GMT-5", "MM-dd-yyyy' 'HH:mm:ss");
  var strike_data_ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('strikeData');

  strike_data_ws.appendRow([strikeData.name, strikeData.steamID, strikeData.strikeCount, strikeData.reason, strikeData.expDate, strikeData.coSteamID, formattedDate, "FALSE", "TRUE"]);  
}

function strikeMemberForm() {
  var activeRange = getCurrentSelectedValues();
  if (!activeRange[0][1] || activeRange[0][1].length != 17) {
    activeRange[0][0] = "";
    activeRange[0][1] = "";
  } 

  var strikeFormTemplate = HtmlService.createTemplateFromFile('StrikeForm');
  strikeFormTemplate.activeRangeHTML = activeRange;
  var html = strikeFormTemplate.evaluate();

  SpreadsheetApp.getUi().showModelessDialog(html, "Strike 104th Member")  
}

function strikeMemberQoata(strikeData) {
  var formattedDate = Utilities.formatDate(new Date(), "GMT-5", "MM-dd-yyyy' 'HH:mm:ss");
  var strike_data_ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('strikeData');

  strike_data_ws.appendRow([strikeData.name, strikeData.steamID, "1", "Didn't Meet Qoata", "0",  strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]); 
}


function strikeMemberQoataForm() {
  var activeRange = getCurrentSelectedValues();
  if (!activeRange[0][1] || activeRange[0][1].length != 17) {
    activeRange[0][0] = "";
    activeRange[0][1] = "";
  } 


  var qoataStrikeFormTemplate = HtmlService.createTemplateFromFile('QoataStrikeForm');
  qoataStrikeFormTemplate.activeRangeHTML = activeRange;  
  var html = qoataStrikeFormTemplate.evaluate();

  SpreadsheetApp.getUi().showModelessDialog(html, "Qoata Strike Member")  
}
    
function strikeMemberQoataAll(strikeData) {
  var strike_data_ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('strikeData');
  var roster_ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
  var totalRosterCount = roster_ws.getLastRow();
  var formattedDate = Utilities.formatDate(new Date(), "GMT-5", "MM-dd-yyyy' 'HH:mm:ss");
  // run through each person on the roster
  for (var rosterRow = 2; rosterRow <= totalRosterCount; rosterRow = rosterRow + 1) {
    // check if qoata strike steamid equals to the steamid on the roster
    // 0=RANKVALUE,1=STEAMID,2=WEEKLYPROD,3=WEEKLYREC,4=WEEKLYTR,5=NAME, 6=ACTIVITY, 7='Exempt.id'
    var rosterRange = roster_ws.getRangeList(['A'+rosterRow, 'E'+rosterRow, 'J'+rosterRow, 'L'+rosterRow, 'N'+rosterRow,'D'+rosterRow, 'F'+rosterRow, 'AJ'+rosterRow]);
    var rosterRangeValue = rosterRange.getRanges().map(range => [range.getValue()]);;          

    if((rosterRangeValue[1] > 0) && (rosterRangeValue[6] != "LOA/ROA") && (rosterRangeValue[7] != "Exempt.id")) {
        
      // XO+ (no qoata)
      if(rosterRangeValue[0] > 19000) {
      

      // check MAJ - COL  
      } else if(rosterRangeValue[0] > 16000) {               
        if(rosterRangeValue[2] < 5) {
            strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]); 
        } else if(rosterRangeValue[4] < 2) {
             strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]); 
          }
        

      // check 2ndLT - CPT
      } else if(rosterRangeValue[0] >= 13000) {
        if(rosterRangeValue[2] < 4) {
          strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]);         
        } else if (rosterRangeValue[4] < 1) {
          strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]);
        }
        
      // check SMB 
      } else if(rosterRangeValue[0] >= 12000) {
        if(rosterRangeValue[2] < 6) {
          strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]);         
        } else if (rosterRangeValue[4] < 1) {
          strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]);
        }
        

      // check 1SG - CSM
      } else if (rosterRangeValue[0] >= 9000) {
        if(rosterRangeValue[2] < 3) {      
          strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]); 
        } else if (rosterRangeValue[3] < 1) {
          strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]);
        }
        
      
      // Check SGT - MSG
      } else if (rosterRangeValue[0] >= 5000) {
        if(rosterRangeValue[2] < 2) {
          strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]); 
          } else if (rosterRangeValue[3] < 1) {    
          strike_data_ws.appendRow([rosterRangeValue[5].toString(), rosterRangeValue[1].toString(), "1", "Didn't Meet Qoata", "0", strikeData.coSteamID, formattedDate, "TRUE", "TRUE"]); 
        }       
      
      // Enlisted (no qoata)            
      } else if (rosterRangeValue[0] <= 5000) { } 
    }
  }
  SpreadsheetApp.getUi().alert("All 104th without reached qoata have been striked!")
}

function strikeMemberQoataAllForm() {

  var qoataStrikeAllFormTemplate = HtmlService.createTemplateFromFile('QoataStrikeAllForm');  
  var html = qoataStrikeAllFormTemplate.evaluate();

  SpreadsheetApp.getUi().showModelessDialog(html, "Qoata Strike All")  
}

function showAlert(message) {
  SpreadsheetApp.getUi().alert(message);
}

function manage_HWL() {
  // create menu
  
}

function initMenu(ui) {  

  var mainMenu = ui.createMenu('104th Officer+');    

    var teamManagementSubMenu = ui.createMenu("104th Team Management");
      teamManagementSubMenu.addItem('Manage Personnel Management', 'manage_PM');
      teamManagementSubMenu.addItem('Manage Specialty Management', 'manage_SM'); 
    mainMenu.addSubMenu(teamManagementSubMenu);

    var memberModerationSubMenu = ui.createMenu("104th Member Moderation");    
      memberModerationSubMenu.addItem('Default Strike', 'strikeMemberForm');
      memberModerationSubMenu.addItem('Qoata Strike Single', 'strikeMemberQoataForm');
      memberModerationSubMenu.addItem('Qoata Strike All', 'strikeMemberQoataAllForm');
      // memberModerationSubMenu.addItem('Check Expired Strikes', 'checkStrikes');      
      memberModerationSubMenu.addItem('Check All Qoata Strikes', 'qoataStrikeCheck');    
      
    mainMenu.addSubMenu(memberModerationSubMenu);
    
    var subbatallionSubMenu = ui.createMenu("104th Subbatallion Management");
      subbatallionSubMenu.addItem('Manage Howlers', 'manage_HWL');
      subbatallionSubMenu.addItem('Manage Wolfpack', 'manage_WP'); 
    mainMenu.addSubMenu(subbatallionSubMenu);

  mainMenu.addToUi();

  var rosterTeamMenu = ui.createMenu('Roster Team officer+');

    var formattingSubMenu = ui.createMenu("Formatting Menu");
      formattingSubMenu.addItem('reformat main roster', 'format_MR');
    rosterTeamMenu.addSubMenu(formattingSubMenu)        

  rosterTeamMenu.addToUi();

} 
