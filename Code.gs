/*
@OnlyCurrentDoc
*/

var mr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
var Roster = mr.getRange("A:AI");

function onOpen() {
  var ui = SpreadsheetApp.getUi();  
  initMenu(ui);
  
  formatRoster();  
}

function everyMinutes() {
  sortRoster();
}

function onEdit() {
  sortRoster();  
}

function sortRoster() {
  Roster.sort({column: 1, ascending: false});
}

function reformatForm() {
  var reformattingFormTemplate = HtmlService.createTemplateFromFile('formatMenu');
  var html = reformattingFormTemplate.evaluate();

  SpreadsheetApp.getUi().showModelessDialog(html, "Reformatting Menu")  
}

function reformatMainRoster() {
  Roster.setBorder(true, true, true, true, true, true, "Black", SpreadsheetApp.BorderStyle.SOLID)  
}

function reformatAll() {
  reformatMainRoster();
}

function initMenu(ui) {  
  
    var rosterTeamMenu = ui.createMenu('Roster Team officer+');
      var formattingSubMenu = ui.createMenu("Formatting Menu (W.I.P)");
        formattingSubMenu.addItem('Reformatting Menu', 'reformatForm');
        formattingSubMenu.addItem('Sort Main roster', 'sortRoster');
    rosterTeamMenu.addSubMenu(formattingSubMenu)        

    rosterTeamMenu.addToUi();
}

// For questions/concerns feel free to contact the author: NightWolf#6326 / Hermes
// Permissions Needed: https://www.googleapis.com/auth/script.container.ui, https://www.googleapis.com/auth/spreadsheets.currentonly
// Github: https://github.com/Amos-de-Vries/104th-Appscript (official roster: Main Branch)
