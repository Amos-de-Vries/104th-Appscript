/*
@OnlyCurrentDoc
*/

var mr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
var Roster = mr.getRange("A:AI");

function everyMinutes() {
  Roster.sort({column: 1, ascending: false});
}

function onEdit() {
  Roster.sort({column: 1, ascending: false});
}
