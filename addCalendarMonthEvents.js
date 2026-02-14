// function addCalendarMonthEvents() {
//   SpreadsheetApp.flush(); // シートの再描画
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB');

//   const startTime = new Date(sheet.getRange('G2').getValue());
//   const endTime = new Date(startTime);
//   endTime.setMonth(endTime.getMonth() + 1);

//   const dateValues = getCalenderAction(startTime, endTime);

//   const lastRow = sheet.getLastRow();
//   sheet.getRange(lastRow + 1, 1, dateValues.length, dateValues[0].length).setValues(dateValues);

// }

// function getCalenderAction(startTime, endTime) {
//   const key = PropertiesService.getScriptProperties().getProperty("CALENDAR_ID");
//   const calendar = CalendarApp.getCalendarById(key);
//   const events = calendar.getEvents(startTime, endTime);

//   const values = [];
//   for (const event of events) {
//     const record = [
//       event.getTitle(),
//       event.getStartTime(),
//       event.getEndTime(),
//       `=INDIRECT("RC[-1]",FALSE)-INDIRECT("RC[-2]",FALSE)`,
//       `=IFERROR(SUBSTITUTE(LEFT(INDIRECT("RC[-4]", FALSE), FIND("】", INDIRECT("RC[-4]", FALSE)) - 1), "【", ""), "")`,
//       `=Int(INDIRECT("RC[-4]",FALSE))`
//     ];
//     values.push(record);
//   }
//   return values
// }