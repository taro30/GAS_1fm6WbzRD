// function addCalendarSelectDayEvents() {
//   SpreadsheetApp.flush(); // シートの再描画
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB');

//   const ui = SpreadsheetApp.getUi();
//   const response = ui.prompt('○○○○/○○/○○形式で日付入力ください', ui.ButtonSet.OK);
//   const inputDate = response.getResponseText();
  
//   if (inputDate) {
//     const startTime = new Date(inputDate);
//     startTime.setHours(0);
//     startTime.setMinutes(0);
//     startTime.setSeconds(0);

//     const endTime = new Date((startTime));
//     endTime.setDate(endTime.getDate() + 1);

//     const dateValues = getCalenderAction(startTime, endTime);

//     if (dateValues.length > 0) {
//       const lastRow = sheet.getLastRow();
//       sheet.getRange(lastRow + 1, 1, dateValues.length, dateValues[0].length).setValues(dateValues);
//     }
//   } else {
//     ui.alert('Invalid date. Please try again.');
//   }
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
//       `=IF(ABS((INDIRECT("RC[-1]",FALSE)-INDIRECT("RC[-2]",FALSE)))=1,"",INDIRECT("RC[-1]",FALSE)-INDIRECT("RC[-2]",FALSE))`,
//       `=IFERROR(MID(INDIRECT("RC[-4]",FALSE),FIND("【",INDIRECT("RC[-4]",FALSE))+1,FIND("】",INDIRECT("RC[-4]",FALSE))-FIND("【",INDIRECT("RC[-4]",FALSE))-1),"")`,
//       `=Int(INDIRECT("RC[-4]",FALSE))`
//     ];
//     values.push(record);
//   }
//   return values;
// }

