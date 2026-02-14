/**
 * Googleカレンダーから指定された日付のイベントを取得し、
 * スプレッドシートに存在しないイベントを追加します。
 */
function addCalendarDayEvents() {
  // 保留中の変更を強制的に更新します。
  SpreadsheetApp.flush();

  // 'DB'という名前のシートを取得します。
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB');

  // G4セルの日付を対象日として取得します。
  const targetDate = new Date(sheet.getRange('G4').getValue());

  // 対象日の開始時刻を設定します（時刻は0時0分0秒）。
  const startTime = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 0, 0, 0);

  // 対象日の終了時刻を設定します（次の日の0時0分0秒）。
  const endTime = new Date(startTime);
  endTime.setDate(endTime.getDate() + 1);

  // 指定された日付のGoogleカレンダーのイベントを取得します。
  const calendarEvents = getCalenderAction(startTime, endTime);

  // スプレッドシート内の指定された日付のイベントタイトルと開始時刻のペアをセットに追加します。
  const sheetEventsSet = getSheetEventsSetForDate(sheet, startTime);

  // カレンダーのイベントのうち、スプレッドシートにまだ存在しないイベントをフィルタリングします。
  const newEvents = calendarEvents.filter(event => {
    const eventKey = createEventKey(event[0], event[1]);
    return !sheetEventsSet.has(eventKey);
  });

  // 新しいイベントがある場合、それをスプレッドシートに追加します。
  if (newEvents.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newEvents.length, newEvents[0].length).setValues(newEvents);
  }
}

/**
 * 指定された期間内のGoogleカレンダーのイベントを取得します。
 * @param {Date} startTime - 開始時刻
 * @param {Date} endTime - 終了時刻
 * @return {Array} カレンダーイベントの配列
 */
function getCalenderAction(startTime, endTime) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const calendarIDs = [scriptProperties.getProperty("CALENDAR_ID"), scriptProperties.getProperty("CALENDAR_ID2")].filter(Boolean);
  const values = [];

  // 各カレンダーIDに対してイベントを取得します。
  calendarIDs.forEach(key => {
    const calendar = CalendarApp.getCalendarById(key);
    const events = calendar.getEvents(startTime, endTime);
    events.forEach(event => {
      values.push(formatEventRecord(event));
    });
  });

  return values;
}

/**
 * Googleカレンダーのイベントからスプレッドシートのレコード用のデータを整形します。
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event - Googleカレンダーのイベント
 * @return {Array} スプレッドシートに追加するイベントデータ
 */
function formatEventRecord(event) {
  return [
    event.getTitle(), // イベントのタイトル
    event.getStartTime(), // イベントの開始時刻
    event.getEndTime(), // イベントの終了時刻
    `=IF(ABS((INDIRECT("RC[-1]",FALSE)-INDIRECT("RC[-2]",FALSE)))*24=24,"",(INDIRECT("RC[-1]",FALSE)-INDIRECT("RC[-2]",FALSE)))`,
    `=IFERROR(MID(INDIRECT("RC[-4]",FALSE),FIND("【",INDIRECT("RC[-4]",FALSE))+1,FIND("】",INDIRECT("RC[-4]",FALSE))-FIND("【",INDIRECT("RC[-4]",FALSE))-1),"")`,
    `=INT(INDIRECT("RC[-4]",FALSE))`
  ];
}

/**
 * スプレッドシートから特定の日付のイベントのタイトルと開始時刻をセットとして取得します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - スプレッドシートのシート
 * @param {Date} startTime - 対象の日付
 * @return {Set} イベントのタイトルと開始時刻のペアのセット
 */
function getSheetEventsSetForDate(sheet, startTime) {
  const dataRange = sheet.getDataRange().getValues();
  const eventsSet = new Set();

  dataRange.forEach(row => {
    const eventDate = row[1];
    if (eventDate instanceof Date && isSameDate(eventDate, startTime)) {
      const eventKey = createEventKey(row[0], row[1]);
      eventsSet.add(eventKey);
    }
  });

  return eventsSet;
}

/**
 * 与えられた2つの日付が同じ日であるかどうかを判断します。
 * @param {Date} d1 - 比較する日付1
 * @param {Date} d2 - 比較する日付2
 * @return {boolean} 2つの日付が同じ日であればtrue
 */
function isSameDate(d1, d2) {
  return d1.getDate() === d2.getDate() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getFullYear() === d2.getFullYear();
}

/**
 * イベントのタイトルと開始時刻からユニークなキーを生成します。
 * @param {string} title - イベントのタイトル
 * @param {Date} startTime - イベントの開始時刻
 * @return {string} イベントのユニークなキー
 */
function createEventKey(title, startTime) {
  return `${title}_${startTime.toISOString()}`;
}
