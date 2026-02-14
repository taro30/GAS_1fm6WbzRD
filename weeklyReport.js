/**
 * Weekly Report Module
 * Calculates calendar statistics (Counts & Durations), generates charts, gets AI commentary, and sends emails.
 */

/**
 * Main function to execute the weekly report process.
 */
function sendWeeklyReport() {
    try {
        const now = new Date();
        // Calculate date ranges: This Week (Mon-Sun) and Last Week (Mon-Sun)
        const thisWeekRange = getWeeklyDateRange(now, 0);
        const lastWeekRange = getWeeklyDateRange(now, -1);

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB');
        if (!sheet) throw new Error("'DB' sheet not found.");

        // 1. Aggregate Data (Count AND Duration)
        const thisWeekStats = aggregateWeeklyStats(sheet, thisWeekRange.start, thisWeekRange.end);
        const lastWeekStats = aggregateWeeklyStats(sheet, lastWeekRange.start, lastWeekRange.end);

        // 2. Compare Data
        const comparison = compareWeeklyStats(thisWeekStats, lastWeekStats);

        // 3. Generate Chart (Stacked Bar by Day)
        const chartBlob = generateStackedDailyChart(sheet, thisWeekRange.start, thisWeekRange.end);

        // 4. Get Gemini Commentary
        const aiCommentary = getGeminiCommentary(comparison);

        // 5. Send Email
        const dateStr = `${Utilities.formatDate(thisWeekRange.start, 'JST', 'yyyy/MM/dd')}–${Utilities.formatDate(thisWeekRange.end, 'JST', 'yyyy/MM/dd')}`;
        const subject = `【週次レポート】${dateStr} カレンダー集計`;

        sendHtmlEmail(subject, comparison, aiCommentary, chartBlob, dateStr);

    } catch (e) {
        console.error(`Weekly Report Failed: ${e.message}`);
        throw e;
    }
}

/**
 * Calculates start and end dates for a week (Monday to Sunday).
 */
function getWeeklyDateRange(refDate, offsetWeeks) {
    const date = new Date(refDate.getTime());
    const day = date.getDay(); // 0(Sun) - 6(Sat)
    const diffToMon = (day === 0 ? -6 : 1 - day);

    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate() + diffToMon + (offsetWeeks * 7), 0, 0, 0);
    const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59, 999);

    return { start, end };
}

/**
 * Aggregates category counts and cumulative hours from 'DB' sheet.
 * Filters for categories enclosed in 【】.
 * @return {Object} { categoryName: { count: number, durationHours: number } }
 */
function aggregateWeeklyStats(sheet, start, end) {
    const data = sheet.getDataRange().getValues();
    const stats = {};

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const title = String(row[0]); // Col A: Title
        const eventDate = new Date(row[5]); // Col F: Date
        const durationRaw = row[3]; // Col D: Duration (Time object or number)

        // Extract category from 【】
        const match = title.match(/【(.*?)】/);
        if (match && eventDate >= start && eventDate <= end) {
            const category = match[1];
            const durationHours = (durationRaw instanceof Date)
                ? (durationRaw.getHours() + durationRaw.getMinutes() / 60 + durationRaw.getSeconds() / 3600)
                : (typeof durationRaw === 'number' ? durationRaw * 24 : 0); // Serial to hours if numeric

            if (!stats[category]) {
                stats[category] = { count: 0, duration: 0 };
            }
            stats[category].count += 1;
            stats[category].duration += durationHours;
        }
    }
    return stats;
}

/**
 * Compares current and previous week stats.
 */
function compareWeeklyStats(current, previous) {
    const allCategories = new Set([...Object.keys(current), ...Object.keys(previous)]);
    const comparison = [];

    allCategories.forEach(cat => {
        const cur = current[cat] || { count: 0, duration: 0 };
        const pre = previous[cat] || { count: 0, duration: 0 };
        comparison.push({
            category: cat,
            currentCount: cur.count,
            currentDuration: cur.duration,
            previousCount: pre.count,
            previousDuration: pre.duration,
            diffCount: cur.count - pre.count,
            diffDuration: cur.duration - pre.duration
        });
    });

    return comparison.sort((a, b) => b.currentDuration - a.currentDuration);
}

/**
 * Generates a Stacked Column Chart based on daily data in the week.
 */
function generateStackedDailyChart(sheet, start, end) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmpSheet = ss.insertSheet('_tmp_chart');

    try {
        const data = sheet.getDataRange().getValues();
        const categoriesSet = new Set();
        const dailyData = {}; // { YYYY/MM/DD: { category: duration } }

        // 1. Collect Categories and Daily Totals
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const title = String(row[0]);
            const eventDate = new Date(row[5]);
            const durationRaw = row[3];

            const match = title.match(/【(.*?)】/);
            if (match && eventDate >= start && eventDate <= end) {
                const cat = match[1];
                const dateStr = Utilities.formatDate(eventDate, 'JST', 'MM/dd');
                const duration = (durationRaw instanceof Date)
                    ? (durationRaw.getHours() + durationRaw.getMinutes() / 60)
                    : (typeof durationRaw === 'number' ? durationRaw * 24 : 0);

                categoriesSet.add(cat);
                if (!dailyData[dateStr]) dailyData[dateStr] = {};
                dailyData[dateStr][cat] = (dailyData[dateStr][cat] || 0) + duration;
            }
        }

        const categories = Array.from(categoriesSet);
        const header = ["Date", ...categories];
        const rows = [];

        // 2. Prepare grid for chart (7 days)
        for (let i = 0; i < 7; i++) {
            const d = new Date(start.getTime());
            d.setDate(d.getDate() + i);
            const ds = Utilities.formatDate(d, 'JST', 'MM/dd');
            const row = [ds];
            categories.forEach(cat => {
                row.push(dailyData[ds] ? (dailyData[ds][cat] || 0) : 0);
            });
            rows.push(row);
        }

        tmpSheet.getRange(1, 1, 1, header.length).setValues([header]);
        if (rows.length > 0) {
            tmpSheet.getRange(2, 1, rows.length, header.length).setValues(rows);
        }

        // 3. Create Stacked Column Chart
        const chart = tmpSheet.newChart()
            .setChartType(Charts.ChartType.COLUMN)
            .addRange(tmpSheet.getRange(1, 1, rows.length + 1, header.length))
            .setOption('isStacked', true)
            .setOption('title', 'Weekly Activity Distribution (Hours)')
            .setOption('hAxis.title', 'Day')
            .setOption('vAxis.title', 'Hours')
            .setOption('legend', { position: 'right' })
            .setOption('backgroundColor', '#fdfdfd')
            .build();

        return chart.getAs('image/png').setName('stacked_weekly_chart.png');

    } finally {
        ss.deleteSheet(tmpSheet);
    }
}

/**
 * Calls Gemini for commentary.
 */
function getGeminiCommentary(comparison) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "※AI寸評はAPIキー未設定のためスキップされました。";

    const endpoint = `https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const prompt = `あなたは生活習慣の分析をサポートするAIです。
以下のカレンダー集計データ（今週の件数・累計時間・前週比）を見て、傾向や示唆を100文字以内の日本語で述べてください。
断定や強い評価は避け、「〜の傾向が見られます」「〜が示唆されます」といった表現を使ってください。

データ:
${JSON.stringify(comparison)}
`;

    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(endpoint, options);
    const json = JSON.parse(response.getContentText());

    return (json.candidates && json.candidates[0].content.parts[0].text) ? json.candidates[0].content.parts[0].text.trim() : "寸評の取得に失敗しました。";
}

/**
 * Sends HTML Email.
 */
function sendHtmlEmail(subject, comparison, aiCommentary, chartBlob, dateRange) {
    const userEmail = Session.getActiveUser().getEmail();

    let tableRows = "";
    comparison.forEach(item => {
        const diffDur = item.diffDuration.toFixed(1);
        const diffColor = item.diffDuration > 0 ? "blue" : (item.diffDuration < 0 ? "red" : "black");
        const durStr = item.currentDuration.toFixed(1);

        tableRows += `
      <tr>
        <td style="border: 1px solid #ddd; padding: 8px;">【${item.category}】</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">${item.currentCount}回</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">${durStr}h</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: center; color: ${diffColor};">${diffDur > 0 ? '+' + diffDur : diffDur}h</td>
      </tr>`;
    });

    const htmlBody = `
    <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; max-width: 650px; color: #333;">
      <h2 style="color: #4285F4; border-bottom: 2px solid #4285F4; padding-bottom: 10px;">週次ライフログ・レポート</h2>
      <p style="font-weight: bold;">📅 対象期間: <span style="color: #555;">${dateRange}</span></p>
      
      <h3>📊 今週の活動サマリー</h3>
      <table style="border-collapse: collapse; width: 100%; border: 1px solid #ddd;">
        <thead>
          <tr style="background-color: #f8f9fa; border-bottom: 2px solid #ddd;">
            <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">カテゴリー</th>
            <th style="padding: 10px; border: 1px solid #ddd;">回数</th>
            <th style="padding: 10px; border: 1px solid #ddd;">累計時間</th>
            <th style="padding: 10px; border: 1px solid #ddd;">前週比(時間)</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>

      <div style="margin-top: 25px;">
        <img src="cid:chart" style="width: 100%; max-width: 600px; border: 1px solid #eee; border-radius: 8px;" />
      </div>

      <h3 style="margin-top: 30px; display: flex; align-items: center;">
        <span style="font-size: 1.5em; margin-right: 10px;">💡</span> AI Insight (Gemini)
      </h3>
      <div style="background-color: #f1f3f4; padding: 20px; border-radius: 8px; border-left: 6px solid #4285F4; line-height: 1.6;">
        ${aiCommentary}
      </div>

      <footer style="margin-top: 40px; padding-top: 10px; border-top: 1px solid #eee; font-size: 0.8em; color: #888; text-align: center;">
        このメールは自動送信されています。今日も良い一日を！
      </footer>
    </div>
  `;

    GmailApp.sendEmail(userEmail, subject, "", {
        htmlBody: htmlBody,
        inlineImages: { chart: chartBlob },
        attachments: [chartBlob]
    });
}

/**
 * Creates trigger.
 */
function createWeeklyTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => { if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t); });

    ScriptApp.newTrigger('sendWeeklyReport')
        .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();
    console.log("Weekly trigger created.");
}
