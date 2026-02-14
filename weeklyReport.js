/**
 * Weekly Report Module
 * Calculates calendar statistics, generates charts, gets AI commentary, and sends emails.
 */

/**
 * Main function to execute the weekly report process.
 * Recommended to be triggered on Sundays.
 */
function sendWeeklyReport() {
    try {
        const now = new Date();
        // Calculate date ranges: This Week (Mon-Sun) and Last Week (Mon-Sun)
        const thisWeekRange = getWeeklyDateRange(now, 0); // Current week's Mon-Sun
        const lastWeekRange = getWeeklyDateRange(now, -1); // Previous week's Mon-Sun

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB');
        if (!sheet) throw new Error("'DB' sheet not found.");

        // 1. Aggregate Data
        const thisWeekStats = aggregateCategoryStats(sheet, thisWeekRange.start, thisWeekRange.end);
        const lastWeekStats = aggregateCategoryStats(sheet, lastWeekRange.start, lastWeekRange.end);

        // 2. Compare Data
        const comparison = compareWeeklyStats(thisWeekStats, lastWeekStats);

        // 3. Generate Chart
        const chartBlob = generateCategoryChart(comparison);

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
 * Calculates start and end dates for a week relative to the reference date.
 * Week is defined as Monday 0:00 to Sunday 23:59:59.
 * @param {Date} refDate Reference date
 * @param {number} offsetWeeks Number of weeks to offset (0 for current week, -1 for last week)
 * @return {Object} {start: Date, end: Date}
 */
function getWeeklyDateRange(refDate, offsetWeeks) {
    const date = new Date(refDate.getTime());
    const day = date.getDay(); // 0(Sun) - 6(Sat)
    const diffToMon = (day === 0 ? -6 : 1 - day); // Distance to this week's Monday

    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate() + diffToMon + (offsetWeeks * 7), 0, 0, 0);
    const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59, 999);

    return { start, end };
}

/**
 * Aggregates category counts from the 'DB' sheet for a given date range.
 * Assumes: Col 6 (F) is Date, Col 5 (E) is Category.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The DB sheet
 * @param {Date} start Start of range
 * @param {Date} end End of range
 * @return {Object} { categoryName: count }
 */
function aggregateCategoryStats(sheet, start, end) {
    const data = sheet.getDataRange().getValues();
    const stats = {};

    // Skip header, loop through rows
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const eventDate = new Date(row[5]); // Column F (Index 5) is Date
        const category = row[4] ? String(row[4]).trim() : ""; // Column E (Index 4) is Category

        if (category && eventDate >= start && eventDate <= end) {
            stats[category] = (stats[category] || 0) + 1;
        }
    }
    return stats;
}

/**
 * Compares two weekly stats objects.
 * @return {Array} Sorted array of objects {category, current, previous, diff}
 */
function compareWeeklyStats(current, previous) {
    const allCategories = new Set([...Object.keys(current), ...Object.keys(previous)]);
    const comparison = [];

    allCategories.forEach(cat => {
        const curVal = current[cat] || 0;
        const preVal = previous[cat] || 0;
        comparison.push({
            category: cat,
            current: curVal,
            previous: preVal,
            diff: curVal - preVal
        });
    });

    // Sort by current volume descending
    return comparison.sort((a, b) => b.current - a.current);
}

/**
 * Generates a bar chart as a PNG Blob.
 * Uses a temporary sheet for data range.
 */
function generateCategoryChart(comparison) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmpSheet = ss.insertSheet('_tmp_chart');

    try {
        // 1. Prepare data for chart
        const chartData = [["Category", "This Week", "Last Week"]];
        comparison.forEach(item => {
            chartData.push([item.category, item.current, item.previous]);
        });

        tmpSheet.getRange(1, 1, chartData.length, 3).setValues(chartData);

        // 2. Create Chart
        const chart = tmpSheet.newChart()
            .setChartType(Charts.ChartType.COLUMN)
            .addRange(tmpSheet.getRange(1, 1, chartData.length, 3))
            .setPosition(1, 5, 0, 0)
            .setOption('title', 'Category Distribution (Current vs Previous)')
            .setOption('vAxis.title', 'Count')
            .setOption('colors', ['#4285F4', '#EA4335'])
            .build();

        // 3. Get Image Blob
        return chart.getAs('image/png').setName('weekly_chart.png');

    } finally {
        ss.deleteSheet(tmpSheet);
    }
}

/**
 * Calls Gemini API to get short commentary.
 */
function getGeminiCommentary(comparison) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "※AI寸評はAPIキー未設定のためスキップされました。";

    // Use gemini-2.5-flash (Stable as of June 2025)
    const endpoint = `https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    const prompt = `あなたは生活習慣の分析をサポートするAIです。
以下のカレンダー集計データ（今週の件数と前週比）を見て、傾向や示唆を100文字以内の日本語で述べてください。
断定や強い評価は避け、「〜の傾向が見られます」「〜が示唆されます」といった表現を使ってください。

データ:
${JSON.stringify(comparison)}
`;

    const payload = {
        contents: [{ parts: [{ text: prompt }] }]
    };

    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(endpoint, options);
    const json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates[0].content.parts[0].text) {
        return json.candidates[0].content.parts[0].text.trim();
    } else {
        throw new Error(`Gemini API Error: ${response.getContentText()}`);
    }
}

/**
 * Sends HTML Email with statistics and chart attachment.
 */
function sendHtmlEmail(subject, comparison, aiCommentary, chartBlob, dateRange) {
    const userEmail = Session.getActiveUser().getEmail();

    let tableRows = "";
    comparison.forEach(item => {
        const diffText = item.diff > 0 ? `+${item.diff}` : `${item.diff}`;
        const diffColor = item.diff > 0 ? "blue" : (item.diff < 0 ? "red" : "black");
        tableRows += `
      <tr>
        <td style="border: 1px solid #ddd; padding: 8px;">${item.category}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">${item.current}</td>
        <td style="border: 1px solid #ddd; padding: 8px; text-align: center; color: ${diffColor};">${diffText}</td>
      </tr>`;
    });

    const htmlBody = `
    <div style="font-family: sans-serif; max-width: 600px;">
      <h2>週次カレンダー集計レポート</h2>
      <p>対象期間: ${dateRange}</p>
      
      <h3>📊 今週の集計結果</h3>
      <table style="border-collapse: collapse; width: 100%;">
        <thead>
          <tr style="background-color: #f2f2f2;">
            <th style="border: 1px solid #ddd; padding: 8px;">カテゴリ</th>
            <th style="border: 1px solid #ddd; padding: 8px;">今週(件)</th>
            <th style="border: 1px solid #ddd; padding: 8px;">前週比</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>

      <h3>💡 AIによる寸評 (Gemini)</h3>
      <div style="background-color: #f9f9f9; padding: 15px; border-left: 5px solid #4285F4; margin: 10px 0;">
        ${aiCommentary}
      </div>

      <p style="color: #666; font-size: 0.9em;">詳細は添付のグラフをご確認ください。</p>
    </div>
  `;

    GmailApp.sendEmail(userEmail, subject, "", {
        htmlBody: htmlBody,
        inlineImages: { chart: chartBlob },
        attachments: [chartBlob]
    });
}

/**
 * Creates a time-driven trigger for the weekly report.
 * Sets to run every Sunday around 9:00 AM.
 */
function createWeeklyTrigger() {
    // Delete existing triggers for same function to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
        if (t.getHandlerFunction() === 'sendWeeklyReport') {
            ScriptApp.deleteTrigger(t);
        }
    });

    // Create new trigger: Weekly on Sunday at 9 AM
    ScriptApp.newTrigger('sendWeeklyReport')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.SUNDAY)
        .atHour(9)
        .create();

    console.log("Weekly trigger created successfully.");
}
