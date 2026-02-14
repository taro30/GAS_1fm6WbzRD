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
        const dateStr = `${Utilities.formatDate(thisWeekRange.start, 'JST', 'yyyy/MM/dd')} - ${Utilities.formatDate(thisWeekRange.end, 'JST', 'yyyy/MM/dd')}`;
        // Emoji usage is restricted to avoid ??? issues in some environments.
        const subject = `[週次レポート] ${dateStr} カレンダー集計`;

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
 */
function aggregateWeeklyStats(sheet, start, end) {
    const data = sheet.getDataRange().getValues();
    const stats = {};

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length < 6) continue;
        const title = String(row[0]);
        const eventDate = new Date(row[5]);
        const durationRaw = row[3];

        const match = title.match(/【(.*?)】/);
        if (match && eventDate >= start && eventDate <= end) {
            const category = match[1];
            let durationHours = 0;
            if (durationRaw instanceof Date) {
                durationHours = durationRaw.getHours() + (durationRaw.getMinutes() / 60);
            } else if (typeof durationRaw === 'number') {
                durationHours = durationRaw * 24;
            }

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
    const tmpSheetName = `ChartData_${new Date().getTime()}`;
    const tmpSheet = ss.insertSheet(tmpSheetName);

    try {
        const data = sheet.getDataRange().getValues();
        const categoriesSet = new Set();
        const dailyMap = {}; // { MM/dd: { category: duration } }

        // 1. Data Collection
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row.length < 6) continue;
            const title = String(row[0]);
            const eventDate = new Date(row[5]);
            const durationRaw = row[3];

            if (isNaN(eventDate.getTime())) continue;

            const match = title.match(/【(.*?)】/);
            if (match && eventDate >= start && eventDate <= end) {
                const cat = match[1];
                const dateKey = Utilities.formatDate(eventDate, 'JST', 'MM/dd');

                let dur = 0;
                if (durationRaw instanceof Date) {
                    dur = durationRaw.getHours() + (durationRaw.getMinutes() / 60);
                } else if (typeof durationRaw === 'number') {
                    dur = durationRaw * 24;
                }

                categoriesSet.add(cat);
                if (!dailyMap[dateKey]) dailyMap[dateKey] = {};
                dailyMap[dateKey][cat] = (dailyMap[dateKey][cat] || 0) + dur;
            }
        }

        const categories = Array.from(categoriesSet).sort();
        if (categories.length === 0) return null;

        const header = ["日付", ...categories];
        const rows = [];

        // 2. Filling 7 days grid
        for (let i = 0; i < 7; i++) {
            const currentDay = new Date(start.getTime());
            currentDay.setDate(currentDay.getDate() + i);
            const ds = Utilities.formatDate(currentDay, 'JST', 'MM/dd');
            const row = [ds];
            categories.forEach(cat => {
                row.push(dailyMap[ds] ? (dailyMap[ds][cat] || 0) : 0);
            });
            rows.push(row);
        }

        // Write to temporary sheet
        tmpSheet.getRange(1, 1, 1, header.length).setValues([header]);
        tmpSheet.getRange(2, 1, rows.length, header.length).setValues(rows);

        // 3. Create Chart
        const chart = tmpSheet.newChart()
            .setChartType(Charts.ChartType.COLUMN)
            .addRange(tmpSheet.getRange(1, 1, rows.length + 1, header.length))
            .setOption('isStacked', true)
            .setOption('title', '日次カテゴリー別時間配分 (h)')
            .setOption('hAxis', { title: '日付' })
            .setOption('vAxis', { title: '時間 (h)' })
            .setOption('width', 600)
            .setOption('height', 400)
            .setOption('legend', { position: 'bottom' })
            .build();

        return chart.getAs('image/png').setName('daily_chart.png');

    } catch (err) {
        console.error("Chart Error: " + err.message);
        return null;
    } finally {
        ss.deleteSheet(tmpSheet);
    }
}

/**
 * Calls Gemini for detailed commentary.
 */
function getGeminiCommentary(comparison) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "※AI寸評はAPIキー未設定のためスキップされました。";

    const endpoint = `https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const prompt = `あなたは生活習慣の分析をサポートするライフログ専門のAIです。
以下のカレンダー集計データを見て、傾向と今後の示唆を【300文字〜400文字程度】で、詳しく日本語で述べてください。

ポイント：
- 活動件数と、特に「時間（h）」の増減に着目してください。
- 各カテゴリーのバランス（仕事、休憩、自己研鑽など）から、現在のライフスタイルの質を分析してください。
- 改善点や、継続すべき良い傾向があれば優しくアドバイスしてください。
- 「〜の傾向が見られます」「〜が示唆されます」といった丁寧なトーンで記述してください。

集計データ:
${JSON.stringify(comparison)}
`;

    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(endpoint, options);
    const json = JSON.parse(response.getContentText());

    return (json.candidates && json.candidates[0].content.parts[0].text) ? json.candidates[0].content.parts[0].text.trim() : "寸評の取得に失敗しました。";
}

/**
 * Sends formatted HTML Email.
 */
function sendHtmlEmail(subject, comparison, aiCommentary, chartBlob, dateRange) {
    const userEmail = Session.getActiveUser().getEmail();

    let tableRows = "";
    comparison.forEach(item => {
        const diffDur = item.diffDuration.toFixed(1);
        const diffColor = item.diffDuration > 0 ? "#4285F4" : (item.diffDuration < 0 ? "#EA4335" : "#333");
        const durStr = item.currentDuration.toFixed(1);

        tableRows += `
      <tr>
        <td style="border: 1px solid #ddd; padding: 10px; font-weight: bold;">【${item.category}】</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${item.currentCount}回</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${durStr}h</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center; color: ${diffColor};">${item.diffDuration > 0 ? '+' : ''}${diffDur}h</td>
      </tr>`;
    });

    // Handle inline image
    const inlineImages = {};
    let chartHtml = "";
    if (chartBlob) {
        inlineImages['chart'] = chartBlob;
        chartHtml = `<div style="margin: 30px 0; text-align: center;"><img src="cid:chart" style="width: 100%; max-width: 600px; border: 1px solid #eee; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);" /></div>`;
    }

    const htmlBody = `
    <div style="font-family: 'Hiragino Kaku Gothic ProN', 'Meiryo', sans-serif; max-width: 650px; margin: 0 auto; padding: 20px; border: 1px solid #f0f0f0; background-color: #ffffff; color: #333;">
      <h2 style="color: #4285F4; border-left: 6px solid #4285F4; padding: 10px 15px; background-color: #f8f9fa;">週次ライフログ・レポート</h2>
      <p style="margin: 20px 0; font-size: 1.1em;">
        <strong>[対象期間]</strong> ${dateRange}
      </p>
      
      <h3 style="border-bottom: 2px solid #eee; padding-bottom: 5px; margin-top: 30px;">(集計) カテゴリー別統計</h3>
      <table style="border-collapse: collapse; width: 100%; margin-top: 15px; font-size: 0.95em;">
        <thead>
          <tr style="background-color: #f2f2f2;">
            <th style="border: 1px solid #ddd; padding: 12px; text-align: left;">カテゴリー</th>
            <th style="border: 1px solid #ddd; padding: 12px;">回数</th>
            <th style="border: 1px solid #ddd; padding: 12px;">累計時間</th>
            <th style="border: 1px solid #ddd; padding: 12px;">前週比</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>

      ${chartHtml}

      <h3 style="border-bottom: 2px solid #eee; padding-bottom: 5px; margin-top: 40px;">AI Insight (Gemini)</h3>
      <div style="background-color: #f1f3f4; padding: 25px; border-radius: 12px; border-left: 8px solid #4285F4; margin-top: 15px; line-height: 1.8; font-size: 1em; white-space: pre-wrap;">
${aiCommentary}
      </div>

      <footer style="margin-top: 50px; padding-top: 20px; border-top: 1px solid #eee; font-size: 0.85em; color: #777; text-align: center;">
        このメールは自動送信されています。本日も充実した一日を過ごしましょう。<br>
        &copy; Living Log Service
      </footer>
    </div>
  `;

    GmailApp.sendEmail(userEmail, subject, "", {
        htmlBody: htmlBody,
        inlineImages: inlineImages,
        attachments: chartBlob ? [chartBlob] : []
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
