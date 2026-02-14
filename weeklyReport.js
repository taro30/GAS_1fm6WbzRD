/**
 * 週次レポート生成モジュール
 * カレンダーの活動件数・時間を集計し、GeminiによるAI寸評とグラフを添えてメール送信します。
 */

/**
 * 週次レポート送信のメイン実行関数
 */
function sendWeeklyReport() {
    try {
        const now = new Date();
        const thisWeekRange = getWeeklyDateRange(now, 0);
        const lastWeekRange = getWeeklyDateRange(now, -1);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('DB');
        if (!sheet) throw new Error("'DB' シートが見つかりません。");

        // 1. 全データを一度だけ読み込み（効率化）
        const allData = sheet.getDataRange().getValues();
        if (allData.length <= 1) {
            console.warn("データがありません。");
            return;
        }

        // 2. イベントのフィルタリングと取得
        const thisWeekEvents = filterEvents(allData, thisWeekRange.start, thisWeekRange.end);
        const lastWeekEvents = filterEvents(allData, lastWeekRange.start, lastWeekRange.end);

        // 3. 統計集計（テーブル・AI用）
        const thisWeekStats = aggregateStatsFromEvents(thisWeekEvents);
        const lastWeekStats = aggregateStatsFromEvents(lastWeekEvents);
        const comparison = compareWeeklyStats(thisWeekStats, lastWeekStats);

        // 4. グラフ用データの生成とBlob化
        const chartBlob = generateStackedDailyChartFromEvents(thisWeekEvents, thisWeekRange.start);

        // 5. Gemini AIによる詳細寸評の取得
        const aiCommentary = getGeminiCommentary(comparison);

        // 6. メールの送信
        const dateStr = `${formatDate(thisWeekRange.start)} - ${formatDate(thisWeekRange.end)}`;
        const subject = `[週次レポート] ${dateStr} カレンダー集計`;

        sendHtmlEmail(subject, comparison, aiCommentary, chartBlob, dateStr);

    } catch (e) {
        console.error(`週次レポート失敗: ${e.message}`);
        throw e;
    }
}

/**
 * 期間内のイベントを抽出
 */
function filterEvents(allData, start, end) {
    const events = [];
    for (let i = 1; i < allData.length; i++) {
        const row = allData[i];
        if (row.length < 6) continue;
        const eventDate = new Date(row[5]); // F列: 日付
        if (!isNaN(eventDate.getTime()) && eventDate >= start && eventDate <= end) {
            events.push({
                title: String(row[0]),     // A列: タイトル
                duration: Number(row[3]), // D列: 所要時間 (シリアル値)
                date: eventDate
            });
        }
    }
    return events;
}

/**
 * イベント群からカテゴリー統計を集計
 */
function aggregateStatsFromEvents(events) {
    const stats = {};
    events.forEach(ev => {
        const match = ev.title.match(/【(.*?)】/);
        if (match) {
            const category = match[1];
            if (!stats[category]) stats[category] = { count: 0, duration: 0 };
            stats[category].count += 1;
            stats[category].duration += (ev.duration * 24); // シリアル値を時間に変換
        }
    });
    return stats;
}

/**
 * 毎日ごとの積み上げグラフを生成
 */
function generateStackedDailyChartFromEvents(events, weekStart) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const categoriesSet = new Set();
    const dailyMap = {}; // { MM/dd: { category: duration } }

    events.forEach(ev => {
        const match = ev.title.match(/【(.*?)】/);
        if (match) {
            const cat = match[1];
            const dateKey = Utilities.formatDate(ev.date, 'JST', 'MM/dd');
            categoriesSet.add(cat);
            if (!dailyMap[dateKey]) dailyMap[dateKey] = {};
            dailyMap[dateKey][cat] = (dailyMap[dateKey][cat] || 0) + (ev.duration * 24);
        }
    });

    const categories = Array.from(categoriesSet).sort();
    if (categories.length === 0) return null;

    const tmpSheetName = `ChartData_${new Date().getTime()}`;
    const tmpSheet = ss.insertSheet(tmpSheetName);

    try {
        // カラム数が足りない場合の「Out of bounds」を防ぐため、十分なサイズを確保
        if (categories.length + 1 > 26) {
            tmpSheet.insertColumnsAfter(26, (categories.length + 1) - 26);
        }

        const header = ["日付", ...categories];
        const rows = [];
        for (let i = 0; i < 7; i++) {
            const day = new Date(weekStart.getTime());
            day.setDate(day.getDate() + i);
            const ds = Utilities.formatDate(day, 'JST', 'MM/dd');
            const row = [ds];
            categories.forEach(cat => row.push(dailyMap[ds] ? (dailyMap[ds][cat] || 0) : 0));
            rows.push(row);
        }

        tmpSheet.getRange(1, 1, 1, header.length).setValues([header]);
        tmpSheet.getRange(2, 1, rows.length, header.length).setValues(rows);

        const chart = tmpSheet.newChart()
            .setChartType(Charts.ChartType.COLUMN)
            .addRange(tmpSheet.getRange(1, 1, rows.length + 1, header.length))
            .setOption('isStacked', true)
            .setOption('title', '日次活動時間配分 (h)')
            .setOption('hAxis.title', '日付')
            .setOption('vAxis.title', '時間 (h)')
            .setOption('width', 600)
            .setOption('height', 400)
            .setOption('legend', { position: 'right' })
            .build();

        return chart.getAs('image/png').setName('daily_activity_chart.png');
    } finally {
        ss.deleteSheet(tmpSheet);
    }
}

/**
 * Gemini AIによる詳細寸評の取得
 */
function getGeminiCommentary(comparison) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "※AI寸評はAPIキー未設定のためスキップされました。";

    const endpoint = `https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const prompt = `あなたはライフログ分析の専門家AIです。
以下の週次集計データを詳細に分析し、ユーザーの今週の生活リズムについて【300文字〜400文字程度】で詳しく日本語でアドバイスしてください。

分析の視点：
- 「時間（h）」の増減に注目し、何に時間を奪われ、どこで時間を捻出したか考察してください。
- 各活動カテゴリーのバランス（仕事、休憩、インプット、生活など）から、理想的なライフスタイルに近づいているか評価してください。
- 来週に向けて、具体的な改善提案や応援メッセージを添えてください。
- 丁寧で温かみのあるトーン（〜の傾向が見られます、〜が示唆されます、といった表現）を使用してください。

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
 * 統計比較
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
            previousDuration: pre.duration,
            diffDuration: cur.duration - pre.duration
        });
    });
    return comparison.sort((a, b) => b.currentDuration - a.currentDuration);
}

/**
 * HTMLメール送信
 */
function sendHtmlEmail(subject, comparison, aiCommentary, chartBlob, dateRange) {
    const userEmail = Session.getActiveUser().getEmail();
    let tableRows = "";
    comparison.forEach(item => {
        const diff = item.diffDuration.toFixed(1);
        const diffColor = item.diffDuration > 0 ? "#4285F4" : (item.diffDuration < 0 ? "#EA4335" : "#333");
        tableRows += `
      <tr>
        <td style="border: 1px solid #ddd; padding: 10px; font-weight: bold;">【${item.category}】</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${item.currentCount}回</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${item.currentDuration.toFixed(1)}h</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center; color: ${diffColor};">${item.diffDuration > 0 ? '+' : ''}${diff}h</td>
      </tr>`;
    });

    const inlineImages = {};
    let chartHtml = "";
    if (chartBlob) {
        inlineImages['chart'] = chartBlob;
        chartHtml = `<div style="margin: 30px 0; text-align: center;"><img src="cid:chart" style="width: 100%; max-width: 600px; border-radius: 12px; border: 1px solid #eee;" /></div>`;
    }

    const htmlBody = `
    <div style="font-family: sans-serif; max-width: 650px; margin: 0 auto; padding: 25px; background-color: #ffffff; color: #333; line-height: 1.6;">
      <h2 style="color: #4285F4; border-bottom: 3px solid #4285F4; padding-bottom: 10px;">週次ライフログ・レポート</h2>
      <p style="font-size: 1.1em; margin: 20px 0;">[対象期間] ${dateRange}</p>
      
      <h3 style="background-color: #f8f9fa; padding: 10px; border-left: 5px solid #ccc;">(集計) カテゴリー別実績</h3>
      <table style="border-collapse: collapse; width: 100%; margin-top: 15px;">
        <thead>
          <tr style="background-color: #eee;">
            <th style="border: 1px solid #ddd; padding: 12px; text-align: left;">カテゴリー</th>
            <th style="border: 1px solid #ddd; padding: 12px;">件数</th>
            <th style="border: 1px solid #ddd; padding: 12px;">累計時間</th>
            <th style="border: 1px solid #ddd; padding: 12px;">前週比</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>

      ${chartHtml}

      <h3 style="background-color: #f8f9fa; padding: 10px; border-left: 5px solid #4285F4; margin-top: 40px;">AI Insight (Gemini)</h3>
      <div style="background-color: #f4f6f8; padding: 25px; border-radius: 12px; border-top: 1px solid #e0e0e0; white-space: pre-wrap; margin-top: 15px;">
${aiCommentary}
      </div>

      <footer style="margin-top: 50px; border-top: 1px solid #eee; padding-top: 20px; font-size: 0.85em; color: #888; text-align: center;">
        このメールはシステムより自動送信されています。<br>
        本日も良い一日をお過ごしください。
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
 * ユーティリティ
 */
function getWeeklyDateRange(refDate, offsetWeeks) {
    const date = new Date(refDate.getTime());
    const day = date.getDay();
    const diffToMon = (day === 0 ? -6 : 1 - day);
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate() + diffToMon + (offsetWeeks * 7), 0, 0, 0);
    const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59, 999);
    return { start, end };
}

function formatDate(date) {
    return Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
}

/**
 * トリガー設定
 */
function createWeeklyTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => { if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();
    console.log("トリガーを設定しました。");
}
