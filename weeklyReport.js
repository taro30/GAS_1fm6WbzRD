/**
 * 週次レポート生成プログラム
 * カレンダーから【カテゴリー】を抽出し、時間集計、グラフ化、AI分析を行って送信します。
 */

/**
 * 週次レポート送信のメイン関数
 */
function sendWeeklyReport() {
    try {
        const now = new Date();
        const thisWeekRange = getWeeklyDateRange(now, 0);
        const lastWeekRange = getWeeklyDateRange(now, -1);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const dbSheet = ss.getSheetByName('DB');
        if (!dbSheet) throw new Error("'DB' シートが見つかりません。");

        // 1. 全データを一度だけ読み込む（効率化）
        const dataMatrix = dbSheet.getDataRange().getValues();
        if (dataMatrix.length <= 1) return;

        // 2. データをメモリ上で処理
        const thisWeekEvents = getFilteredEvents(dataMatrix, thisWeekRange.start, thisWeekRange.end);
        const lastWeekEvents = getFilteredEvents(dataMatrix, lastWeekRange.start, lastWeekRange.end);

        // 3. 統計集計（メモリ上のデータを使用）
        const thisStats = summarizeEvents(thisWeekEvents);
        const lastStats = summarizeEvents(lastWeekEvents);
        const comparison = buildComparison(thisStats, lastStats);

        // 4. グラフの生成（「日ごとの積み上げグラフ」）
        const chartBlob = createStackedDailyChart(thisWeekEvents, thisWeekRange.start);

        // 5. AI寸評の取得
        const aiCommentary = askGeminiForAnalysis(comparison);

        // 6. メールの送信
        const dateStr = formatDate(thisWeekRange.start) + " ～ " + formatDate(thisWeekRange.end);
        const subject = "[週次レポート] カレンダー集計 (" + dateStr + ")";

        sendReportEmail(subject, comparison, aiCommentary, chartBlob, dateStr);

    } catch (e) {
        console.error("実行エラー: " + e.message);
        throw e;
    }
}

/**
 * 期間内のイベントを抽出
 */
function getFilteredEvents(matrix, start, end) {
    const events = [];
    for (let i = 1; i < matrix.length; i++) {
        const row = matrix[i];
        if (row.length < 6) continue;
        const dateRaw = new Date(row[5]); // F列: 日付
        if (isNaN(dateRaw.getTime())) continue;

        if (dateRaw >= start && dateRaw <= end) {
            events.push({
                title: String(row[0]),     // A列: タイトル
                durationRaw: row[3],      // D列: 所要時間
                date: dateRaw
            });
        }
    }
    return events;
}

/**
 * イベント群をカテゴリー別に集計
 */
function summarizeEvents(events) {
    const summary = {};
    events.forEach(ev => {
        const match = ev.title.match(/【(.*?)】/);
        if (match) {
            const cat = match[1];
            if (!summary[cat]) summary[cat] = { count: 0, hours: 0 };

            // 時間の計算 (Date型または数値)
            let h = 0;
            if (ev.durationRaw instanceof Date) {
                h = ev.durationRaw.getHours() + (ev.durationRaw.getMinutes() / 60);
            } else if (typeof ev.durationRaw === 'number') {
                h = ev.durationRaw * 24;
            }

            summary[cat].count += 1;
            summary[cat].hours += h;
        }
    });
    return summary;
}

/**
 * 日ごとの積み上げ棒グラフを作成
 */
function createStackedDailyChart(events, weekStart) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dailyMap = {}; // { MM/dd: { cat: hours } }
    const catSet = new Set();

    events.forEach(ev => {
        const match = ev.title.match(/【(.*?)】/);
        if (match) {
            const cat = match[1];
            const ds = Utilities.formatDate(ev.date, 'JST', 'MM/dd');
            catSet.add(cat);

            let h = 0;
            if (ev.durationRaw instanceof Date) {
                h = ev.durationRaw.getHours() + (ev.durationRaw.getMinutes() / 60);
            } else if (typeof ev.durationRaw === 'number') {
                h = ev.durationRaw * 24;
            }

            if (!dailyMap[ds]) dailyMap[ds] = {};
            dailyMap[ds][cat] = (dailyMap[ds][cat] || 0) + h;
        }
    });

    const categories = Array.from(catSet).sort();
    if (categories.length === 0) return null;

    // 一時シートの作成
    const tmpName = "chart_data_" + new Date().getTime();
    const tmp = ss.insertSheet(tmpName);

    try {
        const header = ["日付", ...categories];
        const dataForSheet = [header];

        for (let i = 0; i < 7; i++) {
            const day = new Date(weekStart.getTime());
            day.setDate(day.getDate() + i);
            const dateStr = Utilities.formatDate(day, 'JST', 'MM/dd');
            const row = [dateStr];
            categories.forEach(c => row.push(dailyMap[dateStr] ? (dailyMap[dateStr][c] || 0) : 0));
            dataForSheet.push(row);
        }

        // 列と行が不足して「Out of bounds」になるのを防ぐ
        const neededRows = dataForSheet.length;
        const neededCols = header.length;
        if (tmp.getMaxRows() < neededRows) tmp.insertRowsAfter(tmp.getMaxRows(), neededRows - tmp.getMaxRows());
        if (tmp.getMaxColumns() < neededCols) tmp.insertColumnsAfter(tmp.getMaxColumns(), neededCols - tmp.getMaxColumns());

        // データの書き込み
        tmp.getRange(1, 1, neededRows, neededCols).setValues(dataForSheet);

        // グラフの作成
        const chart = tmp.newChart()
            .setChartType(Charts.ChartType.COLUMN)
            .addRange(tmp.getRange(1, 1, neededRows, neededCols))
            .setOption('isStacked', true)
            .setOption('title', '日次カテゴリー別 時間配分 (時間)')
            .setOption('hAxis', { title: '日付' })
            .setOption('vAxis', { title: '時間 (h)' })
            .setOption('legend', { position: 'right' })
            .setOption('width', 600)
            .setOption('height', 400)
            .build();

        return chart.getAs('image/png').setName('stacked_chart.png');
    } finally {
        ss.deleteSheet(tmp);
    }
}

/**
 * Geminiによる分析
 */
function askGeminiForAnalysis(comparison) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "※AI寸評はAPIキー未設定のため作成されませんでした。";

    const endpoint = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + apiKey;
    const prompt = `あなたは生活習慣を分析し、より良いライフスタイルを提案する専門AIです。
以下の「今週のカレンダー集計データ（前週比較含む）」を読み解き、ユーザーに向けて【300文字以上400文字以内】で詳しく日本語のアドバイスを作成してください。

分析のポイント：
- どのカテゴリーに最も時間を費やしたか、前週と比べてどう変化したか。
- 仕事、休憩、自己研鑽（インプット）、生活（家事・育児等）のバランスが健全か。
- 具体的な気づき（例：休憩が減りすぎている、仕事効率が上がっている等）と、翌週への提案。
- 丁寧で温かみのある、親身なトーンで記述してください。

集計データ:
${JSON.stringify(comparison)}
`;

    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    const res = UrlFetchApp.fetch(endpoint, options);
    const json = JSON.parse(res.getContentText());

    if (json.candidates && json.candidates[0].content.parts[0].text) {
        return json.candidates[0].content.parts[0].text.trim();
    }
    return "分析レポートの生成中にエラーが発生しました。";
}

/**
 * 統計データの整理
 */
function buildComparison(thisStats, lastStats) {
    const allCats = new Set([...Object.keys(thisStats), ...Object.keys(lastStats)]);
    const res = [];
    allCats.forEach(cat => {
        const t = thisStats[cat] || { count: 0, hours: 0 };
        const l = lastStats[cat] || { count: 0, hours: 0 };
        res.push({
            category: cat,
            currentCount: t.count,
            currentHours: t.hours,
            diffHours: t.hours - l.hours
        });
    });
    return res.sort((a, b) => b.currentHours - a.currentHours);
}

/**
 * HTMLメールの送信
 */
function sendReportEmail(subject, comparison, aiText, chartBlob, dateRange) {
    const email = Session.getActiveUser().getEmail();

    let table = "";
    comparison.forEach(item => {
        const diff = item.diffHours.toFixed(1);
        const color = item.diffHours > 0 ? "#4285F4" : (item.diffHours < 0 ? "#EA4335" : "#333");
        table += `
      <tr>
        <td style="border:1px solid #ddd; padding:10px; font-weight:bold;">【${item.category}】</td>
        <td style="border:1px solid #ddd; padding:10px; text-align:center;">${item.currentCount}回</td>
        <td style="border:1px solid #ddd; padding:10px; text-align:center;">${item.currentHours.toFixed(1)}h</td>
        <td style="border:1px solid #ddd; padding:10px; text-align:center; color:${color};">${item.diffHours > 0 ? '+' : ''}${diff}h</td>
      </tr>`;
    });

    const chartHtml = chartBlob
        ? `<div style="margin:30px 0;text-align:center;"><img src="cid:chartImg" style="width:100%;max-width:600px;border-radius:10px;border:1px solid #eee;"/></div>`
        : "";

    const html = `
    <div style="font-family:sans-serif; max-width:650px; margin:0 auto; padding:20px; color:#333;">
      <h2 style="color:#4285F4; border-bottom:3px solid #4285F4; padding-bottom:5px;">週次ライフログ・レポート</h2>
      <p style="margin:15px 0;">[対象期間] ${dateRange}</p>
      
      <h3 style="background:#f8f9fa; padding:8px; border-left:5px solid #ccc; margin-top:25px;">● カテゴリー別実績サマリー</h3>
      <table style="border-collapse:collapse; width:100%; margin-top:10px;">
        <thead>
          <tr style="background:#eee;">
            <th style="border:1px solid #ddd; padding:10px; text-align:left;">カテゴリー</th>
            <th style="border:1px solid #ddd; padding:10px;">件数</th>
            <th style="border:1px solid #ddd; padding:10px;">累計時間</th>
            <th style="border:1px solid #ddd; padding:10px;">前週比</th>
          </tr>
        </thead>
        <tbody>
          ${table}
        </tbody>
      </table>

      ${chartHtml}

      <h3 style="background:#f8f9fa; padding:8px; border-left:5px solid #4285F4; margin-top:35px;">■ AI Insight (Gemini)</h3>
      <div style="background:#f4f6f8; padding:20px; border-radius:10px; margin-top:10px; line-height:1.7; white-space:pre-wrap;">
${aiText}
      </div>

      <footer style="margin-top:40px; border-top:1px solid #eee; padding-top:15px; font-size:0.85em; color:#888; text-align:center;">
        このレポートはシステムの自動集計により生成されています。<br>
        今日も素晴らしい一日を！
      </footer>
    </div>
  `;

    GmailApp.sendEmail(email, subject, "", {
        htmlBody: html,
        inlineImages: chartBlob ? { chartImg: chartBlob } : {},
        attachments: chartBlob ? [chartBlob] : []
    });
}

/**
 * ユーティリティ
 */
function getWeeklyDateRange(refDate, offsetWeeks) {
    const d = new Date(refDate.getTime());
    const day = d.getDay();
    const diff = (day === 0 ? -6 : 1 - day);
    const start = new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff + (offsetWeeks * 7), 0, 0, 0);
    const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59, 999);
    return { start, end };
}

function formatDate(date) {
    return Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
}

/**
 * トリガー設定用
 */
function createWeeklyTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => { if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();
    console.log("トリガーを設定しました（毎週日曜 9:00）。");
}
