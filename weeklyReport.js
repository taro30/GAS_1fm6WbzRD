/**
 * 週次ライフログ・レポート生成プログラム
 * カレンダーから【カテゴリー】を抽出し、時間集計、グラフ化、AI分析を行って送信します。
 */

/**
 * メイン実行関数
 */
function sendWeeklyReport() {
    try {
        const now = new Date();
        const thisWeek = getWeeklyDateRange(now, 0);
        const lastWeek = getWeeklyDateRange(now, -1);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const dbSheet = ss.getSheetByName('DB');
        if (!dbSheet) throw new Error("'DB' シートが見つかりません。");

        // 1. データの読み込み（1回のみ）
        const dataValues = dbSheet.getDataRange().getValues();
        if (dataValues.length <= 1) return;

        // 2. メモリ上でのフィルタリング
        const thisEvents = filterEventsByRange(dataValues, thisWeek.start, thisWeek.end);
        const lastEvents = filterEventsByRange(dataValues, lastWeek.start, lastWeek.end);

        // 3. カテゴリー集計
        const thisStats = aggregateStats(thisEvents);
        const lastStats = aggregateStats(lastEvents);
        const comparison = buildComparison(thisStats, lastStats);

        // 4. グラフ生成 (シートを使わないDataTable方式)
        const chartBlob = createChartImage(thisEvents, thisWeek.start);

        // 5. AI分析の取得
        const aiInsight = getGeminiAnalysis(comparison);

        // 6. メール送信
        const dateRangeStr = formatDate(thisWeek.start) + " ～ " + formatDate(thisWeek.end);
        const subject = "[週次レポート] カレンダー実績集計 (" + dateRangeStr + ")";

        sendHtmlEmail(subject, comparison, aiInsight, chartBlob, dateRangeStr);

    } catch (e) {
        console.error("実行エラー: " + e.message);
        throw e;
    }
}

/**
 * 期間内のデータを抽出
 */
function filterEventsByRange(values, start, end) {
    const list = [];
    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (row.length < 6) continue;
        const d = new Date(row[5]);
        if (!isNaN(d.getTime()) && d >= start && d <= end) {
            list.push({
                title: String(row[0]),
                duration: row[3], // シリアル値
                date: d
            });
        }
    }
    return list;
}

/**
 * カテゴリー統計
 */
function aggregateStats(events) {
    const res = {};
    events.forEach(ev => {
        const match = ev.title.match(/【(.*?)】/);
        if (match) {
            const cat = match[1];
            if (!res[cat]) res[cat] = { count: 0, hours: 0 };

            let h = 0;
            if (ev.duration instanceof Date) {
                h = ev.duration.getHours() + (ev.duration.getMinutes() / 60);
            } else if (typeof ev.duration === 'number') {
                h = ev.duration * 24;
            }
            res[cat].count++;
            res[cat].hours += h;
        }
    });
    return res;
}

/**
 * グラフ用の画像生成 (シートを使わない方式)
 */
function createChartImage(events, weekStart) {
    const dailyData = {}; // { MM/dd: { cat: hours } }
    const catSet = new Set();

    events.forEach(ev => {
        const match = ev.title.match(/【(.*?)】/);
        if (match) {
            const cat = match[1];
            const ds = Utilities.formatDate(ev.date, 'JST', 'MM/dd');
            catSet.add(cat);

            let h = 0;
            if (ev.duration instanceof Date) h = ev.duration.getHours() + (ev.duration.getMinutes() / 60);
            else if (typeof ev.duration === 'number') h = ev.duration * 24;

            if (!dailyData[ds]) dailyData[ds] = {};
            dailyData[ds][cat] = (dailyMap = dailyData[ds][cat] || 0) + h;
        }
    });

    const categories = Array.from(catSet).sort();
    if (categories.length === 0) return null;

    // DataTableの構築
    const dataTable = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "日付");
    categories.forEach(c => dataTable.addColumn(Charts.ColumnType.NUMBER, c));

    for (let i = 0; i < 7; i++) {
        const day = new Date(weekStart.getTime());
        day.setDate(day.getDate() + i);
        const ds = Utilities.formatDate(day, 'JST', 'MM/dd');
        const row = [ds];
        categories.forEach(c => row.push(dailyData[ds] ? (dailyData[ds][c] || 0) : 0));
        dataTable.addRow(row);
    }

    // グラフ構築 (シート不要)
    const chart = Charts.newColumnChart()
        .setDataTable(dataTable.build())
        .setStacked()
        .setTitle('日次活動配分 (時間)')
        .setXAxisTitle('日付')
        .setYAxisTitle('時間(h)')
        .setDimensions(600, 400)
        .setLegendPosition(Charts.Position.RIGHT)
        .setColors(["#4285F4", "#34A853", "#FBBC05", "#EA4335", "#673AB7", "#00ACC1", "#FF6D00", "#4E342E"])
        .build();

    return chart.getAs('image/png').setName('activity_chart.png');
}

/**
 * Gemini分析
 */
function getGeminiAnalysis(comparison) {
    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return "※AI寸評はAPIキー未設定のため表示されません。";

    const endpoint = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + key;
    const prompt = `あなたはライフログ分析のプロフェッショナルAIです。
以下の集計データを分析し、ユーザーの今週の生活リズムと質の変化について【350文字〜400文字程度】で詳しく日本語のアドバイスを記述してください。

着眼点：
- カテゴリー別の「時間(h)」の変化（何が増え、何が減ったか）。
- ワークライフバランスの観点からの気づき（仕事、学習、休息の偏り）。
- データから読み取れるユーザーの頑張りや、来週に向けた具体的な提案。
- 温かみがあり、かつ客観的な洞察に満ちたトーンで答えてください。

データ: ${JSON.stringify(comparison)}`;

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
 * データ比較
 */
function buildComparison(thisStats, lastStats) {
    const keys = new Set([...Object.keys(thisStats), ...Object.keys(lastStats)]);
    const res = [];
    keys.forEach(k => {
        const t = thisStats[k] || { count: 0, hours: 0 };
        const l = lastStats[k] || { count: 0, hours: 0 };
        res.push({
            category: k,
            curCount: t.count,
            curHours: t.hours,
            diffHours: t.hours - l.hours
        });
    });
    return res.sort((a, b) => b.curHours - a.curHours);
}

/**
 * メール送信
 */
function sendHtmlEmail(subject, comparison, aiText, chartBlob, dateRange) {
    const email = Session.getActiveUser().getEmail();

    let table = "";
    comparison.forEach(item => {
        const d = item.diffHours.toFixed(1);
        const c = item.diffHours > 0 ? "#4285F4" : (item.diffHours < 0 ? "#EA4335" : "#333");
        table += `<tr><td style="border:1px solid #ddd;padding:10px;font-weight:bold;">【${item.category}】</td>` +
            `<td style="border:1px solid #ddd;padding:10px;text-align:center;">${item.curCount}回</td>` +
            `<td style="border:1px solid #ddd;padding:10px;text-align:center;">${item.curHours.toFixed(1)}h</td>` +
            `<td style="border:1px solid #ddd;padding:10px;text-align:center;color:${c};">${item.diffHours > 0 ? '+' : ''}${d}h</td></tr>`;
    });

    const chartImg = chartBlob
        ? `<div style="margin:30px 0;text-align:center;"><img src="cid:chartImg" style="width:100%;max-width:600px;border-radius:12px;border:1px solid #eee;"/></div>`
        : "";

    const html = `
    <div style="font-family:sans-serif;max-width:650px;margin:0 auto;padding:25px;color:#333;background-color:#fff;line-height:1.7;">
      <h2 style="color:#4285F4;border-bottom:3px solid #4285F4;padding-bottom:10px;">週次ライフログ・レポート</h2>
      <p style="font-size:1.1em;">[対象期間] ${dateRange}</p>
      
      <h3 style="background:#f8f9fa;padding:10px;border-left:5px solid #ccc;margin-top:25px;">● カテゴリー別実績</h3>
      <table style="border-collapse:collapse;width:100%;margin-top:10px;">
        <tr style="background:#eee;"><th style="border:1px solid #ddd;padding:12px;text-align:left;">カテゴリー</th><th>回数</th><th>時間(h)</th><th>前週比</th></tr>
        ${table}
      </table>
      ${chartImg}
      <h3 style="background:#f8f9fa;padding:10px;border-left:5px solid #4285F4;margin-top:40px;">■ AI Insight (Gemini)</h3>
      <div style="background:#f4f6f8;padding:20px;border-radius:10px;margin-top:15px;white-space:pre-wrap;">${aiText}</div>
      <footer style="margin-top:50px;font-size:0.85em;color:#888;text-align:center;border-top:1px solid #eee;padding-top:20px;">
         自動生成メールです。本日も充実した一日をお過ごしください。
      </footer>
    </div>`;

    GmailApp.sendEmail(email, subject, "", {
        htmlBody: html,
        inlineImages: chartBlob ? { chartImg: chartBlob } : {},
        attachments: chartBlob ? [chartBlob] : []
    });
}

function getWeeklyDateRange(ref, offset) {
    const d = new Date(ref.getTime());
    const day = d.getDay();
    const diff = (day === 0 ? -6 : 1 - day);
    const start = new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff + (offset * 7), 0, 0, 0);
    const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59, 999);
    return { start, end };
}

function formatDate(date) {
    return Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
}

function createWeeklyTrigger() {
    const ts = ScriptApp.getProjectTriggers();
    ts.forEach(t => { if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();
}
