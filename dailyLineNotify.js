/**
 * @fileoverview 日次LINE通知プログラム
 * 前日の活動実績（カテゴリー別件数・時間）をDBシートから集計し、
 * 過去1週間の平均と比較した分析結果をLINE Messaging APIを使用して通知します。
 */

/**
 * 【メイン関数】前日のカテゴリー別統計と分析をLINEに通知します。
 * 毎朝 5:00 に前日の実績をブロードキャストすることを想定しています。
 */
function dailyLineNotify() {
    try {
        const now = new Date();
        const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);

        // 比較用の過去7日間（昨日のさらに前の7日間）
        const weekStart = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate() - 7);
        const weekEnd = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate() - 1);

        const sheetName = 'DB';

        // 1. スプレッドシートデータの取得
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const dbSheet = ss.getSheetByName(sheetName);
        if (!dbSheet) throw new Error(`'${sheetName}' シートが見つかりません。`);

        const allRows = dbSheet.getDataRange().getValues();
        if (allRows.length <= 1) return;

        // 2. データの抽出（昨日分と過去1週間分）
        const yesterdayEvents = [];
        const pastWeekEvents = [];

        const rangeStart = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate(), 0, 0, 0);
        const rangeEnd = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate(), 23, 59, 59, 999);

        for (let i = 1; i < allRows.length; i++) {
            const row = allRows[i];
            if (row.length < 6) continue;

            const eventDate = new Date(row[5]);
            if (isNaN(eventDate.getTime())) continue;

            // 昨日分
            if (eventDate.toDateString() === yesterday.toDateString()) {
                yesterdayEvents.push({
                    title: String(row[0]),
                    durationSerial: row[3]
                });
            }
            // 過去1週間分（昨日を含まない直近7日間）
            else if (eventDate >= rangeStart && eventDate <= weekEnd) {
                pastWeekEvents.push({
                    title: String(row[0]),
                    durationSerial: row[3]
                });
            }
        }

        if (yesterdayEvents.length === 0) {
            console.log("前日のデータは見つかりませんでした。");
            return;
        }

        // 3. カテゴリー別に集計
        const statsYesterday = aggregateDailyEvents(yesterdayEvents);
        const statsWeek = aggregateDailyEvents(pastWeekEvents);

        // 4. 分析データの構築（昨日 vs 1週間の1日平均）
        const analysisData = buildDailyAnalysis(statsYesterday, statsWeek);

        // 5. Geminiによる寸評の取得
        const aiInsight = getGeminiDailyInsight(analysisData);

        // 6. LINEメッセージの構築
        const dateStr = Utilities.formatDate(yesterday, 'JST', 'yyyy/MM/dd(E)');
        let message = `【昨日の活動実績】\n📅 ${dateStr}\n\n`;

        analysisData.forEach(item => {
            const diff = item.diff.toFixed(1);
            const mark = item.diff > 0 ? "▲" : (item.diff < 0 ? "▼" : " ");
            message += `■${item.category}\n  ${item.hours.toFixed(1)}h (平均比:${mark}${Math.abs(diff)}h)\n`;
        });

        message += `\n【AIリフレクション】\n${aiInsight}\n\n`;
        message += `今日も素晴らしい一日を！`;

        // 7. LINE送信
        sendLineMessage(message);

    } catch (e) {
        console.error(`日次LINE通知エラー: ${e.message}`);
    }
}

/**
 * イベント群からカテゴリー統計を算出
 */
function aggregateDailyEvents(events) {
    const stats = {};
    events.forEach(ev => {
        const match = ev.title.match(/【(.*?)】/);
        if (match) {
            const cat = match[1];
            if (!stats[cat]) stats[cat] = { hours: 0 };

            let h = 0;
            if (ev.durationSerial instanceof Date) h = ev.durationSerial.getHours() + (ev.durationSerial.getMinutes() / 60);
            else if (typeof ev.durationSerial === 'number') h = ev.durationSerial * 24;

            stats[cat].hours += h;
        }
    });
    return stats;
}

/**
 * 昨日と過去1週間平均の比較データを構築
 */
function buildDailyAnalysis(yesterday, week) {
    const allCats = new Set([...Object.keys(yesterday), ...Object.keys(week)]);
    const res = [];
    allCats.forEach(cat => {
        const yHours = yesterday[cat] ? yesterday[cat].hours : 0;
        const wAvgHours = week[cat] ? week[cat].hours / 7 : 0; // 7日間の平均

        // 昨日活動があった、または平均的に活動があるもののみ
        if (yHours > 0 || wAvgHours > 0.1) {
            res.push({
                category: cat,
                hours: yHours,
                avg: wAvgHours,
                diff: yHours - wAvgHours
            });
        }
    });
    return res.sort((a, b) => b.hours - a.hours);
}

/**
 * Geminiによる短寸評の取得
 */
function getGeminiDailyInsight(data) {
    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return "（AI分析はAPIキー未設定のためスキップします）";

    const endpoint = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=" + key;
    const prompt = `あなたはライフログコーチです。昨日の活動実績と直近1週間の1日平均の比較データを見て、短く鋭い日本語の寸評を【120文字以内】で作成してください。
LINEで読むため、簡潔かつ前向きなアドバイスにしてください。

比較データ(昨日 vs 1日平均):
${JSON.stringify(data)}
`;

    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

    try {
        const res = UrlFetchApp.fetch(endpoint, options);
        const json = JSON.parse(res.getContentText());
        if (json.candidates && json.candidates[0].content.parts[0].text) {
            return json.candidates[0].content.parts[0].text.trim();
        }
    } catch (e) {
        console.error("Gemini Error: " + e.message);
    }
    return "分析中...";
}

/**
 * LINEメッセージ送信用の共通関数
 */
function sendLineMessage(text) {
    const url = 'https://api.line.me/v2/bot/message/broadcast';
    const token = PropertiesService.getScriptProperties().getProperty('line_personal_channel_token');

    if (!token) {
        console.warn("line_personal_channel_token が未設定です。");
        return;
    }

    const payload = { messages: [{ type: 'text', text: text }] };
    const params = {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        UrlFetchApp.fetch(url, params);
    } catch (e) {
        console.error(`LINE通信例外: ${e.message}`);
    }
}

/**
 * 毎日午前5時ごろに実行するトリガーを作成
 */
function createDailyLineTrigger() {
    const handler = 'dailyLineNotify';
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => { if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t); });

    ScriptApp.newTrigger(handler)
        .timeBased()
        .everyDays(1)
        .atHour(5)
        .create();

    console.log("日次分析LINE通知のトリガーを設定しました（毎日 05:00）。");
}
