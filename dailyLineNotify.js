/**
 * @fileoverview 日次LINE通知プログラム
 * 前日の活動実績（カテゴリー別件数・時間）をDBシートから集計し、
 * LINE Messaging APIを使用して通知します。
 */

/**
 * 【メイン関数】前日のカテゴリー別統計をLINEに通知します。
 * 毎朝 5:00 に前日の実績をブロードキャストすることを想定しています。
 */
function dailyLineNotify() {
    try {
        const now = new Date();
        // 前日の日付を取得
        const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);

        const sheetName = 'DB';

        // 1. スプレッドシートデータの取得
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const dbSheet = ss.getSheetByName(sheetName);
        if (!dbSheet) throw new Error(`'${sheetName}' シートが見つかりません。`);

        const allRows = dbSheet.getDataRange().getValues();
        if (allRows.length <= 1) return;

        // 2. 前日のイベントのみを抽出
        // 時刻を00:00:00と23:59:59に設定した範囲を作成
        const startOfDate = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate(), 0, 0, 0);
        const endOfDate = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate(), 23, 59, 59, 999);

        const yesterdayEvents = [];
        for (let i = 1; i < allRows.length; i++) {
            const row = allRows[i];
            if (row.length < 6) continue;

            const eventDate = new Date(row[5]); // F列: 日付
            if (!isNaN(eventDate.getTime()) && eventDate >= startOfDate && eventDate <= endOfDate) {
                yesterdayEvents.push({
                    title: String(row[0]),      // A列: タイトル
                    durationSerial: row[3]      // D列: 所要時間
                });
            }
        }

        if (yesterdayEvents.length === 0) {
            console.log("前日のデータは見つかりませんでした。");
            return;
        }

        // 3. カテゴリー別に集計
        const stats = {};
        let totalDayHours = 0;

        yesterdayEvents.forEach(ev => {
            const match = ev.title.match(/【(.*?)】/);
            if (match) {
                const cat = match[1];
                if (!stats[cat]) stats[cat] = { count: 0, hours: 0 };

                let hours = 0;
                if (ev.durationSerial instanceof Date) {
                    hours = ev.durationSerial.getHours() + (ev.durationSerial.getMinutes() / 60);
                } else if (typeof ev.durationSerial === 'number') {
                    hours = ev.durationSerial * 24;
                }

                stats[cat].count++;
                stats[cat].hours += hours;
                totalDayHours += hours;
            }
        });

        // 4. LINEメッセージの構築
        const dateStr = Utilities.formatDate(yesterday, 'JST', 'yyyy/MM/dd(E)');
        let message = `【昨日の活動実績】\n📅 ${dateStr}\n\n`;

        // 時間の長い順に並び替え
        const sortedCats = Object.keys(stats).sort((a, b) => stats[b].hours - stats[a].hours);

        sortedCats.forEach(cat => {
            const s = stats[cat];
            message += `■${cat}\n  ${s.count}回 / ${s.hours.toFixed(1)}h\n`;
        });

        message += `\n合計記録時間: ${totalDayHours.toFixed(1)}h\n`;
        message += `今日も一日、充実した日になりますように！`;

        // 5. LINE送信
        sendLineMessage(message);

    } catch (e) {
        console.error(`日次LINE通知エラー: ${e.message}`);
    }
}

/**
 * LINEメッセージ送信用の共通関数
 * 宛先はMessaging APIのブロードキャスト機能を使用します。
 * @param {string} text 送信するテキスト内容
 */
function sendLineMessage(text) {
    const url = 'https://api.line.me/v2/bot/message/broadcast';
    const token = PropertiesService.getScriptProperties().getProperty('line_personal_channel_token');

    if (!token) {
        console.warn("line_personal_channel_token が未設定です。");
        return;
    }

    const payload = {
        messages: [
            { type: 'text', text: text }
        ]
    };

    const params = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            Authorization: 'Bearer ' + token
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(url, params);
        const code = response.getResponseCode();
        if (code === 200) {
            console.log("LINE通知の送信に成功しました。");
        } else {
            console.error(`LINE APIエラー (Status:${code}): ${response.getContentText()}`);
        }
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

    console.log("日次LINE通知のトリガーを設定しました（毎日 05:00）。");
}
