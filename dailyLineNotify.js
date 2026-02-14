/**
 * @fileoverview 日次LINE通知プログラム
 * その日の活動実績（カテゴリー別件数・時間）をDBシートから集計し、
 * LINE Messaging APIを使用して通知します。
 */

/**
 * 【メイン関数】その日のカテゴリー別統計をLINEに通知します。
 * 毎日 23:00〜など、一日の終わりに実行するトリガー設定を想定しています。
 */
function dailyLineNotify() {
    try {
        const today = new Date();
        const sheetName = 'DB';

        // 1. スプレッドシートデータの取得
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const dbSheet = ss.getSheetByName(sheetName);
        if (!dbSheet) throw new Error(`'${sheetName}' シートが見つかりません。`);

        const allRows = dbSheet.getDataRange().getValues();
        if (allRows.length <= 1) return;

        // 2. 本日のイベントのみを抽出
        // 時刻を00:00:00にリセットした比較用の日付を作成
        const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
        const endOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59, 999);

        const todayEvents = [];
        for (let i = 1; i < allRows.length; i++) {
            const row = allRows[i];
            if (row.length < 6) continue;

            const eventDate = new Date(row[5]); // F列: 日付
            if (!isNaN(eventDate.getTime()) && eventDate >= startOfToday && eventDate <= endOfToday) {
                todayEvents.push({
                    title: String(row[0]),      // A列: タイトル
                    durationSerial: row[3]      // D列: 所要時間
                });
            }
        }

        if (todayEvents.length === 0) {
            console.log("本日のデータはまだ記録されていません。");
            return;
        }

        // 3. カテゴリー別に集計
        const stats = {};
        let totalDayHours = 0;

        todayEvents.forEach(ev => {
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
        const dateStr = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd(E)');
        let message = `【本日の活動実績】\n📅 ${dateStr}\n\n`;

        // 時間の長い順に並び替え
        const sortedCats = Object.keys(stats).sort((a, b) => stats[b].hours - stats[a].hours);

        sortedCats.forEach(cat => {
            const s = stats[cat];
            message += `■${cat}\n  ${s.count}回 / ${s.hours.toFixed(1)}h\n`;
        });

        message += `\n合計記録時間: ${totalDayHours.toFixed(1)}h\n`;
        message += `今日もお疲れ様でした！`;

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
 * 毎日23時ごろに実行するトリガーを作成
 */
function createDailyLineTrigger() {
    const handler = 'dailyLineNotify';
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => { if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t); });

    ScriptApp.newTrigger(handler)
        .timeBased()
        .everyDays(1)
        .atHour(23)
        .create();

    console.log("日次LINE通知のトリガーを設定しました（毎日 23:00）。");
}
