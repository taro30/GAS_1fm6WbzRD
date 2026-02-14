/**
 * @fileoverview 週次ライフログ・レポート生成プログラム
 * Googleカレンダーからエクスポートされたスプレッドシートのデータ（DBシート）を元に、
 * 1週間の活動をカテゴリー別に集計し、グラフ化、AIによる分析を行ってメール送信します。
 * 
 * 主な機能:
 * 1. 【カテゴリー】形式のタイトルから活動を自動分類
 * 2. カテゴリー別の累積時間・構成比・前週比の算出
 * 3. インプット・アウトプット比率の計算
 * 4. Google Chartsを用いた日次積み上げ棒グラフの生成
 * 5. Gemini API (AI) による高度な活動分析とアドバイス
 */

// --- 設定定数 ---
const CONFIG = {
    SHEET_NAME: 'DB',
    TIME_ZONE: 'JST',
    DATE_FORMAT: 'yyyy/MM/dd',
    CHART_WIDTH: 600,
    CHART_HEIGHT: 400,
    IO_GOAL_RATIO: '3:7', // 目指すべきインプット:アウトプット比
    CHART_COLORS: ["#4285F4", "#34A853", "#FBBC05", "#EA4335", "#673AB7", "#00ACC1", "#FF6D00", "#4E342E"],
    GEMINI_MODEL: 'gemini-2.5-flash', // 使用するAIモデル
    MAX_ANALYSIS_CHARS: 450 // AI分析の最大文字数目安
};

/**
 * 【メイン関数】週次レポートの生成・送信プロセスを実行します。
 * トリガーによって毎週日曜日の朝などに実行されることを想定しています。
 */
function sendWeeklyReport() {
    try {
        const now = new Date();

        // 1. 集計期間の計算 (offset 0: 今週(直近の日～土), offset -1: 前週)
        const thisWeek = getWeeklyDateRange(now, 0);
        const lastWeek = getWeeklyDateRange(now, -1);

        // 2. スプレッドシートデータの取得
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const dbSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
        if (!dbSheet) throw new Error(`'${CONFIG.SHEET_NAME}' シートが見つかりません。`);

        // 効率化のため、全データをメモリ上に一度だけ読み込みます
        const allRows = dbSheet.getDataRange().getValues();
        if (allRows.length <= 1) {
            console.warn("DBシートにヘッダー以外のデータが存在しません。");
            return;
        }

        // 3. メモリ上でのイベント抽出（フィルタリング）
        const thisEvents = filterEventsByRange(allRows, thisWeek.start, thisWeek.end);
        const lastEvents = filterEventsByRange(allRows, lastWeek.start, lastWeek.end);

        // 4. カテゴリー別の統計集計
        const thisStats = aggregateStats(thisEvents);
        const lastStats = aggregateStats(lastEvents);

        // 5. 分析用指標の算出 (構成比、前週比、インアウト比)
        const comparison = buildComparison(thisStats, lastStats);
        const ioMetrics = calculateIOMetrics(thisStats);

        // 6. ビジュアル（グラフ）の生成 (DataTable方式: 一時シート不要で安定)
        const chartBlob = createChartImage(thisEvents, thisWeek.start);

        // 7. AIによる分析レポートの取得
        const aiInsight = getGeminiAnalysis(comparison, ioMetrics);

        // 8. メールの構築と送信
        const dateRangeStr = `${formatDate(thisWeek.start)} ～ ${formatDate(thisWeek.end)}`;
        const subject = `[週次レポート] カレンダー実績集計 (${dateRangeStr})`;

        sendHtmlEmail(subject, comparison, ioMetrics, aiInsight, chartBlob, dateRangeStr);

        console.log(`週次レポート送信成功: ${dateRangeStr}`);

    } catch (e) {
        console.error(`週次レポート生成中にエラーが発生しました: ${e.message}`);
        // 致命的なエラーはスタックトレースと共に再スローし、GASの実行ログに残します
        throw e;
    }
}

/**
 * 渡されたシートデータから期間内のイベントのみを抽出します。
 * @param {Array<Array>} allRows スプレッドシートの全行データ
 * @param {Date} start 始業日時
 * @param {Date} end 終業日時
 * @return {Array<Object>} フィルタリングされたイベントオブジェクトの配列
 */
function filterEventsByRange(allRows, start, end) {
    const events = [];
    // インデックス1から（ヘッダーを飛ばして）ループ
    for (let i = 1; i < allRows.length; i++) {
        const row = allRows[i];
        if (row.length < 6) continue; // 不完全な行はスキップ

        // スプレッドシートの列定義に合わせてインデックスを指定
        // A列(0):タイトル, D列(3):所要時間(シリアル値), F列(5):日付
        const eventDate = new Date(row[5]);
        if (!isNaN(eventDate.getTime()) && eventDate >= start && eventDate <= end) {
            events.push({
                title: String(row[0]),
                durationSerial: row[3],
                date: eventDate
            });
        }
    }
    return events;
}

/**
 * イベントのリストをカテゴリー【 】ごとに集計し、時間と回数を算出します。
 * @param {Array<Object>} events 対象イベントのリスト
 * @return {Object} カテゴリー名をキーとした統計オブジェクト
 */
function aggregateStats(events) {
    const summary = {};
    events.forEach(ev => {
        // タイトルから【 】で囲まれた文字列をカテゴリーとして抽出
        const match = ev.title.match(/【(.*?)】/);
        if (match) {
            const category = match[1];
            if (!summary[category]) summary[category] = { count: 0, hours: 0 };

            // 所要時間の変換 (Google Sheetsのシリアル値またはDateオブジェクトを時間に変換)
            let hours = 0;
            if (ev.durationSerial instanceof Date) {
                // 時刻形式(例 1:30:00) の場合
                hours = ev.durationSerial.getHours() + (ev.durationSerial.getMinutes() / 60);
            } else if (typeof ev.durationSerial === 'number') {
                // 数値形式(例 0.0416... = 1時間) の場合
                hours = ev.durationSerial * 24;
            }

            summary[category].count++;
            summary[category].hours += hours;
        }
    });
    return summary;
}

/**
 * 「インプット」と「アウトプット」に関連する時間の比率を計算します。
 * @param {Object} stats 集計済み統計データ
 * @return {Object} インアウト比率と各種時間のオブジェクト
 */
function calculateIOMetrics(stats) {
    let inputHours = 0;
    let outputHours = 0;

    Object.keys(stats).forEach(cat => {
        // 特定のキーワードを含むカテゴリーを振り分け
        if (cat.includes("インプット")) {
            inputHours += stats[cat].hours;
        } else if (cat.includes("アウトプット") || cat === "中小") {
            outputHours += stats[cat].hours;
        }
    });

    const totalIO = inputHours + outputHours;
    return {
        inputHours: inputHours,
        outputHours: outputHours,
        inputRatio: totalIO > 0 ? (inputHours / totalIO * 100).toFixed(1) : 0,
        outputRatio: totalIO > 0 ? (outputHours / totalIO * 100).toFixed(1) : 0
    };
}

/**
 * 日ごとのカテゴリー別時間を集計し、積み上げ棒グラフの画像を生成します。
 * DataTableを使用することで、一時シートを作成せずにメモリ上で完結させます。
 * @param {Array<Object>} events 1週間のイベント
 * @param {Date} weekStart 週の開始日（日曜日）
 * @return {Blob} グラフ画像のBlobデータ
 */
function createChartImage(events, weekStart) {
    const dailyAllocation = {}; // { MM/dd: { category: hours } }
    const uniqueCategories = new Set();

    // 日次・カテゴリー別のマトリックスを作成
    events.forEach(ev => {
        const match = ev.title.match(/【(.*?)】/);
        if (!match) return;

        const cat = match[1];
        const dateLabel = Utilities.formatDate(ev.date, CONFIG.TIME_ZONE, 'MM/dd');
        uniqueCategories.add(cat);

        let h = (ev.durationSerial instanceof Date)
            ? (ev.durationSerial.getHours() + ev.durationSerial.getMinutes() / 60)
            : (typeof ev.durationSerial === 'number' ? ev.durationSerial * 24 : 0);

        if (!dailyAllocation[dateLabel]) dailyAllocation[dateLabel] = {};
        dailyAllocation[dateLabel][cat] = (dailyAllocation[dateLabel][cat] || 0) + h;
    });

    const sortedCats = Array.from(uniqueCategories).sort();
    if (sortedCats.length === 0) return null;

    // 1. Google Chartsのデータテーブルを定義
    const dataTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "日付");
    sortedCats.forEach(c => dataTable.addColumn(Charts.ColumnType.NUMBER, c));

    // 2. 7日間分（日〜土）の行データを追加
    for (let i = 0; i < 7; i++) {
        const cursor = new Date(weekStart.getTime());
        cursor.setDate(cursor.getDate() + i);
        const label = Utilities.formatDate(cursor, CONFIG.TIME_ZONE, 'MM/dd');

        const row = [label];
        sortedCats.forEach(c => {
            row.push(dailyAllocation[label] ? (dailyAllocation[label][c] || 0) : 0);
        });
        dataTable.addRow(row);
    }

    // 3. グラフの構築
    const chart = Charts.newColumnChart()
        .setDataTable(dataTable.build())
        .setStacked()
        .setTitle('日次カテゴリー別 時間配分 (時間)')
        .setXAxisTitle('日付')
        .setYAxisTitle('時間(h)')
        .setDimensions(CONFIG.CHART_WIDTH, CONFIG.CHART_HEIGHT)
        .setLegendPosition(Charts.Position.RIGHT)
        .setColors(CONFIG.CHART_COLORS)
        .build();

    return chart.getAs('image/png').setName('activity_chart.png');
}

/**
 * 集計データを元に、Gemini APIを使用して分析レポートを生成します。
 * @param {Array<Object>} comparison 各カテゴリーの比較データ
 * @param {Object} ioMetrics インプット・アウトプット比率
 * @return {string} AIによる日本語寸評
 */
function getGeminiAnalysis(comparison, ioMetrics) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "※AI寸評は、スクリプトプロパティに「GEMINI_API_KEY」が設定されていないためスキップされました。";

    const apiUrl = `https://generativelanguage.googleapis.com/v1/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

    // AIへの指示（プロンプト）の構築
    const prompt = `あなたはライフログ分析のプロフェッショナルAIです。
以下の集計データ（活動構成比、インプット/アウトプット比率、前週比較）を読み解き、ユーザーの生活リズムと質の変化について【350文字～450文字程度】で日本語のアドバイスを記述してください。

着眼点：
- カテゴリー別の「時間(h)」の変化（何が増え、何が減り、それが生活にどう影響しているか）。
- プロフェッショナルな視点でのワークライフバランスや自己研鑽の評価。
- インプット(${ioMetrics.inputRatio}%) 対 アウトプット(${ioMetrics.outputRatio}%)の比率（理想の黄金比は3:7とされています）から見た示唆。
- 温かみがあり、かつ気づきを与えるトーンで記述してください。

データ: ${JSON.stringify({ comparison, ioMetrics })}`;

    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const params = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(apiUrl, params);
        const json = JSON.parse(response.getContentText());

        if (json.candidates && json.candidates[0].content.parts[0].text) {
            return json.candidates[0].content.parts[0].text.trim();
        }
    } catch (err) {
        console.error(`Gemini API呼び出しエラー: ${err.message}`);
    }
    return "AI分析レポートの生成中に不具合が発生しました。集計数値のみご確認ください。";
}

/**
 * 今週と前週のデータを比較し、構成比と増減を計算した配列を生成します。
 */
function buildComparison(thisStats, lastStats) {
    const allCategories = new Set([...Object.keys(thisStats), ...Object.keys(lastStats)]);

    // 今週の総記録時間を算出 (構成比の計算用)
    let totalHours = 0;
    Object.keys(thisStats).forEach(k => totalHours += thisStats[k].hours);

    const result = [];
    allCategories.forEach(cat => {
        const t = thisStats[cat] || { count: 0, hours: 0 };
        const l = lastStats[cat] || { count: 0, hours: 0 };

        // 構成比と差分の計算
        const ratio = totalHours > 0 ? (t.hours / totalHours * 100).toFixed(1) : 0;
        const diff = t.hours - l.hours;

        result.push({
            category: cat,
            currentCount: t.count,
            currentHours: t.hours,
            ratio: ratio,
            diffHours: diff
        });
    });

    // 時間が長い順にソートします
    return result.sort((a, b) => b.currentHours - a.currentHours);
}

/**
 * リッチなHTMLメールを構築して送信します。
 */
function sendHtmlEmail(subject, comparison, ioMetrics, aiText, chartBlob, dateRange) {
    const me = Session.getActiveUser().getEmail();

    // 1. 各カテゴリーの統計表（行）を作成
    let tableRows = "";
    comparison.forEach(item => {
        const diff = item.diffHours.toFixed(1);
        const diffColor = item.diffHours > 0 ? "#4285F4" : (item.diffHours < 0 ? "#EA4335" : "#333");

        tableRows += `
      <tr>
        <td style="border: 1px solid #ddd; padding: 12px; font-weight: bold; background-color: #fafafa;">【${item.category}】</td>
        <td style="border: 1px solid #ddd; padding: 12px; text-align: center;">${item.currentCount}回</td>
        <td style="border: 1px solid #ddd; padding: 12px; text-align: center; font-weight: 500;">${item.currentHours.toFixed(1)}h (${item.ratio}%)</td>
        <td style="border: 1px solid #ddd; padding: 12px; text-align: center; color: ${diffColor}; font-weight: bold;">
          ${item.diffHours > 0 ? '+' : ''}${diff}h
        </td>
      </tr>`;
    });

    // 2. グラフHTML（画像が生成できた場合のみ）
    const chartHtml = chartBlob
        ? `<div style="margin: 35px 0; text-align: center;">
         <img src="cid:chartImg" style="width: 100%; max-width: ${CONFIG.CHART_WIDTH}px; border-radius: 12px; border: 1px solid #eee; box-shadow: 0 4px 8px rgba(0,0,0,0.05);" />
       </div>`
        : "";

    // 3. 全体のHTMLメールテンプレート
    const htmlBody = `
    <div style="font-family: 'Helvetica Neue', Arial, sans-serif; max-width: 650px; margin: 0 auto; color: #333; line-height: 1.8; background-color: #fff; border: 1px solid #eee; padding: 30px; border-radius: 12px;">
      <h2 style="color: #4285F4; border-bottom: 3px solid #4285F4; padding-bottom: 12px; margin-top: 0;">週次ライフログ・レポート</h2>
      <p style="font-size: 1.1em; color: #666; font-weight: bold; margin-bottom: 30px;">[対象期間] ${dateRange}</p>
      
      <!-- インアウト比率セクション -->
      <div style="background: #FFF8E1; padding: 20px; border-radius: 10px; border-left: 6px solid #FFC107; margin-bottom: 35px;">
        <h4 style="margin: 0 0 10px 0; color: #FF8F00;">● インプット・アウトプット比率</h4>
        理想バランス ${CONFIG.IO_GOAL_RATIO} にむけて<br>
        <span style="font-size: 1.3em; font-weight: bold; color: #1B5E20;">
          IN ${ioMetrics.inputRatio}% (${ioMetrics.inputHours.toFixed(1)}h) : OUT ${ioMetrics.outputRatio}% (${ioMetrics.outputHours.toFixed(1)}h)
        </span>
      </div>

      <!-- メイン集計テーブル -->
      <h3 style="background: #f8f9fa; padding: 10px; border-left: 5px solid #ccc; font-size: 1em;">● カテゴリー別実績</h3>
      <table style="border-collapse: collapse; width: 100%; margin-top: 15px; border-radius: 8px; overflow: hidden;">
        <thead>
          <tr style="background-color: #eee; color: #444; font-size: 0.9em;">
            <th style="border: 1px solid #ddd; padding: 12px; text-align: left;">カテゴリー</th>
            <th style="border: 1px solid #ddd; padding: 12px;">件数</th>
            <th style="border: 1px solid #ddd; padding: 12px;">時間(構成比)</th>
            <th style="border: 1px solid #ddd; padding: 12px;">前週比</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>

      ${chartHtml}

      <!-- AI分析セクション -->
      <h3 style="background: #f8f9fa; padding: 10px; border-left: 5px solid #4285F4; margin-top: 45px; font-size: 1em;">■ AI Insight (Gemini)</h3>
      <div style="background-color: #F1F3F4; padding: 25px; border-radius: 12px; line-height: 1.9; white-space: pre-wrap; font-size: 1em; color: #202124;">
${aiText}
      </div>

      <footer style="margin-top: 50px; font-size: 0.85em; color: #999; text-align: center; border-top: 1px solid #f0f0f0; padding-top: 25px;">
        このレポートはシステムの自動集計により構成されています。<br>
        毎日を丁寧に振り返る、あなたのパートナーでありたい。
      </footer>
    </div>`;

    // 送信の設定
    const options = {
        htmlBody: htmlBody,
        inlineImages: chartBlob ? { chartImg: chartBlob } : {},
        attachments: chartBlob ? [chartBlob] : []
    };

    GmailApp.sendEmail(me, subject, "", options);
}

/**
 * 集計期間の計算ロジック
 * 日〜土の範囲を計算し、日曜日の朝に実行された場合に前週分（直近の土曜まで）を集計します。
 * @param {Date} refDate 基準日時（通常は今日）
 * @param {number} offsetWeeks 週のオフセット (0:今週, -1:前週)
 */
function getWeeklyDateRange(refDate, offsetWeeks) {
    const d = new Date(refDate.getTime());
    const day = d.getDay(); // 0:日, 6:土

    // 本日が日曜(0)なら7日前を、それ以外ならその週の日曜を特定します
    const diffToSunday = (day === 0) ? 7 : day;

    // 基準となる週の日曜日
    const start = new Date(d.getFullYear(), d.getMonth(), d.getDate() - diffToSunday + (offsetWeeks * 7), 0, 0, 0);
    // そこから6日後の土曜日
    const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59, 999);

    return { start, end };
}

/**
 * 日付を指定のフォーマットで文字列化します。
 */
function formatDate(date) {
    return Utilities.formatDate(date, CONFIG.TIME_ZONE, CONFIG.DATE_FORMAT);
}

/**
 * 定期実行用トリガーの設定
 * 毎週日曜日の午前9時にsendWeeklyReportを実行するようにセットします。
 */
function createWeeklyTrigger() {
    const handler = 'sendWeeklyReport';
    const existingTriggers = ScriptApp.getProjectTriggers();

    // 既存の重複トリガーを削除
    existingTriggers.forEach(t => {
        if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
    });

    // 日曜日、午前9時に実行
    ScriptApp.newTrigger(handler)
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.SUNDAY)
        .atHour(9)
        .create();

    console.log("定期レポートのトリガーを設定しました（毎週日曜 9:00）。");
}
