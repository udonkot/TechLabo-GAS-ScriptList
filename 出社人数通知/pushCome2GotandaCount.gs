/**
 * 出社人数通知スプレッドシート操作スクリプト
 * スプレッドシートから値を取得してメッセージを生成、Slackへ通知
 */

// Slackトークン、プロパティから取得
const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN");

// スプレッドシート情報、プロパティから取得
const SHEET_URL = PropertiesService.getScriptProperties().getProperty("SPREADSHEETS_URL");

// SlackAPI
const POST_URL = PropertiesService.getScriptProperties().getProperty("SLACK_APIURL_POSTMESSAGE");

// Slack ChannelID
// bot_test
const CHANNEL_ID = PropertiesService.getScriptProperties().getProperty("SLACK_CHANNELID_BOTTEST");
// cmn-本社7f出退勤連絡
// const CHANNEL_ID = PropertiesService.getScriptProperties().getProperty("SLACK_CHANNELID_COME2GOTANDA");

/**
 * main function
 */
function pushCome2GotandaCount() {

  let sendMessage = '';

  try {

    // Spreadsheetファイルを開く
    var spreadSheet = SpreadsheetApp.openByUrl(SHEET_URL);

    // システム日付取得
    today = new Date();

    // シート名取得
    const sheetName = getSheetName(today);
    // Sheetを開く
    var sheet = spreadSheet.getSheetByName(sheetName);

    // 翌日取得
    const nextDay = getNextWorkDay(today);

    // 翌日の列番号取得
    let targetCol = 0;
    // 1列ごとのに判定、3か月分判定できればよいので100で仮置き。
    for(let i=3; i < 100; i++) {
      // 日付セル取得
      const dateCell = sheet.getRange(15,i);

      // 通知対象の日付か判定
      if(dateCell.getValue().toString() == nextDay.toString()) {
        targetCol = i;
        break;
      }

      // 空のセルがきたら終了
      if(dateCell == undefined || dateCell.getValue() == '') {
        break;
      }
    }

    // ***************
    // 出力メッセージ生成
    // ***************
    const header = '【' + nextDay.getFullYear() + '/' + (nextDay.getMonth()+1) + '/' + nextDay.getDate() + 'の出社人数】';
    const earth = createSummaryMessage(sheet, 4, targetCol);
    const venus = createSummaryMessage(sheet, 5, targetCol);
    const mars = createSummaryMessage(sheet, 6, targetCol);
    const mercury = createSummaryMessage(sheet, 7, targetCol);
    const sixA = createSummaryMessage(sheet, 8, targetCol);
    const sixB = createSummaryMessage(sheet, 9, targetCol);
    const sixZ = createSummaryMessage(sheet, 10, targetCol);

    // 改行コードをつけて整形
    sendMessage = header + '\n' + earth + '\n' + venus + '\n' + mars + '\n' + mercury + '\n' + sixA + '\n' + sixB + '\n' + sixZ

  } catch(e) {
    Logger.log(e);
    throw new Error("SpreadSheet Error:");
  }

  try {
    // slackにpostするデータ生成
    const res = sendMessage;
    const options = getSlackConnectData(CHANNEL_ID, res);
    Logger.log(options);

    // slackに通知
    return UrlFetchApp.fetch(POST_URL, options);
    // return true
  } catch(e) {
    Logger.log(e);
    throw new Error("UrlFetchApp Error:");
  }
}

/**
 * 集計用メッセージ生成
 */
const createSummaryMessage = (sheet, rowNumber, colNumber) => {

  // 会議室名
  const roomName = sheet.getRange(rowNumber, 2).getValue();
  // 人数
  const count = sheet.getRange(rowNumber, colNumber).getValue().toString(); 
  // 
  return roomName + ':' + count + '人'

}

/**
 * Slack送信オブジェクト生成
 */
const getSlackConnectData = (channel, res) => {
  var payload = {
    token   : SLACK_TOKEN,
    channel : channel,
    text    : res,
  }

  var options = {
    'method' : 'post',
    // 'contentType': 'application/json',
    // 'payload': JSON.stringify(payload)
    'payload': payload
  }

  return options;
 
}

/**
 * TODO：シート名取得
 * 参照するシート名を取得する
 * 
 * nowDate:システム日付のオブジェクト
 */
const getSheetName = (nowDate) => {
  // システム日付から年月を取得

  // 取得した月から現在Qを特定して文字列を生成
  // 1Q(10月~12月) → 10-12
  // 2Q(01月~3月) → 01-03
  // 3Q(04月~06月) → 04-06
  // 4Q(07月~09月) → 07-09

  // "システム年_上記文字列のシート名を生成して返却

  // FIXME: 動的にシート名を生成するよう修正が必要
  const sheetName = "2023_07-09"

  return sheetName;


}

/**
 * TODO：土日チェック
 * システム日付が土日が判定する(true：土日)
 * 
 * nowDate:システム日付のオブジェクト
 * 
 */
const isHoliday = (nowDate) => {
  // 

  return false;

}

/**
 * TODO：翌営業日取得
 * システム日付の翌営業日を返却する
 * 
 * nowDate:システム日付のオブジェクト
 * 
 */
const getNextWorkDay = (nowDate) => {

  // FIXME:翌日が土日の場合は翌週月曜の日付を返却するよう修正が必要
  const nextWorkDay = nowDate;
  nextWorkDay.setDate(nextWorkDay.getDate()+1);
  nextWorkDay.setHours(0,0,0)

  return nextWorkDay;


}

