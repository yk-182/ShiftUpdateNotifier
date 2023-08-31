/**
 * Apps Scriptのトリガー設定は1分毎にファイルの更新日時をチェック
 */

// 各メンバーの通知設定情報
const membersSettings = {
  tanaka: {
    type: "email", address: "efdfe24jod@example.com", token: ""
  },
  yamada: {
    type: "none", address: "", token: ""
  },
  kobayashi: {
    type: "line", address: "", token: "example1NMKAnci909sgtxrbRXxJktW7GeZxa6cqXbm"
  },
  sato: {
    type: "email", address: "efj9302fod@example.com", token: ""
  },
  ito: {
    type: "none", address: "", token: ""
  },
  kato: {
    type: "line", address: "", token: "exampleRzmrPAh8wa3jiltRO7L3b0EfTsUDI3qnEBt"
  }
};
// 対象のGoogleDriveフォルダID
const shiftFolderId = "1m9cljyfUaMM_x1nxKoexampleG9Tom_K";
// Excelのシフトファイル名
const excelFileName = "shift.xlsm";
// Excelファイルから変換作成されるGSファイル名
const gsFileName = "shift.gs";
// 予定表Excelファイル
let excelFile = "";
// 予定表Excelファイルのリンクアドレス
let excelFileUrl = "";
// シフトPDFファイルのURL
let pdfFileUrl = "";
// 通知に記載する日付
let dateForTitle = "";

/**
 * シフト更新通知処理
 */
function noticeUpdate() {
  // 1分以内に更新があった場合
  if (detectFileUpdate()) {
    // ExcelファイルからGSファイル作成
    excel2Gs();
    // GSファイルからPDFファイル作成しそのリンクアドレスを変数に設定
    gs2Pdf();
    // 対象日の出勤者を取得
    let targetMembersArr = getTargetMembers();
    // 通知送信処理
    sendNotification(targetMembersArr);
  } else {
    // 更新がなかった場合そのまま終了
    return;
  }
}

/**
 * 1分以内にファイルの更新があったかを調査
 */ 
function detectFileUpdate() {
  let isUpdated = false;
  // 判定用に現在日時を取得
  let now = new Date();
  // 予定表Excelファイルを取得
  excelFile = DriveApp.getFilesByName(excelFileName).next();
  // 予定表Excelファイルのリンクアドレス取得
  excelFileUrl = getShortLink(excelFile.getUrl());
  // 予定表Excelファイルの更新日時を取得
  let updatedDate = excelFile.getLastUpdated();
  // 差分計算
  let timeDiff = (now.getTime() - updatedDate.getTime()) / (60 * 1000);
  // ファイルが1分以内に更新された場合
  if (1 >= timeDiff) {
    isUpdated = true;
  }
  return isUpdated;
}

/**
 * ExcelファイルからGSファイルを作成
 */
function excel2Gs() {
  // 前回作成したGSファイルを削除
  if (DriveApp.getFilesByName(gsFileName).hasNext()) {
    let oldGsFile = DriveApp.getFilesByName(gsFileName);
    oldGsFile.next().setTrashed(true);
  } else {
    Logger.log("there is no Spreadsheets to delete...");
  }
  // GSファイル作成
  let option = {
    // Google sheets
    mimeType: MimeType.GOOGLE_SHEETS,
    // 出力先フォルダ
    parents: [{ id: shiftFolderId }],
    // 出力先ファイル名
    title: gsFileName,
  };
  Drive.Files.insert(option, excelFile);
}

/**
 * GSファイルからPDFファイルを作成しそのリンクアドレスを取得
 */
function gs2Pdf() {
  // 作成するPDFファイル名
  let pdfFileName = "shift.pdf";
  // 変換元のGSファイルを取得
  let gsFile = DriveApp.getFilesByName(gsFileName).next();
  
  // PDF用にGSファイルの列幅を調整
  let targetSheetObj = SpreadsheetApp.open(gsFile).getSheetByName("Sheet1");
  targetSheetObj.setColumnWidth(5, 910);
  targetSheetObj.deleteColumn(4);

  // 前回作成したPDFファイルを削除
  if (DriveApp.getFilesByName(pdfFileName).hasNext()) {
    let oldPdfFiles = DriveApp.getFilesByName(pdfFileName);
    oldPdfFiles.next().setTrashed(true);
  } else {
    Logger.log("there is no PDF file to delete...");
  }
  // PDFファイル作成
  let option = {
    mimeType: MimeType.pdf,
    parents: [{ id: shiftFolderId }],
    title: pdfFileName,
  };
  Drive.Files.insert(option, gsFile);
  // PDFファイルのリンクアドレス取得（フォルダは共有設定済み）
  let pdfFile = DriveApp.getFilesByName(pdfFileName).next();
  pdfFileUrl = getShortLink(pdfFile.getUrl());
}

/**
 * スプレッドシートから当日の出勤者のみの情報を取得
 */
function getTargetMembers() {
  // shift.gsからシフト対象日を取得
  let gsFile = DriveApp.getFilesByName(gsFileName).next();
  let targetSheetObj = SpreadsheetApp.open(gsFile).getSheetByName("Sheet1");
  // シフト対象日を取得
  let dateObj = targetSheetObj.getRange("A2").getValue();
  // 日にちを取得
  let targetDay = dateObj.getDate();
  // 通知本文に記載するの日付を取得
  dateForTitle = dateObj.getMonth() + 1 + "月" + targetDay + "日";

  // noticeUpadateのmonthly_scheduleシートから対象日の出勤者を取得
  let workScheduleSheet = SpreadsheetApp.getActive().getActiveSheet();
  // 全メンバーの名前を取得
  let membersList = workScheduleSheet.getRange(6, 1, 12).getValues();
  // 出勤有無が記載された範囲を取得（+2はセルの位置調整のため）
  let attendanceInfo = workScheduleSheet.getRange(6, targetDay + 2, 12).getValues();
  // 戻り値用の出勤者のみの名前を設定する配列
  let targetMembersArr = [];
  let targetMembersMap = new Map();
  // MAP作成（key:名前、value:出勤有無）
  for (let i = 0; i < membersList.length; i++) {
    targetMembersMap.set(membersList[i], attendanceInfo[i]);
  }
  // 非出勤者は削除
  targetMembersMap.forEach(function (value, key) {
    if (value[0] == "◯" || value[0] == "◎") {
      targetMembersMap.delete(key);
    } else {
      targetMembersArr.push(key[0]);
    }
  });
  return targetMembersArr;
}

/**
 * 通知送信処理
 * targetMembersArr 配列 対象日の出勤者一覧
 */
function sendNotification(targetMembersArr) {
  // 通知されるメンバーの確認用メソッド
  // checkMembers(targetMembersArr)
  for (const targetMember of targetMembersArr) {
    // 名前の存在確認
    if (typeof membersSettings[targetMember] !== 'undefined') {
      // 通知方法判定
      if (membersSettings[targetMember]["type"] == "email") {
        // メール送信
        sendMail(membersSettings[targetMember]["address"]);
      } else if (membersSettings[targetMember]["type"] == "line"){
        // LINE送信
        postLine(membersSettings[targetMember]["token"]);
      }
    } else {
      Logger.log("unkonwn member to send...")
    }
  }
}

/**
 * メール送信処理
 */
function sendMail(targetAddress) {
  // メールタイトル
  let titletext = "【" + dateForTitle + "】 " + "シフト更新のお知らせ";
  // メール本文
  let bodyText =
    "新しいシフトが作成されました。\n・PDF\n" + pdfFileUrl + "\n・Excel\n" + excelFileUrl
  // メール送信
  MailApp.sendEmail(targetAddress, titletext, bodyText);
}

// LINE投稿処理
function postLine(token) {
  // tokenの存在チェック
  if (!token) {
    return
  }
  let lineNotifyUrl  = "https://notify-api.line.me/api/notify"
  let messageText = "\n" + dateForTitle + "のシフトが作成されました。\n・PDF\n" + pdfFileUrl + "\n・Excel\n" + excelFileUrl
  let options = {
    "method" : "post",
    "headers" : {
      "Authorization" : "Bearer "+ token
    },
    "payload" : {
      "message" : messageText
    }
  }
  // LINE送信
  UrlFetchApp.fetch(lineNotifyUrl, options)
}

// リンクアドレスの短縮URLを取得
function getShortLink(url) {
  try {
    if (url == undefined) {
      throw 'url is empty or is not a valid url...'
    }
    let content = UrlFetchApp.fetch('https://tinyurl.com/api-create.php?url=' + encodeURI(url));
    if (content.getResponseCode() != 200) {
      return 'An error occured: [ ' + content.getContentText() + ' ]';
    }
    // HTTPレスポンスを文字列に変換して返却
    return content.getContentText();
  } catch (e) {
    return 'An error occured: [ ' + e + ' ]';
  }
}

// 通知するメンバー名をメール送信
function checkMembers(targetMembersArr) {
  targetAddress = "baysik126@gmail.com"
  let titletext = "更新通知メンバーのお知らせ"
  let bodyText = targetMembersArr
  MailApp.sendEmail(targetAddress, titletext, bodyText);
}
