// ============================================================
// Google Apps Script - 出席・ふりかえり管理システム
// ============================================================
//
// 【セットアップ手順】
// 1. Google Spreadsheet を新規作成
// 2. シートを16枚作成し、以下の名前にする：
//    「第1回」「第2回」...「第15回」「出席サマリー」
// 3. 各回シート（第1回〜第15回）の1行目にヘッダーを入力：
//    A1: タイムスタンプ / B1: 学籍番号 / C1: 名前 / D1: メールアドレス / E1: ふりかえり
// 4. 「出席サマリー」シートの1行目にヘッダーを入力：
//    A1: 学籍番号 / B1: 名前 / C1: 第1回 / D1: 第2回 / ... / Q1: 第15回 / R1: 出席回数
// 5. スプレッドシートのIDをコピーし、下記 SPREADSHEET_ID に貼り付ける
// 6. 「拡張機能」→「Apps Script」でこのコードを貼り付け
// 7. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
//    - 実行するユーザー: 自分
//    - アクセスできるユーザー: 全員
// 8. デプロイ後に表示されるURLをコピーし、各セッションページの
//    SCRIPT_URL 変数に設定する
//
// 【注意】
// - メール送信には Gmail の1日あたりの送信制限があります（無料: 100通/日）
// - 初回デプロイ時に Google アカウントの認証が必要です
// ============================================================

// ★ ここにスプレッドシートのIDを入力してください
const SPREADSHEET_ID = 'ここにスプレッドシートIDを貼り付け';

// 授業名（メール件名に使用）
const COURSE_NAME = '人文学とデジタル技術（2026）';

/**
 * POST リクエストを受け取り、出席データを記録する
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sessionNum = data.session;      // 回数（1〜15）
    const studentId = data.studentId;     // 学籍番号
    const name = data.name;               // 名前
    const email = data.email;             // メールアドレス
    const reflection = data.reflection;   // ふりかえり

    // バリデーション
    if (!sessionNum || !studentId || !name || !email || !reflection) {
      return createResponse(false, '全ての項目を入力してください。');
    }

    if (sessionNum < 1 || sessionNum > 15) {
      return createResponse(false, '無効な授業回です。');
    }

    if (reflection.length < 200) {
      return createResponse(false, 'ふりかえりは200文字以上で記入してください（現在 ' + reflection.length + ' 文字）。');
    }

    // メールアドレスの簡易バリデーション
    if (!email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
      return createResponse(false, '有効なメールアドレスを入力してください。');
    }

    // スプレッドシートに書き込み
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = '第' + sessionNum + '回';
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return createResponse(false, 'シート「' + sheetName + '」が見つかりません。');
    }

    // 同じ学籍番号で既に提出済みかチェック
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][1] === studentId) {
        return createResponse(false, 'この学籍番号では既に第' + sessionNum + '回のふりかえりが提出済みです。');
      }
    }

    // データを追記
    const timestamp = new Date();
    sheet.appendRow([timestamp, studentId, name, email, reflection]);

    // サマリーシートを更新
    updateSummary(ss, studentId, name, sessionNum);

    // 確認メールを送信
    sendConfirmationEmail(email, name, sessionNum, studentId, reflection, timestamp);

    return createResponse(true, '第' + sessionNum + '回のふりかえりを提出しました。確認メールを ' + email + ' に送信しました。');

  } catch (error) {
    console.error(error);
    return createResponse(false, 'エラーが発生しました: ' + error.message);
  }
}

/**
 * 出席サマリーシートを更新する
 */
function updateSummary(ss, studentId, name, sessionNum) {
  const summarySheet = ss.getSheetByName('出席サマリー');
  if (!summarySheet) return;

  const data = summarySheet.getDataRange().getValues();
  let rowIndex = -1;

  // 既存の学籍番号を探す
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      rowIndex = i + 1; // シートの行番号（1始まり）
      break;
    }
  }

  // 該当回の列: A=学籍番号, B=名前, C=第1回, D=第2回, ... Q=第15回, R=出席回数
  const sessionCol = sessionNum + 2; // 第1回 = 列C(3) → sessionNum(1) + 2

  if (rowIndex === -1) {
    // 新規学生：行を追加
    const newRow = summarySheet.getLastRow() + 1;
    summarySheet.getRange(newRow, 1).setValue(studentId);
    summarySheet.getRange(newRow, 2).setValue(name);
    summarySheet.getRange(newRow, sessionCol).setValue('○');
    // 出席回数の数式を設定（C列〜Q列の「○」をカウント）
    summarySheet.getRange(newRow, 18).setFormula(
      '=COUNTIF(C' + newRow + ':Q' + newRow + ',"○")'
    );
  } else {
    // 既存学生：該当回を更新
    summarySheet.getRange(rowIndex, sessionCol).setValue('○');
  }
}

/**
 * 確認メールを送信する
 */
function sendConfirmationEmail(email, name, sessionNum, studentId, reflection, timestamp) {
  const subject = '【' + COURSE_NAME + '】第' + sessionNum + '回 ふりかえり提出確認';

  const formattedDate = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

  const body = name + ' さん\n\n'
    + '以下の内容でふりかえりを受け付けました。\n\n'
    + '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n'
    + '授業回: 第' + sessionNum + '回\n'
    + '学籍番号: ' + studentId + '\n'
    + '名前: ' + name + '\n'
    + '提出日時: ' + formattedDate + '\n'
    + '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'
    + '【ふりかえり内容】\n'
    + reflection + '\n\n'
    + '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n'
    + 'このメールは自動送信されています。\n'
    + '内容に誤りがある場合は、担当教員にご連絡ください。\n';

  GmailApp.sendEmail(email, subject, body, {
    name: COURSE_NAME + ' 出席管理システム'
  });
}

/**
 * JSON レスポンスを作成する
 */
function createResponse(success, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: success, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET リクエスト（ブラウザからのアクセス確認用）
 */
function doGet() {
  return ContentService
    .createTextOutput('出席管理システムは稼働中です。')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * 初期セットアップ：スプレッドシートのシートとヘッダーを自動作成する
 * （Apps Script エディタから手動で1回だけ実行してください）
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 各回のシートを作成
  for (let i = 1; i <= 15; i++) {
    const sheetName = '第' + i + '回';
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    // ヘッダー設定
    const headers = ['タイムスタンプ', '学籍番号', '名前', 'メールアドレス', 'ふりかえり'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // 列幅の調整
    sheet.setColumnWidth(1, 160); // タイムスタンプ
    sheet.setColumnWidth(2, 120); // 学籍番号
    sheet.setColumnWidth(3, 120); // 名前
    sheet.setColumnWidth(4, 200); // メールアドレス
    sheet.setColumnWidth(5, 500); // ふりかえり
  }

  // サマリーシートを作成
  let summarySheet = ss.getSheetByName('出席サマリー');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('出席サマリー');
  }
  const summaryHeaders = ['学籍番号', '名前'];
  for (let i = 1; i <= 15; i++) {
    summaryHeaders.push('第' + i + '回');
  }
  summaryHeaders.push('出席回数');
  summarySheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
  summarySheet.getRange(1, 1, 1, summaryHeaders.length).setFontWeight('bold');
  summarySheet.setFrozenRows(1);

  // デフォルトの「シート1」を削除（存在する場合）
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.getUi().alert('セットアップが完了しました！');
}
