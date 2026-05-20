// =================================================================
// === 【設定が必要な箇所】 ========================================
// =================================================================

// 1. 処理対象のシート名（順番にロックしていく順に並べてください）
const TARGET_SHEET_NAMES = ['一週目', '二週目', '三週目', '四週目']; 

// 2. 保護・クリア（空白にする）範囲
const TARGET_RANGE_A1 = 'B3:H50'; 

// 3. 基準日のセル（ここを14日進めます）
const DATE_CELL_A1 = 'B1'; 

// =================================================================


/**
 * 【金曜日用】次に処理すべきシートを1つだけロックする関数
 */
function protectNextSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();

  // 次にロックすべきシートのインデックスを取得（0〜3）
  let currentIndex = parseInt(scriptProperties.getProperty('CURRENT_INDEX') || 0);
  const sheetName = TARGET_SHEET_NAMES[currentIndex];
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`エラー: シート「${sheetName}」が見つかりません。`);
    return;
  }

  const range = sheet.getRange(TARGET_RANGE_A1);
  const targetA1 = `${sheetName}!${TARGET_RANGE_A1}`;

  // 1. 既存の保護があれば削除（二重防止）
  const allProtections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  allProtections.forEach(p => {
    if (p.getRange().getA1Notation() === targetA1) p.remove();
  });


// 2. 保護を設定し、自分以外編集不可にする
  const protection = range.protect().setDescription(`LOCK_${sheetName}`);
  
  // 一旦、現在の編集者をすべて削除（これであなただけが編集できる状態になります）
  protection.removeEditors(protection.getEditors());

  // ★ここに「編集を許可する人のメールアドレス」を追加します
  protection.addEditors(['xxxxxxxxxxx']); 

  if (protection.canDomainEdit()) protection.setDomainEdit(false);
  // 3. 【重要】このシートのロック開始日を個別に記録
  scriptProperties.setProperty(`LOCK_DATE_${sheetName}`, new Date().toISOString());

  // 4. 次の週のためにインデックスを更新
  scriptProperties.setProperty('CURRENT_INDEX', (currentIndex + 1) % TARGET_SHEET_NAMES.length);

  spreadsheet.toast(`${sheetName} をロックしました。`);
  Logger.log(`${sheetName} をロックし、日付を記録しました。`);
}


/**
 * 【毎日実行用】ロックから2週間経ったシートを探して、解除・空白化・日付更新する関数
 */
/**
 * 【毎日実行用】ロックから2週間経ったシートを探して、解除・空白化・日付更新する関数
 */
function checkAndUnprotectSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date();
  
  const TWO_WEEKS_MS = 14 * 24 * 60 * 60 * 1000; // 判定用の14日間
  const FOUR_WEEKS_MS = 28 * 24 * 60 * 60 * 1000; // ★更新用の28日間（追加）

  TARGET_SHEET_NAMES.forEach(sheetName => {
    // 各シートのロック開始日を取得
    const lockDateStr = scriptProperties.getProperty(`LOCK_DATE_${sheetName}`);
    if (!lockDateStr) return;

    const lockDate = new Date(lockDateStr);
    
    // 今日がロックから14日以上経過しているか判定
    if (today.getTime() - lockDate.getTime() >= TWO_WEEKS_MS) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) return;

      const targetA1 = `${sheetName}!${TARGET_RANGE_A1}`;

      // --- 手順1: ロック解除（より確実な方法に変更） ---
      // シートにある「範囲の保護」をすべて取得して削除する
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      protections.forEach(p => {
        p.remove();
        Logger.log(`${sheetName} の保護を一つ解除しました。`);
      });

      // --- 手順2: セルを消して空白にする ---
      sheet.getRange(TARGET_RANGE_A1).clearContent();

      // --- 手順3: 日付を28日更新する（ここを修正しました） ---
      const dateRange = sheet.getRange(DATE_CELL_A1);
      const currentDate = dateRange.getValue();
      if (currentDate instanceof Date) {
        // 14日ではなく28日分(FOUR_WEEKS_MS)を加算します
        const newDate = new Date(currentDate.getTime() + FOUR_WEEKS_MS);
        dateRange.setValue(newDate);
        Logger.log(`${sheetName} の日付を28日進めて ${newDate.toLocaleDateString()} にしました。`);
      }

      // --- 手順4: このシートのロック記録を削除してリセット完了 ---
      scriptProperties.deleteProperty(`LOCK_DATE_${sheetName}`);
      
      spreadsheet.toast(`${sheetName} の2週間経過。空白化と28日後の日付更新を完了しました。`);
    }
  });
}
function test_prepareLockDate() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date();
  
  // 15日前の日付を作成
  const fifteenDaysAgo = new Date(today.getTime() - (15 * 24 * 60 * 60 * 1000));
  
  // プロパティに確実に存在する「二週目」でテストします
  scriptProperties.setProperty('LOCK_DATE_二週目', fifteenDaysAgo.toISOString());
  
  console.log('テスト準備完了：二週目を15日前のロックに設定しました。');
}




/**
 * シート上の編集を検知して自動で動く関数
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // 1. 「管理」シート以外での編集は無視する
  if (sheet.getName() !== "管理") return;
  
  // 2. B列（2列目）かつ 2行目以降の編集のみ対象にする
  if (range.getColumn() === 2 && range.getRow() >= 2) {
    const row = range.getRow();
    const sheetName = sheet.getRange(row, 1).getValue(); // A列のシート名を取得
    const shouldLock = range.getValue(); // チェックボックスの真偽値（true/false）
    
    // ターゲットとなるシートを操作する
    processSingleSheetLock(sheetName, shouldLock);
  }
}

/**
 * 特定の1シートに対してロックまたは解除を行う
 */
function processSingleSheetLock(sheetName, shouldLock) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(sheetName);
  const me = Session.getEffectiveUser();
  
  if (!targetSheet) {
    ss.toast(`エラー：シート「${sheetName}」が見つかりません。`);
    return;
  }

  // 以前のコードで指定していた範囲（B3:H50）
  const targetRangeA1 = 'B3:H50'; 
  const range = targetSheet.getRange(targetRangeA1);
  const protections = targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  if (shouldLock === true) {
    // 【ロックをかける】
    // 既存の保護がないか確認
    let alreadyProtected = protections.some(p => p.getRange().getA1Notation() === targetRangeA1);
    
    if (!alreadyProtected) {
      const p = range.protect().setDescription(`LOCK_${sheetName}`);
      p.removeEditors(p.getEditors());
      if (p.canEdit()) p.addEditor(me);
      ss.toast(`${sheetName} をロックしました。`);
    }
  } else if (shouldLock === false) {
    // 【ロックを外す】
    protections.forEach(p => {
      // 範囲が一致する保護をすべて解除
      if (p.getRange().getA1Notation() === targetRangeA1) {
        p.remove();
      }
    });
    ss.toast(`${sheetName} のロックを解除しました。`);
  }
}


const LINE_API_TOKEN = "YOUR_LINE_CHANNEL_ACCESS_TOKEN";
const TARGET_GROUP_ID = "YOUR_TARGET_GROUP_ID";

function sendShiftReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();

  // 次にロックされる予定のシート名を自動判定
  let currentIndex = parseInt(scriptProperties.getProperty('CURRENT_INDEX') || 0);
  const targetSheetName = TARGET_SHEET_NAMES[currentIndex];
  
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    ss.toast(`エラー：シート「${targetSheetName}」が見つかりません。`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  let unpaidMembers = [];

  // 判定（3行目開始、A列[0]が名前、J列[9]が「未完了」）
  for (let i = 2; i < data.length; i++) { 
    const name = data[i][0];   
    const status = data[i][8]; 
    
    if (name !== "" && status === "未完了") {
      unpaidMembers.push(name);
    }
  }

  // LINE送信（URLなし）
  if (unpaidMembers.length > 0) {
    const message = 
      "\n【" + targetSheetName + "のシフト提出のお願い】\n" + 
      "以下のメンバーの入力が完了しておりません：\n" + 
      unpaidMembers.join("さん、") + "さん\n" +
      "速やかに入力してください";
                    
    const url = 'https://api.line.me/v2/bot/message/push';
    UrlFetchApp.fetch(url, {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + LINE_API_TOKEN
      },
      'payload': JSON.stringify({
        'to': TARGET_GROUP_ID,
        'messages': [{'type': 'text', 'text': message}]
      })
    });
    ss.toast(`${targetSheetName}のリマインドを送信しました。`);
  } else {
    ss.toast(`${targetSheetName}に未記入者はいませんでした。`);
  }
}
