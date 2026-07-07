function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('図書室 督促システム V7')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 設定の保存と読み込み
function saveSettings(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('システム設定');
  if (!sheet) sheet = ss.insertSheet('システム設定');
  
  sheet.clearContents();
  sheet.appendRow(['設定項目（変更不可）', '設定値']);
  for (let key in settings) {
    sheet.appendRow([key, settings[key]]);
  }
}

function loadSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('システム設定');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for(let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function getContactsMap(sheetName, idColIdx, emailColIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  let map = {};
  let idCol = parseInt(idColIdx) - 1;
  let emailCol = parseInt(emailColIdx) - 1;
  
  for (let i = 1; i < data.length; i++) {
    let id = String(data[i][idCol]).trim();
    let email = String(data[i][emailCol]).trim();
    if (id && email) map[id] = email;
  }
  return map;
}

// IDの正規化（中高・教職員）
function normalizeTo4DigitId(classInfo, name) {
  let match = classInfo.match(/(中学|高校).*?(\d+)年([A-Za-z0-9]+)組(\d+)番/);
  if (!match) return "STAFF_" + name.replace(/[\s ]+/g, ""); 
  return `${match[2]}${match[3]}${String(match[4]).padStart(2, '0')}`;
}

// 送信と予約のメイン処理
function processMails(payload) {
  const records = payload.records;
  saveSettings(payload.settings);
  
  const settings = payload.settings;
  const contacts = getContactsMap(settings.sheetName, settings.idCol, settings.emailCol);
  
  let subOverdue = settings.subOverdue || "【図書室】未返却図書のお知らせ";
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let grouped = { overdue: {}, scheduled: [] };
  let missingEmails = new Set();

  records.forEach(r => {
    let id = normalizeTo4DigitId(r.classInfo, r.name);
    let email = contacts[id] || "";
    r.id = id;
    r.email = email;

    let isAllowed = (r.classInfo.includes("中学") && String(settings.sendJ) === "true") || 
                    (r.classInfo.includes("高校") && String(settings.sendH) === "true") || 
                    ((!r.classInfo.match(/(中学|高校)/)) && String(settings.sendStaff) === "true");
    if (!isAllowed) return;

    if (!email) {
      missingEmails.add(`${r.classInfo || '教職員等'} ${r.name} (ID: ${id})`);
    }

    let d = r.dueDate.split('/');
    let dueDate = new Date(d[0], d[1] - 1, d[2]);
    dueDate.setHours(0, 0, 0, 0);

    if (dueDate < today) {
      if (!grouped.overdue[id]) grouped.overdue[id] = { user: r, books: [] };
      grouped.overdue[id].books.push(`・『${r.title}』（期限: ${r.dueDate}）`);
    } else {
      grouped.scheduled.push(r);
    }
  });

  let sentCount = 0;
  for (let id in grouped.overdue) {
    let data = grouped.overdue[id];
    if (data.user.email) {
      // ★件名の変数置換（書籍名は横並びになるようスペース区切り）
      let compiledSubject = subOverdue.replace(/【氏名】/g, data.user.name)
                                      .replace(/【クラス】/g, data.user.classInfo || "")
                                      .replace(/【書籍名】/g, data.books.join(' '));
      
      // ★本文の変数置換（書籍名は縦並びになるよう改行区切り）
      let compiledBody = settings.tplOverdue.replace(/【氏名】/g, data.user.name)
                                            .replace(/【クラス】/g, data.user.classInfo || "")
                                            .replace(/【書籍名】/g, data.books.join('\n'));
                                            
      GmailApp.sendEmail(data.user.email, compiledSubject, compiledBody);
      sentCount++;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('送信予約キュー');
  if (!sheet) sheet = ss.insertSheet('送信予約キュー');
  
  sheet.clearContents(); 
  sheet.appendRow(['検索用ID', '氏名', 'クラス', '書籍名', '貸出日', '返却期限', 'メールアドレス']); 
  
  grouped.scheduled.forEach(r => {
    sheet.appendRow([r.id, r.name, r.classInfo, r.title, r.borrowDate, r.dueDate, r.email]);
  });

  let msg = `即時督促メール送信: ${sentCount}件\n最新の予約リスト更新: ${grouped.scheduled.length}件`;
  if (missingEmails.size > 0) {
    msg += `\n\n⚠️ 【注意】以下の ${missingEmails.size} 名は名簿からメールアドレスが見つかりませんでした。\n` 
         + Array.from(missingEmails).join('\n');
  }
  return msg;
}

// 毎朝の自動実行用
function processDailyScheduledEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('送信予約キュー');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const today = new Date(); today.setHours(0,0,0,0);
  const tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);

  let groupedReminder = {};
  let newData = [data[0]]; 

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let d = String(row[5]).split('/');
    let dueDate = new Date(d[0], d[1] - 1, d[2]);
    dueDate.setHours(0,0,0,0);

    if (dueDate.getTime() === tomorrow.getTime()) {
      let id = row[0];
      if (!groupedReminder[id]) {
        // ★予約データにもクラス情報を保持
        groupedReminder[id] = { name: row[1], classInfo: row[2], email: row[6], books: [] };
      }
      groupedReminder[id].books.push(`・『${row[3]}』（期限: ${row[5]}）`);
    } else if (dueDate > tomorrow) {
      newData.push(row);
    }
  }

  // スプレッドシートから設定をロード（自動実行時にも最新設定を反映）
  let settings = loadSettings() || {};
  let subReminder = settings.subReminder || "【図書室】返却期限間近のお知らせ";
  let tplReminder = settings.tplReminder || "【氏名】 様\n\n図書室です。\n以下の図書の返却期限が【明日】となっております。\n\n【書籍名】\n\n期限内のご返却をお願いいたします。";

  for (let id in groupedReminder) {
    let user = groupedReminder[id];
    if (user.email) {
      // ★件名の変数置換
      let compiledSubject = subReminder.replace(/【氏名】/g, user.name)
                                       .replace(/【クラス】/g, user.classInfo || "")
                                       .replace(/【書籍名】/g, user.books.join(' '));
      
      // ★本文の変数置換
      let compiledBody = tplReminder.replace(/【氏名】/g, user.name)
                                    .replace(/【クラス】/g, user.classInfo || "")
                                    .replace(/【書籍名】/g, user.books.join('\n'));
                                    
      GmailApp.sendEmail(user.email, compiledSubject, compiledBody);
    }
  }
  
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}