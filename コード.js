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

function parseYmdDate(value) {
  if (!value) return null;
  const parts = String(value).trim().split(/[\/\-]/);
  if (parts.length !== 3) return null;
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
  const dt = new Date(y, m, d);
  dt.setHours(0, 0, 0, 0);
  return dt;
}

function detectCategory(classInfo) {
  if (String(classInfo).includes("中学")) return "中学生";
  if (String(classInfo).includes("高校")) return "高校生";
  return "職員教職員";
}

function normalizeCategory(raw) {
  const value = String(raw || "").trim();
  if (!value) return "";
  if (value.includes("中")) return "中学生";
  if (value.includes("高")) return "高校生";
  if (value.includes("職")) return "職員教職員";
  return value;
}

function buildRosterContactMap(rosterConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(rosterConfig.sheetName);
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  let map = {};
  let idCol = parseInt(rosterConfig.idCol, 10) - 1;
  let emailCol = parseInt(rosterConfig.emailCol, 10) - 1;
  let categoryCol = rosterConfig.categoryCol ? parseInt(rosterConfig.categoryCol, 10) - 1 : -1;
  let nameCol = rosterConfig.nameCol ? parseInt(rosterConfig.nameCol, 10) - 1 : -1;
  
  for (let i = 1; i < data.length; i++) {
    let id = String(data[i][idCol]).trim();
    let email = String(data[i][emailCol]).trim();
    let category = categoryCol >= 0 ? normalizeCategory(data[i][categoryCol]) : "";
    let studentName = nameCol >= 0 ? String(data[i][nameCol]).trim() : "";
    if (!id || !email) continue;

    map[id] = { email: email, name: studentName, category: category };
    if (category) {
      map[`${category}|${id}`] = { email: email, name: studentName, category: category };
    }
  }
  return map;
}

function getRosterConfigs(settings) {
  const sheetConfigs = loadRosterConfigsFromSheet();
  if (sheetConfigs.length > 0) return sheetConfigs;

  let parsed = [];
  if (settings.rosterConfigs) {
    try {
      parsed = JSON.parse(settings.rosterConfigs);
    } catch (e) {
      throw new Error("名簿設定（rosterConfigs）のJSON形式が不正です。");
    }
  } else if (settings.sheetName && settings.idCol && settings.emailCol) {
    // 旧設定形式の後方互換
    parsed = [{
      name: "既定名簿",
      sheetName: settings.sheetName,
      startDate: "1900/01/01",
      endDate: "2999/12/31",
      idCol: settings.idCol,
      emailCol: settings.emailCol
    }];
  }

  if (!Array.isArray(parsed) || parsed.length === 0) {
    throw new Error("名簿設定がありません。設定画面で名簿設定（JSON）を登録してください。");
  }
  return parsed;
}

function loadRosterConfigsFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("名簿設定");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(v => String(v || "").trim());
  const idx = {
    name: headers.indexOf("name"),
    sheetName: headers.indexOf("sheetName"),
    startDate: headers.indexOf("startDate"),
    endDate: headers.indexOf("endDate"),
    idCol: headers.indexOf("idCol"),
    emailCol: headers.indexOf("emailCol"),
    categoryCol: headers.indexOf("categoryCol"),
    nameCol: headers.indexOf("nameCol")
  };
  if (idx.name < 0 || idx.sheetName < 0 || idx.startDate < 0 || idx.endDate < 0 || idx.idCol < 0 || idx.emailCol < 0) {
    throw new Error("名簿設定シートのヘッダーが不足しています。name,sheetName,startDate,endDate,idCol,emailCol を用意してください。");
  }

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sheetName = String(row[idx.sheetName] || "").trim();
    if (!sheetName) continue;
    result.push({
      name: String(row[idx.name] || "").trim() || sheetName,
      sheetName: sheetName,
      startDate: String(row[idx.startDate] || "").trim(),
      endDate: String(row[idx.endDate] || "").trim(),
      idCol: Number(row[idx.idCol]),
      emailCol: Number(row[idx.emailCol]),
      categoryCol: idx.categoryCol >= 0 && row[idx.categoryCol] !== "" ? Number(row[idx.categoryCol]) : undefined,
      nameCol: idx.nameCol >= 0 && row[idx.nameCol] !== "" ? Number(row[idx.nameCol]) : undefined
    });
  }
  return result;
}

function setupRosterConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("名簿設定");
  if (!sheet) sheet = ss.insertSheet("名簿設定");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["name", "sheetName", "startDate", "endDate", "idCol", "emailCol", "categoryCol", "nameCol"]);
  } else if (sheet.getLastRow() >= 1) {
    const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v || "").trim());
    if (firstRow[0] !== "name") {
      sheet.insertRows(1, 1);
      sheet.getRange(1, 1, 1, 8).setValues([["name", "sheetName", "startDate", "endDate", "idCol", "emailCol", "categoryCol", "nameCol"]]);
    }
  }
  return "名簿設定シートを準備しました。";
}

function selectRosterConfig(settings, recordDate) {
  const rosterConfigs = getRosterConfigs(settings);
  for (let i = 0; i < rosterConfigs.length; i++) {
    const cfg = rosterConfigs[i];
    const start = parseYmdDate(cfg.startDate);
    const end = parseYmdDate(cfg.endDate);
    if (!start || !end) continue;
    if (recordDate >= start && recordDate <= end) return cfg;
  }
  return null;
}

function resolveRecordsWithContacts(records, settings) {
  const dateBase = settings.rosterDateBase === "dueDate" ? "dueDate" : "borrowDate";
  const rosterMaps = {};
  const resolved = [];

  records.forEach(r => {
    let id = normalizeTo4DigitId(r.classInfo, r.name);
    let category = detectCategory(r.classInfo);
    const baseDate = parseYmdDate(dateBase === "dueDate" ? r.dueDate : r.borrowDate);
    const selectedRoster = baseDate ? selectRosterConfig(settings, baseDate) : null;
    const rosterName = selectedRoster ? selectedRoster.name || selectedRoster.sheetName : "";

    let email = "";
    let resolvedName = "";
    if (selectedRoster) {
      const rosterKey = selectedRoster.name || selectedRoster.sheetName;
      if (!rosterMaps[rosterKey]) {
        rosterMaps[rosterKey] = buildRosterContactMap(selectedRoster);
      }
      const rosterMap = rosterMaps[rosterKey];
      const contact = rosterMap[`${category}|${id}`] || rosterMap[id];
      if (contact) {
        email = contact.email || "";
        resolvedName = contact.name || "";
      }
    }

    resolved.push(Object.assign({}, r, {
      id: id,
      category: category,
      rosterName: rosterName,
      resolvedName: resolvedName,
      email: email,
      dateBase: dateBase,
      hasRoster: !!selectedRoster
    }));
  });

  return resolved;
}

function getScheduledPreview(payload) {
  const records = payload.records || [];
  const settings = payload.settings || {};
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const subReminder = settings.subReminder || "【図書室】返却期限間近のお知らせ";
  const tplReminder = settings.tplReminder || "【氏名】 様\n\n図書室です。\n以下の図書の返却期限が【明日】となっております。\n\n【書籍名】\n\n期限内のご返却をお願いいたします。";

  const resolvedRecords = resolveRecordsWithContacts(records, settings);
  const rows = [];

  resolvedRecords.forEach(r => {
    const dueDate = parseYmdDate(r.dueDate);
    const isAllowed = (r.classInfo.includes("中学") && String(settings.sendJ) === "true") ||
      (r.classInfo.includes("高校") && String(settings.sendH) === "true") ||
      ((!r.classInfo.match(/(中学|高校)/)) && String(settings.sendStaff) === "true");

    const status = dueDate && dueDate < today ? "督促（期限超過）" : "予約（前日通知対象）";
    const recipientName = r.resolvedName || r.name;
    rows.push({
      id: r.id,
      category: r.category,
      rosterName: r.rosterName,
      classInfo: r.classInfo,
      borrowerName: r.name,
      recipientName: recipientName,
      email: r.email,
      title: r.title,
      borrowDate: r.borrowDate,
      dueDate: r.dueDate,
      status: status,
      isAllowed: isAllowed,
      warning: !r.hasRoster ? "基準日に一致する名簿がありません" : (!r.email ? "メールアドレス未登録" : "")
    });
  });
  return { rows: rows };
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
  const dateBase = settings.rosterDateBase === "dueDate" ? "dueDate" : "borrowDate";
  getRosterConfigs(settings);
  const resolvedRecords = resolveRecordsWithContacts(records, settings);
  
  let subOverdue = settings.subOverdue || "【図書室】未返却図書のお知らせ";
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let grouped = { overdue: {}, scheduled: [] };
  let missingEmails = new Set();

  resolvedRecords.forEach(r => {

    let isAllowed = (r.classInfo.includes("中学") && String(settings.sendJ) === "true") || 
                    (r.classInfo.includes("高校") && String(settings.sendH) === "true") || 
                    ((!r.classInfo.match(/(中学|高校)/)) && String(settings.sendStaff) === "true");
    if (!isAllowed) return;

    if (!email) {
      if (!r.hasRoster) {
        missingEmails.add(`${r.classInfo || '教職員等'} ${r.name} (ID: ${r.id}) / 基準日 ${dateBase === "dueDate" ? r.dueDate : r.borrowDate} に一致する名簿なし`);
      } else {
        missingEmails.add(`${r.classInfo || '教職員等'} ${r.name} (ID: ${r.id}, 区分: ${r.category}, 名簿: ${r.rosterName})`);
      }
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
      const displayName = data.user.resolvedName || data.user.name;
      let compiledSubject = subOverdue.replace(/【氏名】/g, displayName)
                                      .replace(/【クラス】/g, data.user.classInfo || "")
                                      .replace(/【書籍名】/g, data.books.join(' '));
      
      // ★本文の変数置換（書籍名は縦並びになるよう改行区切り）
      let compiledBody = settings.tplOverdue.replace(/【氏名】/g, displayName)
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
  sheet.appendRow(['検索用ID', '区分', '氏名', '名簿名', 'クラス', '書籍名', '貸出日', '返却期限', 'メールアドレス']); 
  
  grouped.scheduled.forEach(r => {
    sheet.appendRow([r.id, r.category, r.resolvedName || r.name, r.rosterName, r.classInfo, r.title, r.borrowDate, r.dueDate, r.email]);
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
    let d = String(row[7]).split('/');
    let dueDate = new Date(d[0], d[1] - 1, d[2]);
    dueDate.setHours(0,0,0,0);

    if (dueDate.getTime() === tomorrow.getTime()) {
      let id = row[0];
      if (!groupedReminder[id]) {
        groupedReminder[id] = { name: row[2], classInfo: row[4], email: row[8], books: [] };
      }
      groupedReminder[id].books.push(`・『${row[5]}』（期限: ${row[7]}）`);
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