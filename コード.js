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

function normalizeRoster4DigitId(raw) {
  return String(raw || "").trim();
}

function build4DigitIdFromParts(grade, classCode, numberValue) {
  const g = String(grade || "").trim();
  const c = String(classCode || "").trim();
  const n = String(numberValue || "").trim();
  if (!g || !c || !n) return "";
  const n2 = String(parseInt(n, 10));
  if (n2 === "NaN") return "";
  return `${g}${c}${n2.padStart(2, "0")}`;
}

function getHeaderIndexMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || "").trim();
    if (key) map[key] = i;
  });
  return map;
}

function parseStaffEmailOverrides(raw) {
  const map = {};
  const text = String(raw || "").trim();
  if (!text) return map;
  text.split(/\r?\n/).forEach(line => {
    const row = String(line || "").trim();
    if (!row) return;
    const parts = row.split(",");
    if (parts.length < 2) return;
    const name = String(parts[0] || "").trim();
    const email = String(parts[1] || "").trim();
    if (name && email) map[name] = email;
  });
  return map;
}

function buildRosterContactMap(rosterConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(rosterConfig.sheetName);
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const headerRow = Math.max(1, parseInt(rosterConfig.headerRow || 1, 10));
  if (data.length < headerRow) return {};
  const headers = data[headerRow - 1];
  const idxMap = getHeaderIndexMap(headers);

  let map = {};
  const mode = rosterConfig.idMode === "parts" ? "parts" : "direct";
  const idHeader = String(rosterConfig.idHeader || "").trim();
  const gradeHeader = String(rosterConfig.gradeHeader || "").trim();
  const classHeader = String(rosterConfig.classHeader || "").trim();
  const numberHeader = String(rosterConfig.numberHeader || "").trim();
  const nameHeader = String(rosterConfig.nameHeader || "").trim();
  const emailHeader = String(rosterConfig.emailHeader || "").trim();
  const categoryHeader = String(rosterConfig.categoryHeader || "").trim();
  const emailCol = idxMap[emailHeader];
  const nameCol = nameHeader ? idxMap[nameHeader] : undefined;
  const categoryCol = categoryHeader ? idxMap[categoryHeader] : undefined;
  const idCol = idHeader ? idxMap[idHeader] : undefined;
  const gradeCol = gradeHeader ? idxMap[gradeHeader] : undefined;
  const classCol = classHeader ? idxMap[classHeader] : undefined;
  const numberCol = numberHeader ? idxMap[numberHeader] : undefined;
  
  if (emailCol === undefined) return {};

  for (let i = headerRow; i < data.length; i++) {
    let row = data[i];
    let id = "";
    if (mode === "parts") {
      id = build4DigitIdFromParts(
        gradeCol === undefined ? "" : row[gradeCol],
        classCol === undefined ? "" : row[classCol],
        numberCol === undefined ? "" : row[numberCol]
      );
    } else {
      id = normalizeRoster4DigitId(idCol === undefined ? "" : row[idCol]);
    }
    let email = String(row[emailCol] || "").trim();
    let category = categoryCol === undefined ? "" : normalizeCategory(row[categoryCol]);
    let studentName = nameCol === undefined ? "" : String(row[nameCol] || "").trim();
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
    startYear: headers.indexOf("startYear"),
    headerRow: headers.indexOf("headerRow"),
    idMode: headers.indexOf("idMode"),
    idHeader: headers.indexOf("idHeader"),
    gradeHeader: headers.indexOf("gradeHeader"),
    classHeader: headers.indexOf("classHeader"),
    numberHeader: headers.indexOf("numberHeader"),
    nameHeader: headers.indexOf("nameHeader"),
    emailHeader: headers.indexOf("emailHeader"),
    categoryHeader: headers.indexOf("categoryHeader")
  };
  if (idx.name < 0 || idx.sheetName < 0 || idx.startDate < 0 || idx.endDate < 0 || idx.idMode < 0 || idx.emailHeader < 0) {
    throw new Error("名簿設定シートのヘッダーが不足しています。最新版のUIで名簿設定を保存し直してください。");
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
      startYear: idx.startYear >= 0 ? Number(row[idx.startYear]) : undefined,
      headerRow: idx.headerRow >= 0 && row[idx.headerRow] !== "" ? Number(row[idx.headerRow]) : 1,
      idMode: String(row[idx.idMode] || "direct").trim(),
      idHeader: idx.idHeader >= 0 ? String(row[idx.idHeader] || "").trim() : "",
      gradeHeader: idx.gradeHeader >= 0 ? String(row[idx.gradeHeader] || "").trim() : "",
      classHeader: idx.classHeader >= 0 ? String(row[idx.classHeader] || "").trim() : "",
      numberHeader: idx.numberHeader >= 0 ? String(row[idx.numberHeader] || "").trim() : "",
      nameHeader: idx.nameHeader >= 0 ? String(row[idx.nameHeader] || "").trim() : "",
      emailHeader: idx.emailHeader >= 0 ? String(row[idx.emailHeader] || "").trim() : "",
      categoryHeader: idx.categoryHeader >= 0 ? String(row[idx.categoryHeader] || "").trim() : ""
    });
  }
  return result;
}

function setupRosterConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("名簿設定");
  if (!sheet) sheet = ss.insertSheet("名簿設定");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["name", "sheetName", "startYear", "startDate", "endDate", "headerRow", "idMode", "idHeader", "gradeHeader", "classHeader", "numberHeader", "nameHeader", "emailHeader", "categoryHeader"]);
  } else if (sheet.getLastRow() >= 1) {
    const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v || "").trim());
    if (firstRow[0] !== "name") {
      sheet.insertRows(1, 1);
      sheet.getRange(1, 1, 1, 14).setValues([["name", "sheetName", "startYear", "startDate", "endDate", "headerRow", "idMode", "idHeader", "gradeHeader", "classHeader", "numberHeader", "nameHeader", "emailHeader", "categoryHeader"]]);
    }
  }
  return "名簿設定シートを準備しました。";
}

function getRosterConfigsForUi() {
  setupRosterConfigSheet();
  const configs = loadRosterConfigsFromSheet();
  return configs;
}

function saveRosterConfigsFromUi(configs) {
  if (!Array.isArray(configs) || configs.length === 0) {
    throw new Error("名簿設定が空です。1件以上登録してください。");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("名簿設定") || ss.insertSheet("名簿設定");
  const headers = ["name", "sheetName", "startYear", "startDate", "endDate", "headerRow", "idMode", "idHeader", "gradeHeader", "classHeader", "numberHeader", "nameHeader", "emailHeader", "categoryHeader"];
  const values = [headers];

  configs.forEach(c => {
    if (!c.sheetName || !c.startDate || !c.endDate || !c.emailHeader) return;
    const idMode = String(c.idMode || "direct");
    const canUseDirect = idMode === "direct" && String(c.idHeader || "").trim();
    const canUseParts = idMode === "parts" && String(c.gradeHeader || "").trim() && String(c.classHeader || "").trim() && String(c.numberHeader || "").trim();
    if (!(canUseDirect || canUseParts)) return;
    values.push([
      String(c.name || c.sheetName).trim(),
      String(c.sheetName).trim(),
      Number(c.startYear),
      String(c.startDate).trim(),
      String(c.endDate).trim(),
      Number(c.headerRow || 1),
      idMode,
      String(c.idHeader || "").trim(),
      String(c.gradeHeader || "").trim(),
      String(c.classHeader || "").trim(),
      String(c.numberHeader || "").trim(),
      String(c.nameHeader || "").trim(),
      String(c.emailHeader || "").trim(),
      String(c.categoryHeader || "").trim()
    ]);
  });

  if (values.length <= 1) {
    throw new Error("有効な名簿設定がありません。必須項目を入力してください。");
  }

  sheet.clearContents();
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  return `名簿設定を保存しました（${values.length - 1}件）。`;
}

function getRosterUiMeta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const excluded = { "システム設定": true, "送信予約キュー": true, "名簿設定": true };
  const sheetNames = ss.getSheets()
    .map(s => s.getName())
    .filter(name => !excluded[name]);
  return {
    sheetNames: sheetNames,
    currentYear: new Date().getFullYear()
  };
}

function getSheetHeadersForUi(sheetName, headerRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const row = Math.max(1, parseInt(headerRow || 1, 10));
  const maxCol = sheet.getLastColumn();
  if (maxCol <= 0) return [];
  const values = sheet.getRange(row, 1, 1, maxCol).getValues()[0];
  return values.map(v => String(v || "").trim()).filter(v => v);
}

function selectRosterConfig(rosterConfigs, recordDate) {
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
  const rosterConfigs = getRosterConfigs(settings);
  const rosterMaps = {};
  const staffOverrides = parseStaffEmailOverrides(settings.staffOverrides);
  const resolved = [];

  records.forEach(r => {
    let id = normalizeTo4DigitId(r.classInfo, r.name);
    let category = detectCategory(r.classInfo);
    const baseDate = parseYmdDate(dateBase === "dueDate" ? r.dueDate : r.borrowDate);
    const selectedRoster = baseDate ? selectRosterConfig(rosterConfigs, baseDate) : null;
    const rosterName = selectedRoster ? selectedRoster.name || selectedRoster.sheetName : "";

    let email = "";
    let resolvedName = "";
    let contactSource = "none";
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
        contactSource = "roster";
      }
    }

    // 名簿未解決の教職員は個別登録を適用
    if (!email && category === "職員教職員") {
      const overrideEmail = staffOverrides[r.name] || "";
      if (overrideEmail) {
        email = overrideEmail;
        resolvedName = r.name;
        contactSource = "staff_override";
      }
    }

    resolved.push(Object.assign({}, r, {
      id: id,
      category: category,
      rosterName: rosterName,
      resolvedName: resolvedName,
      email: email,
      contactSource: contactSource,
      dateBase: dateBase,
      hasRoster: !!selectedRoster
    }));
  });

  return resolved;
}

function getScheduledPreview(payload) {
  try {
    const records = payload.records || [];
    const settings = payload.settings || {};
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const resolvedRecords = resolveRecordsWithContacts(records, settings);
    const rows = [];

    resolvedRecords.forEach(r => {
      const dueDate = parseYmdDate(r.dueDate);
      const isAllowed = (r.classInfo.includes("中学") && String(settings.sendJ) === "true") ||
        (r.classInfo.includes("高校") && String(settings.sendH) === "true") ||
        ((!r.classInfo.match(/(中学|高校)/)) && String(settings.sendStaff) === "true");

      const status = dueDate && dueDate < today ? "督促（期限超過）" : "予約（前日通知対象）";
      const recipientName = r.resolvedName || r.name;
      let deliveryStatus = "";
      let warning = "";
      if (!isAllowed) {
        deliveryStatus = "送信対象外（区分オフ）";
      } else if (!r.hasRoster && r.contactSource !== "staff_override") {
        deliveryStatus = "名簿未登録で送信不可";
        warning = "基準日に対応する名簿なし";
      } else if (!r.email) {
        if (r.category === "職員教職員") {
          deliveryStatus = "メールアドレス未登録";
          warning = "氏名をクリックしてメール登録";
        } else {
          deliveryStatus = "名簿に記載なしで送信不可";
          warning = "名簿に該当者/メールなし";
        }
      } else if (r.contactSource === "staff_override") {
        deliveryStatus = "送信可（教職員個別登録）";
      } else {
        deliveryStatus = "送信可";
      }

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
        deliveryStatus: deliveryStatus,
        isAllowed: isAllowed,
        warning: warning
      });
    });
    return { rows: rows };
  } catch (e) {
    return { rows: [], error: e.message || String(e) };
  }
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

    if (!r.email) {
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
      if (!grouped.overdue[r.id]) grouped.overdue[r.id] = { user: r, books: [] };
      grouped.overdue[r.id].books.push(`・『${r.title}』（期限: ${r.dueDate}）`);
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