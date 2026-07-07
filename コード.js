function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('図書室 督促システム V12')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 現在のスプレッドシート内のシート名一覧を取得
function getSpreadsheetSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

// 列番号をアルファベットに変換するヘルパー関数
function getColumnLetter(colNum) {
  let letter = "";
  while (colNum > 0) {
    let temp = (colNum - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colNum = (colNum - temp - 1) / 26;
  }
  return letter;
}

// 指定したシートの1行目（見出し行）を取得し、重複があれば列記号を付与する
function getSheetHeaders(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  let lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  let counts = {};
  headers.forEach(h => {
    let name = String(h).trim() || "(空欄)";
    counts[name] = (counts[name] || 0) + 1;
  });
  
  return headers.map((h, idx) => {
    let name = String(h).trim() || "(空欄)";
    let letter = getColumnLetter(idx + 1);
    let text = name;
    if (counts[name] > 1 || name === "(空欄)") {
      text = `${name}(${letter}列)`;
    }
    return { index: idx + 1, text: text };
  });
}

// TSVデータから新しいシートを作成する
function createSheetFromTSV(sheetName, tsvData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(sheetName)) {
    throw new Error(`シート「${sheetName}」は既に存在します。別の名前を指定してください。`);
  }
  
  let sheet = ss.insertSheet(sheetName);
  let rows = tsvData.trim().split('\n');
  let values = rows.map(row => row.split('\t'));
  
  if (values.length > 0 && values[0].length > 0) {
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  }
  return `シート「${sheetName}」を新規作成し、名簿データを登録しました。`;
}

// 基本設定の保存
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

// ルール設定の保存と読み込み
function saveRules(rulesArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('システム設定_ルール');
  if (!sheet) sheet = ss.insertSheet('システム設定_ルール');
  sheet.clearContents();
  sheet.appendRow(['ルールデータ']);
  sheet.appendRow([JSON.stringify(rulesArray)]);
}

function loadRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('システム設定_ルール');
  if (!sheet) return [];
  let val = sheet.getRange(2, 1).getValue();
  return val ? JSON.parse(val) : [];
}

// 教職員専用名簿を取得する関数
function loadStaffRoster() {
  const maps = loadStaffRosterMaps();
  let map = {};
  for (let id in maps) map[id] = maps[id].email;
  return map;
}

function loadStaffRosterMaps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('教職員名簿');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let map = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      map[String(data[i][0])] = {
        email: data[i][2] ? String(data[i][2]).trim() : "",
        name: data[i][1] ? String(data[i][1]).trim() : ""
      };
    }
  }
  return map;
}

function registerStaffEmail(id, name, email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('教職員名簿');
  if (!sheet) {
    sheet = ss.insertSheet('教職員名簿');
    sheet.appendRow(['4桁ID', '氏名', 'メールアドレス']);
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 3).setValue(email);
      return name + " 先生のメールアドレスを更新しました。";
    }
  }
  sheet.appendRow([id, name, email]);
  return name + " 先生のメールアドレスを新規登録しました。";
}

function normalizeName(name) {
  return String(name || "").trim().replace(/[\s　]+/g, "");
}

function normalizeTo4DigitId(classInfo, name) {
  let match = classInfo.match(/(中学|高校).*?(\d+)年([A-Za-z0-9]+)組(\d+)番/);
  if (!match) return "STAFF_" + name.replace(/[\s ]+/g, ""); 
  return `${match[2]}${match[3]}${String(match[4]).padStart(2, '0')}`;
}

// 日付文字列から貸出年度（4月始まり）を算出
function normalizeDateString(input) {
  if (input === null || input === undefined) return "";

  if (Object.prototype.toString.call(input) === "[object Date]" && !isNaN(input.getTime())) {
    const y = input.getFullYear();
    const m = String(input.getMonth() + 1).padStart(2, "0");
    const d = String(input.getDate()).padStart(2, "0");
    return `${y}/${m}/${d}`;
  }

  const raw = String(input).trim();
  if (!raw) return "";

  // "2026年4月9日", "2026.4.9", "2026-04-09", "2026/4/9" を統一
  const replaced = raw
    .replace(/[年月]/g, "/")
    .replace(/日/g, "")
    .replace(/[.\-]/g, "/")
    .replace(/\s+/g, "");

  const parts = replaced.split("/").filter(Boolean);
  if (parts.length < 3) return "";

  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  const d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return "";
  if (m < 1 || m > 12 || d < 1 || d > 31) return "";

  return `${y}/${String(m).padStart(2, "0")}/${String(d).padStart(2, "0")}`;
}

function getNendoFromDateStr(dateStr) {
  const norm = normalizeDateString(dateStr);
  if (!norm) return null;
  let parts = norm.split("/");
  let y = parseInt(parts[0], 10);
  let m = parseInt(parts[1], 10);
  if (isNaN(y) || isNaN(m)) return null;
  return m <= 3 ? y - 1 : y;
}

function parseRuleDate(dateStr) {
  const ms = parseDateMs(dateStr);
  return new Date(ms);
}

function parseDateMs(dateStr) {
  const norm = normalizeDateString(dateStr);
  if (!norm) return null;
  let parts = norm.split("/");
  if (parts.length < 3) return null;
  let y = parseInt(parts[0], 10);
  let m = parseInt(parts[1], 10);
  let d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
  return new Date(y, m - 1, d).getTime();
}

function getRuleNendoLabel(rule) {
  if (!rule || !rule.start) return "";
  let nendo = getNendoFromDateStr(rule.start);
  return nendo != null ? `${nendo}年度` : "";
}

// 1件の貸出データに対して適用ルールを解決
function resolveRuleForRecord(rules, borrowDateStr, dueDateStr) {
  const bDate = parseDateMs(borrowDateStr);
  const dDate = parseDateMs(dueDateStr);
  if (bDate === null || dDate === null) {
    return { error: "日付形式不正", matchedRules: [] };
  }

  let matchedRules = [];
  for (let i = 0; i < rules.length; i++) {
    let rule = rules[i];
    let dateType = rule.dateType || "borrow";
    let targetDate = dateType === "due" ? dDate : bDate;
    let start = parseRuleDate(rule.start).getTime();
    let end = parseRuleDate(rule.end);
    end.setHours(23, 59, 59, 999);
    if (targetDate >= start && targetDate <= end.getTime()) {
      matchedRules.push(rule);
    }
  }

  if (matchedRules.length > 1) {
    return { error: "名簿ルール重複", matchedRules };
  }
  return { rule: matchedRules[0] || null, bDate, dDate, matchedRules };
}

// ルールに基づいて特定シートの名簿Mapを生成する（学年・組・番号の4桁IDで照合）
function getRosterMapsForRule(rule) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(rule.sheetName);
  if (!sheet) return null;
  
  let data = sheet.getDataRange().getValues();
  let byId = {};
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let id = "";
    
    if (rule.idCol) {
      let val = row[parseInt(rule.idCol) - 1];
      if(val) id = String(val).trim();
    } 
    else if (rule.gradeCol && rule.classCol && rule.numCol) {
      let g = String(row[parseInt(rule.gradeCol) - 1] || "").trim();
      let c = String(row[parseInt(rule.classCol) - 1] || "").trim();
      let n = String(row[parseInt(rule.numCol) - 1] || "").padStart(2, '0');
      if (g && c && n) id = `${g}${c}${n}`;
    }
    
    let email = rule.emailCol ? String(row[parseInt(rule.emailCol) - 1] || "").trim() : "";
    let rosterName = rule.nameCol ? String(row[parseInt(rule.nameCol) - 1] || "").trim() : "";
    
    if (id && email) byId[id] = { email: email, rosterName: rosterName };
  }
  return { byId };
}

// 名簿からメールアドレスを照合（参照名簿のIDで照合）
function lookupInRoster(maps, id) {
  if (id && maps.byId[id]) {
    return {
      email: maps.byId[id].email || "",
      rosterName: maps.byId[id].rosterName || ""
    };
  }
  return { email: "", rosterName: "" };
}

function buildMailPlan(records, settings, rules) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let grouped = { overdue: {}, scheduled: [] };
  let previewRows = [];
  let missingEmails = new Set();
  let rosterCache = {};

  records.forEach(r => {
    const borrowDateNorm = normalizeDateString(r.borrowDate);
    const dueDateNorm = normalizeDateString(r.dueDate);
    let isStaff = !r.classInfo.match(/(中学|高校)/);
    let borrowNendo = getNendoFromDateStr(borrowDateNorm);
    let lookupId = normalizeTo4DigitId(r.classInfo, r.name);
    r.id = lookupId;

    let email = "";
    let rosterName = "";
    let ruleMatched = false;
    let matchedRule = null;
    let matchError = "";
    let resolved = resolveRuleForRecord(rules, borrowDateNorm, dueDateNorm);

    if (isStaff) {
      if (!rosterCache['STAFF']) rosterCache['STAFF'] = loadStaffRosterMaps();
      let staffInfo = rosterCache['STAFF'][lookupId] || { email: "", name: "" };
      email = staffInfo.email || "";
      rosterName = staffInfo.name || r.name;
      ruleMatched = true;
    } else {
      if (resolved.error) {
        matchError = resolved.error;
      } else if (resolved.rule) {
        matchedRule = resolved.rule;
        ruleMatched = true;
        let cacheKey = matchedRule.sheetName;
        if (!rosterCache[cacheKey]) {
          let maps = getRosterMapsForRule(matchedRule);
          rosterCache[cacheKey] = maps !== null ? maps : { byId: {} };
        }
        let found = lookupInRoster(rosterCache[cacheKey], lookupId);
        email = found.email;
        rosterName = found.rosterName || "";
      }
    }

    r.email = email;

    let isAllowed = (r.classInfo.includes("中学") && String(settings.sendJ) === "true") ||
                    (r.classInfo.includes("高校") && String(settings.sendH) === "true") ||
                    (isStaff && String(settings.sendStaff) === "true");

    let dueMs = parseDateMs(dueDateNorm);
    let isOverdue = dueMs !== null && dueMs < today.getTime();

    if (isOverdue) {
      let status = "";
      let statusDetail = "";
      if (!isAllowed) {
        status = "対象外";
        statusDetail = "送信設定オフ";
      } else if (!isStaff && matchError === "名簿ルール重複") {
        status = "送信不可";
        statusDetail = "名簿ルール重複";
      } else if (!isStaff && matchError === "日付形式不正") {
        status = "送信不可";
        statusDetail = "日付形式不正";
      } else if (!ruleMatched) {
        status = "送信不可";
        statusDetail = "名簿ルールなし";
      } else if (!email) {
        status = "送信不可";
        statusDetail = isStaff ? "メール未登録" : "学年・組・番未一致";
      } else {
        status = "送信予定";
        statusDetail = "";
      }

      previewRows.push({
        borrowName: r.name,
        rosterName: rosterName,
        matchId: lookupId,
        email: email,
        status: status,
        statusDetail: statusDetail,
        title: r.title,
        classInfo: r.classInfo,
        borrowDate: borrowDateNorm || r.borrowDate,
        dueDate: dueDateNorm || r.dueDate,
        nendo: borrowNendo != null ? String(borrowNendo) : (r.nendo || ""),
        rosterSheet: isStaff ? "教職員名簿" : (matchedRule ? matchedRule.sheetName : ""),
        rosterNendo: isStaff ? "教職員" : (matchedRule ? getRuleNendoLabel(matchedRule) : ""),
        id: lookupId,
        isStaff: isStaff,
        willSend: status === "送信予定"
      });
    }

    if (!isAllowed) return;

    if (!email) {
      if (isStaff) {
        missingEmails.add(`教職員: ${r.name} (プレビュー画面の氏名をクリックして登録可能)`);
      } else if (matchError === "名簿ルール重複") {
        missingEmails.add(`生徒: ${r.classInfo} ${r.name} (貸出日: ${r.borrowDate} に対して名簿ルールが重複しています。期間設定を見直してください)`);
      } else if (matchError === "日付形式不正") {
        missingEmails.add(`生徒: ${r.classInfo} ${r.name} (貸出日または返却日の形式が不正です: 貸出=${r.borrowDate}, 返却=${r.dueDate})`);
      } else if (!ruleMatched) {
        missingEmails.add(`生徒: ${r.classInfo} ${r.name} (貸出日: ${borrowDateNorm || r.borrowDate} に対応する名簿ルールがありません)`);
      } else {
        missingEmails.add(`生徒: ${r.classInfo} ${r.name} (名簿「${matchedRule.sheetName}」に照合ID[${lookupId}]が見つかりません)`);
      }
    }

    if (email || isStaff) {
      if (isOverdue) {
        let groupKey = email || ('NOEMAIL_' + normalizeName(r.name) + '_' + r.borrowDate);
        if (!grouped.overdue[groupKey]) grouped.overdue[groupKey] = { user: r, books: [] };
        grouped.overdue[groupKey].books.push(`・『${r.title}』（期限: ${r.dueDate}）`);
      } else {
        grouped.scheduled.push(r);
      }
    }
  });

  return { grouped, previewRows, missingEmails };
}

function previewOverdueMails(payload) {
  const plan = buildMailPlan(payload.records, payload.settings, payload.rules);
  return plan.previewRows;
}

// 送信と予約のメイン処理
function processMails(payload) {
  const records = payload.records;
  saveSettings(payload.settings);
  saveRules(payload.rules);
  
  const settings = payload.settings;
  const rules = payload.rules;
  let subOverdue = settings.subOverdue || "【図書室】未返却図書のお知らせ";

  const plan = buildMailPlan(records, settings, rules);
  const grouped = plan.grouped;
  const missingEmails = plan.missingEmails;

  let sentCount = 0;
  for (let id in grouped.overdue) {
    let data = grouped.overdue[id];
    if (data.user.email) {
      let compiledSubject = subOverdue.replace(/【氏名】/g, data.user.name)
                                      .replace(/【クラス】/g, data.user.classInfo || "")
                                      .replace(/【書籍名】/g, data.books.join(' '));
      
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
    msg += `\n\n⚠️ 【注意】以下の ${missingEmails.size} 件はメールアドレスが見つからずスキップされました：\n` 
         + Array.from(missingEmails).join('\n');
  }
  return msg;
}

// 毎朝の自動実行用（変更なし）
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
      let groupKey = row[6] || ('NOEMAIL_' + normalizeName(String(row[1])));
      if (!groupedReminder[groupKey]) {
        groupedReminder[groupKey] = { name: row[1], classInfo: row[2], email: row[6], books: [] };
      }
      groupedReminder[groupKey].books.push(`・『${row[3]}』（期限: ${row[5]}）`);
    } else if (dueDate > tomorrow) {
      newData.push(row);
    }
  }

  let settings = loadSettings() || {};
  let subReminder = settings.subReminder || "【図書室】返却期限間近のお知らせ";
  let tplReminder = settings.tplReminder || "【氏名】 様\n\n図書室です。\n以下の図書の返却期限が【明日】となっております。\n\n【書籍名】\n\n期限内のご返却をお願いいたします。";

  for (let id in groupedReminder) {
    let user = groupedReminder[id];
    if (user.email) {
      let compiledSubject = subReminder.replace(/【氏名】/g, user.name)
                                       .replace(/【クラス】/g, user.classInfo || "")
                                       .replace(/【書籍名】/g, user.books.join(' '));
      
      let compiledBody = tplReminder.replace(/【氏名】/g, user.name)
                                    .replace(/【クラス】/g, user.classInfo || "")
                                    .replace(/【書籍名】/g, user.books.join('\n'));
                                    
      GmailApp.sendEmail(user.email, compiledSubject, compiledBody);
    }
  }
  
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}