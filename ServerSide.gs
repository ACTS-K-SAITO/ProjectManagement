/**
 * Google Apps Script for Milestone Manager
 * Features: WBS Search & Info Extraction (Updated Column Mapping)
 * Added: Active User Tracking (Heartbeat)
 * Added: Estimate File Search (Recursive with !old exclusion)
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_PROJECTS = 'Projects';
const SHEET_ACTIVE_USERS = 'ActiveUsers'; // ★閲覧中ユーザー管理用シート

const TEMPLATE_WBS_ID = '1T7QAlk5rxKE_6-oOZvf8XAkmxuqCpMPu6LPo7vFy4Dc'; 
const DEST_FOLDER_ID = '1GeIiZezt8EF6rwEIVBDVgoSEE_rIlfjw'; 
const SETTING_SOURCE_ID = '1Dxp7XahAOJEXYcUxuyOJH-xTnHVecLNGYeRNPGZbnSc'; 

// ★見積書検索用フォルダID
const ESTIMATE_FOLDER_ID = '1LK2bIRQg9sqOhf2I4xZ8sXHReLEuKc92';

// --- WBS設定 ---
const WBS_SHEET_NAME = 'Schedule'; 
const WBS_PROGRESS_CELL = 'M3';    

const WBS_COST_SHEET_NAME = '工数管理'; 
const WBS_MANDAYS_CELL = 'E2';        
const WBS_MANMONTHS_CELL = 'F2';      

const ACTIVE_THRESHOLD_MS = 5 * 60 * 1000; // ★5分以内のアクセスを「閲覧中」とみなす

// ★タスク検索設定
const TARGET_TASKS = [
  { key: 'revInternal', name: '定義レビュー(社内)', shortName: '社内レビュー' },
  { key: 'revMurasys',  name: '定義レビュー(ムラシス)', shortName: 'ムラシスレビュー' },
  { key: 'testInteg',   name: 'リンクテスト(統合試験)', shortName: '統合試験' },
  { key: 'testActs',    name: 'ACTS社内検証', shortName: 'ACTS検証' },
  { key: 'testSystem',  name: 'システム検証', shortName: 'システム検証' },
  { key: 'test3rd',     name: '第三者検証', shortName: '第三者検証' }
];

// 列インデックス (A列=0, B列=1, ... H列=7, K=10, L=11, O=14, P=15, Q=16)
const WBS_COL_TASK_NAME = 7;    // H列: 検索対象(タスク名)
const WBS_COL_ASSIGNEE = 10;    // K列: 担当者
const WBS_COL_STATUS = 11;      // L列: ステータス
const WBS_COL_PLAN_DATE_O = 14; // O列: 予定日(標準)
const WBS_COL_PLAN_DATE_P = 15; // P列: 予定日(優先)
const WBS_COL_ACTUAL_DATE = 16; // Q列: 完了日(実績)

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ムラシスクリーン作業一覧2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- API Methods ---

function apiGetData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pSheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  
  return {
    projects: getRowsAsObjects(pSheet)
  };
}

// ★ Heartbeat Function (閲覧中ユーザーの更新と取得)
function apiHeartbeat() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_ACTIVE_USERS);
  const lock = LockService.getScriptLock();
  
  const email = Session.getActiveUser().getEmail() || "Unknown";
  const shortName = email.split('@')[0];
  const now = new Date();
  
  let activeUsers = [];

  try {
    if (lock.tryLock(3000)) {
      const data = sheet.getDataRange().getValues();
      let newData = [];
      let found = false;
      
      if (data.length === 0 || data[0][0] !== 'email') {
        newData.push(['email', 'name', 'lastSeen']);
      } else {
        newData.push(data[0]);
      }
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowTimeStr = row[2];
        if (!rowTimeStr) continue;
        
        const rowTime = new Date(rowTimeStr);
        if (now.getTime() - rowTime.getTime() <= ACTIVE_THRESHOLD_MS) {
          if (row[0] === email) {
            newData.push([email, shortName, now.toISOString()]);
            found = true;
          } else {
            newData.push(row);
          }
        }
      }
      
      if (!found) {
        newData.push([email, shortName, now.toISOString()]);
      }
      
      sheet.clearContents();
      if (newData.length > 0) {
        sheet.getRange(1, 1, newData.length, 3).setValues(newData);
      }
      
      activeUsers = newData.slice(1).map(row => row[1]);
      lock.releaseLock();
    } else {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
         if (data[i][2]) {
             const t = new Date(data[i][2]);
             if (now.getTime() - t.getTime() <= ACTIVE_THRESHOLD_MS) {
                 activeUsers.push(data[i][1]);
             }
         }
      }
      if (!activeUsers.includes(shortName)) activeUsers.push(shortName);
    }
    
  } catch (e) {
    console.warn("Heartbeat error: " + e.message);
    activeUsers.push(shortName); 
  }
  
  return { activeUsers: [...new Set(activeUsers)].sort() };
}

// ★ 見積ファイル検索API (サブフォルダ対応版)
function apiGetFileUrlByEstimateNo(estimateNo) {
  if (!estimateNo) return null;
  
  try {
    const rootFolder = DriveApp.getFolderById(ESTIMATE_FOLDER_ID);
    
    // 再帰的に検索を実行
    return findFileRecursive(rootFolder, estimateNo);
    
  } catch (e) {
    console.warn("Estimate file search error: " + e.message);
  }
  return null;
}

// 再帰検索用ヘルパー関数
function findFileRecursive(folder, targetName) {
  // 1. カレントフォルダ内でファイルを検索
  // DriveApp.searchFiles は検索精度（トークナイズ）の問題でヒットしない場合があるため、
  // getFiles() で全ファイルを取得して JavaScript 側で判定することで確実性を高める。
  try {
    const files = folder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      // ファイル名が検索語で始まるかチェック（前方一致）
      if (file.getName().startsWith(targetName)) {
        return file.getUrl(); // 見つかったらURLを返して終了
      }
    }
  } catch (e) {
    console.warn(`Search error in folder ${folder.getName()}: ${e.message}`);
  }
  
  // 2. サブフォルダを探索
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    
    // "!old" フォルダはスキップ
    if (subFolder.getName() === '!old') {
      continue;
    }
    
    // 再帰呼び出し
    const foundUrl = findFileRecursive(subFolder, targetName);
    if (foundUrl) {
      return foundUrl; // 見つかったらバブルアップして終了
    }
  }
  
  return null; // 見つからなかった
}

function apiAddMemo(projectId, text) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(5000);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return apiGetData();
    
    const headers = data[0];
    const idIdx = headers.indexOf('id');
    let descIdx = headers.indexOf('description');
    
    if (idIdx === -1) return apiGetData();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === projectId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex > 0) {
      if (descIdx === -1) {
        descIdx = headers.length;
        sheet.getRange(1, descIdx + 1).setValue('description');
      }
      
      const cell = sheet.getRange(rowIndex, descIdx + 1);
      const currentDesc = String(cell.getValue());
      const now = new Date();
      const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
      const email = Session.getActiveUser().getEmail() || "Unknown";
      const shortName = email.split('@')[0];
      
      const newEntry = `[${timeStr} ${shortName}]\n${text}`;
      const newDesc = currentDesc ? (newEntry + "\n\n" + currentDesc) : newEntry;
      
      cell.setValue(newDesc);
    }
  } finally {
    lock.releaseLock();
  }
  return apiGetData();
}

function apiEditMemo(projectId, oldEntryContent, newText) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('id');
    const descIdx = headers.indexOf('description');
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === projectId) { rowIndex = i + 1; break; }
    }
    if (rowIndex > 0) {
      const cell = sheet.getRange(rowIndex, descIdx + 1);
      const currentDesc = String(cell.getValue());
      if (currentDesc) {
        const entries = currentDesc.split(/(?=\[\d{4}\/\d{2}\/\d{2} \d{2}:\d{2} )/g);
        const targetIndex = entries.findIndex(e => e.trim() === oldEntryContent.trim());
        if (targetIndex !== -1) {
          const originalEntry = entries[targetIndex];
          const trailingSpace = originalEntry.replace(originalEntry.trimEnd(), '');
          const match = originalEntry.trim().match(/^\[(.*?)\]/);
          let newEntry = match ? match[0] + '\n' + newText : newText;
          entries[targetIndex] = newEntry + trailingSpace;
          cell.setValue(entries.join(''));
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
  return apiGetData();
}

function apiDeleteMemo(projectId, entryContent) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('id');
    const descIdx = headers.indexOf('description');
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === projectId) { rowIndex = i + 1; break; }
    }
    if (rowIndex > 0) {
      const cell = sheet.getRange(rowIndex, descIdx + 1);
      const currentDesc = String(cell.getValue());
      if (currentDesc) {
        const entries = currentDesc.split(/(?=\[\d{4}\/\d{2}\/\d{2} \d{2}:\d{2} )/g);
        const targetIndex = entries.findIndex(e => e.trim() === entryContent.trim());
        if (targetIndex !== -1) {
          entries.splice(targetIndex, 1);
          cell.setValue(entries.join('').trim());
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
  return apiGetData();
}

function apiDeleteProject(projectId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('id');
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === projectId) { sheet.deleteRow(i + 1); break; }
    }
  } finally {
    lock.releaseLock();
  }
  return apiGetData();
}

function apiSyncWbsProgress() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pSheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  
  // 最新のプロジェクト一覧を一度取得
  const projects = getRowsAsObjects(pSheet);
  
  projects.forEach(p => {
    if (p.wbsUrl) {
      try {
        const wbsSs = SpreadsheetApp.openByUrl(p.wbsUrl);
        let hasUpdates = false;
        
        // ★修正: saveRowによる上書きを防ぐため、このタイミングで最新の行データを再取得する
        const latestData = getRowsAsObjects(pSheet);
        const currentP = latestData.find(item => item.id === p.id) || p;

        const wbsSheet = wbsSs.getSheetByName(WBS_SHEET_NAME);
        if (wbsSheet) {
          const val = wbsSheet.getRange(WBS_PROGRESS_CELL).getValue();
          if (currentP.wbsProgress != val) { currentP.wbsProgress = val; hasUpdates = true; }
          const data = wbsSheet.getDataRange().getValues();
          TARGET_TASKS.forEach(target => {
            const foundRow = data.find(row => String(row[WBS_COL_TASK_NAME] || '').includes(target.name));
            if (foundRow) {
              const rawPlanDate = foundRow[WBS_COL_PLAN_DATE_P] ? foundRow[WBS_COL_PLAN_DATE_P] : foundRow[WBS_COL_PLAN_DATE_O];
              const planDate = formatDate(rawPlanDate);
              const actualDate = formatDate(foundRow[WBS_COL_ACTUAL_DATE]);
              const status = String(foundRow[WBS_COL_STATUS] || '');
              const assignee = String(foundRow[WBS_COL_ASSIGNEE] || '');
              const kPlan = `task_${target.key}_plan`;
              const kActual = `task_${target.key}_actual`;
              const kStatus = `task_${target.key}_status`;
              const kAssignee = `task_${target.key}_assignee`;
              if (currentP[kPlan] !== planDate) { currentP[kPlan] = planDate; hasUpdates = true; }
              if (currentP[kActual] !== actualDate) { currentP[kActual] = actualDate; hasUpdates = true; }
              if (currentP[kStatus] !== status) { currentP[kStatus] = status; hasUpdates = true; }
              if (currentP[kAssignee] !== assignee) { currentP[kAssignee] = assignee; hasUpdates = true; }
            }
          });
        }
        const costSheet = wbsSs.getSheetByName(WBS_COST_SHEET_NAME);
        if (costSheet) {
          const md = costSheet.getRange(WBS_MANDAYS_CELL).getValue();
          const mm = costSheet.getRange(WBS_MANMONTHS_CELL).getValue();
          if (currentP.manDays != md) { currentP.manDays = md; hasUpdates = true; }
          if (currentP.manMonths != mm) { currentP.manMonths = mm; hasUpdates = true; }
        }
        if (hasUpdates) saveRow(pSheet, currentP); 
      } catch (e) { console.warn(`WBS Sync Failed for ${p.name}: ${e.message}`); }
    }
  });
  return apiGetData();
}

function apiCreateProject(project) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  
  if (!project.id) {
    project.id = Utilities.getUuid();
    project.createdAt = new Date().toISOString();
  }

  if (project.type !== 'trip' && !project.skipWbs) {
    try {
      const templateFile = DriveApp.getFileById(TEMPLATE_WBS_ID);
      const destFolder = DriveApp.getFolderById(DEST_FOLDER_ID);
      
      const pwbsStr = project.pwbs || 'NoPWBS';
      const nameStr = project.name || 'NoName';
      const fileName = `WBS(${pwbsStr})${nameStr}`;
      
      const copiedFile = templateFile.makeCopy(fileName, destFolder);
      project.wbsUrl = copiedFile.getUrl();

      try {
        const sourceSpreadsheet = SpreadsheetApp.openById(SETTING_SOURCE_ID);
        const sourceSheet = sourceSpreadsheet.getSheetByName('setting');
        const targetSpreadsheet = SpreadsheetApp.openById(copiedFile.getId());
        const targetSheet = targetSpreadsheet.getSheetByName(WBS_SHEET_NAME);
        
        if (sourceSheet && targetSheet) {
          const val = sourceSheet.getRange('B3').getValue();
          targetSheet.getRange('G1').setValue(val);
        }
        
        if (targetSheet) {
           project.wbsProgress = targetSheet.getRange(WBS_PROGRESS_CELL).getValue();
        }
        const costSheet = targetSpreadsheet.getSheetByName(WBS_COST_SHEET_NAME);
        if (costSheet) {
           project.manDays = costSheet.getRange(WBS_MANDAYS_CELL).getValue();
           project.manMonths = costSheet.getRange(WBS_MANMONTHS_CELL).getValue();
        }

      } catch (e) {
        console.warn("WBS初期処理エラー: " + e.message);
      }

    } catch (e) {
      console.error("WBS作成失敗: " + e.message);
    }
  }

  saveRow(sheet, project);
  return apiGetData();
}

function apiUpdateProject(project) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  saveRow(sheet, project);
  return apiGetData();
}

// --- Helpers ---
function formatDate(dateVal) {
  if (!dateVal) return '';
  if (dateVal instanceof Date) return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
  return String(dateVal); 
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function getRowsAsObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const idIndex = headers.indexOf('id');
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (idIndex !== -1 && !row[idIndex]) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      if (val instanceof Date) val = val.toISOString();
      obj[headers[j]] = (val === undefined || val === null) ? '' : val;
    }
    results.push(obj);
  }
  return results;
}

function saveRow(sheet, obj) {
  const data = sheet.getDataRange().getValues();
  let headers = (data.length > 0) ? data[0] : [];
  const newKeys = Object.keys(obj).filter(k => !headers.includes(k));
  if (newKeys.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, newKeys.length).setValues([newKeys]);
    headers = [...headers, ...newKeys];
  }
  let rowIndex = -1;
  const idIdx = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === obj.id) { rowIndex = i + 1; break; }
  }
  const rowToSave = headers.map(h => {
    const val = obj[h];
    return (val instanceof Date) ? val.toISOString() : (val === undefined || val === null ? '' : val);
  });
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, rowToSave.length).setValues([rowToSave]);
  else sheet.appendRow(rowToSave);
}