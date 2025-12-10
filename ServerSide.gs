/**
 * Google Apps Script for Milestone Manager
 * Features: WBS Search & Info Extraction (Updated Column Mapping)
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_PROJECTS = 'Projects';

const TEMPLATE_WBS_ID = '1T7QAlk5rxKE_6-oOZvf8XAkmxuqCpMPu6LPo7vFy4Dc'; 
const DEST_FOLDER_ID = '1GeIiZezt8EF6rwEIVBDVgoSEE_rIlfjw'; 
const SETTING_SOURCE_ID = '1Dxp7XahAOJEXYcUxuyOJH-xTnHVecLNGYeRNPGZbnSc'; 

// --- WBS設定 ---
const WBS_SHEET_NAME = 'Schedule'; 
const WBS_PROGRESS_CELL = 'M3';    

const WBS_COST_SHEET_NAME = '工数管理'; 
const WBS_MANDAYS_CELL = 'E2';        
const WBS_MANMONTHS_CELL = 'F2';      

// ★タスク検索設定
const TARGET_TASKS = [
  { key: 'revInternal', name: '定義レビュー(社内)', shortName: '社内レビュー' },
  { key: 'revMurasys',  name: '定義レビュー(ムラシス)', shortName: 'ムラシスレビュー' },
  { key: 'testInteg',   name: '統合試験', shortName: '統合試験' },
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

function apiSyncWbsProgress() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pSheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  
  const projects = getRowsAsObjects(pSheet);
  
  projects.forEach(p => {
    if (p.wbsUrl) {
      try {
        const wbsSs = SpreadsheetApp.openByUrl(p.wbsUrl);
        let hasUpdates = false;

        // 1. 進捗率
        const wbsSheet = wbsSs.getSheetByName(WBS_SHEET_NAME);
        if (wbsSheet) {
          const val = wbsSheet.getRange(WBS_PROGRESS_CELL).getValue();
          if (p.wbsProgress != val) {
            p.wbsProgress = val;
            hasUpdates = true;
          }

          // ★タスク情報の抽出
          const data = wbsSheet.getDataRange().getValues();
          
          TARGET_TASKS.forEach(target => {
            // タスク名で検索（H列）
            const foundRow = data.find(row => String(row[WBS_COL_TASK_NAME] || '').includes(target.name));
            
            if (foundRow) {
              // 予定日の決定ロジック: P列があればP列、なければO列
              const rawPlanDate = foundRow[WBS_COL_PLAN_DATE_P] ? foundRow[WBS_COL_PLAN_DATE_P] : foundRow[WBS_COL_PLAN_DATE_O];
              
              const planDate = formatDate(rawPlanDate);
              const actualDate = formatDate(foundRow[WBS_COL_ACTUAL_DATE]); // Q列
              const status = String(foundRow[WBS_COL_STATUS] || '');        // L列
              const assignee = String(foundRow[WBS_COL_ASSIGNEE] || '');    // K列
              
              const kPlan = `task_${target.key}_plan`;
              const kActual = `task_${target.key}_actual`;
              const kStatus = `task_${target.key}_status`;
              const kAssignee = `task_${target.key}_assignee`;
              
              if (p[kPlan] !== planDate) { p[kPlan] = planDate; hasUpdates = true; }
              if (p[kActual] !== actualDate) { p[kActual] = actualDate; hasUpdates = true; }
              if (p[kStatus] !== status) { p[kStatus] = status; hasUpdates = true; }
              if (p[kAssignee] !== assignee) { p[kAssignee] = assignee; hasUpdates = true; }
            }
          });
        }

        // 2. 工数
        const costSheet = wbsSs.getSheetByName(WBS_COST_SHEET_NAME);
        if (costSheet) {
          const md = costSheet.getRange(WBS_MANDAYS_CELL).getValue();
          const mm = costSheet.getRange(WBS_MANMONTHS_CELL).getValue();
          
          if (p.manDays != md) { p.manDays = md; hasUpdates = true; }
          if (p.manMonths != mm) { p.manMonths = mm; hasUpdates = true; }
        }

        if (hasUpdates) {
          saveRow(pSheet, p); 
        }

      } catch (e) {
        console.warn(`Failed to fetch WBS data for ${p.name}: ${e.message}`);
      }
    }
  });

  return {
    projects: getRowsAsObjects(pSheet)
  };
}

function apiCreateProject(project) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  
  if (!project.id) {
    project.id = Utilities.getUuid();
    project.createdAt = new Date().toISOString();
  }

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

  saveRow(sheet, project);
  return apiGetData();
}

function apiUpdateProject(project) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_PROJECTS);
  saveRow(sheet, project);
  return apiGetData();
}

// --- Helper Functions ---

function formatDate(dateVal) {
  if (!dateVal) return '';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }
  return String(dateVal); 
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function getRowsAsObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  if (data.length === 1 && data[0].length === 1 && data[0][0] === '') return [];

  const headers = data[0];
  const idIndex = headers.indexOf('id');
  
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (idIndex !== -1 && !row[idIndex]) continue;
    
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      const header = headers[j];
      if (!header) continue;
      let val = row[j];
      if (val instanceof Date) val = val.toISOString();
      obj[header] = (val === undefined || val === null) ? '' : val;
    }
    results.push(obj);
  }
  return results;
}

function saveRow(sheet, obj) {
  const data = sheet.getDataRange().getValues();
  let headers = [];
  const isSheetEmpty = data.length === 0 || (data.length === 1 && data[0].length === 1 && data[0][0] === '');

  if (!isSheetEmpty) {
    headers = data[0];
    const newKeys = Object.keys(obj).filter(k => !headers.includes(k));
    if (newKeys.length > 0) {
      const lastCol = sheet.getLastColumn();
      const startCol = lastCol === 0 ? 1 : lastCol + 1;
      sheet.getRange(1, startCol, 1, newKeys.length).setValues([newKeys]);
      headers = [...headers, ...newKeys];
    }
  } else {
    const keys = Object.keys(obj);
    headers = ['id', ...keys.filter(k => k !== 'id')];
    if (data.length === 1) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
        sheet.appendRow(headers);
    }
  }
  
  let rowIndex = -1;
  let idColIndex = headers.indexOf('id');
  
  if (idColIndex === -1) {
    idColIndex = headers.length;
    headers.push('id');
    sheet.getRange(1, headers.length).setValue('id');
  }
  
  if (!isSheetEmpty && data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === obj.id) {
        rowIndex = i + 1;
        break;
      }
    }
  }
  
  const rowToSave = headers.map(h => {
    const val = obj[h];
    if (val instanceof Date) return val.toISOString();
    return val === undefined || val === null ? '' : val;
  });
  
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowToSave.length).setValues([rowToSave]);
  } else {
    sheet.appendRow(rowToSave);
  }
}