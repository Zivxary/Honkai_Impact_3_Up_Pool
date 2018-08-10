//---------------取得命名範圍---------------
function getNameRange(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getRangeByName(name);
}

//---------------取得表格---------------
function getMySheetByName(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

//---------------取得線上版表格---------------
function getOnlineSheetByName(name) {
  var id = "1ZTldkLA36q6eDyjryIywD1AeCUCCc6GpVTz9HpPz1ZE";
  var spreadSheet = SpreadsheetApp.openById(id);
  return spreadSheet.getSheetByName(name);
}

//---------------取得 "精準分析" 表格---------------
function getLotteryAnalysisSheet() {
  return getMySheetByName('精準分析');
}

//---------------取得 "WIKI評分" 表格---------------
function getEvaluationsSheet() {
  return getMySheetByName('WIKI評分');
}

//---------------取得 線上"WIKI評分" 表格---------------
function getOnlineEvaluationsSheet() {
  return getOnlineSheetByName('WIKI評分');
}

//---------------取得 "評分備份" 表格---------------
function getEvaluationsBackUpSheet() {
  return getMySheetByName('評分備份');
}

//---------------取得 "精準UP" 表格---------------
function getUpPoolSheet() {
  return getMySheetByName('精準UP');
}

//---------------取得 線上"精準UP" 表格---------------
function getOnlineUpPoolSheet() {
  return getOnlineSheetByName('精準UP');
}

//---------------取得 "個人評分" 表格---------------
function getPersonalEvaluationsSheet() {
  return getMySheetByName('個人評分');
}

//---------------取得 "同步用" 表格---------------
function getSynchronousSheet() {
  return getMySheetByName('同步用');
}

//---------------取得 "GAS設定" ---------------
function getSetting(row) {
  return getMySheetByName('GAS設定').getRange(row,1,1,4).getValues();
}

//---------------取得 "GAS設定" 預設公式---------------
function getSettingFormulas(row,column,height,width) {
  return getMySheetByName('GAS設定').getRange(row,column,height,width).getFormulas();
}

//---------------取得範圍資料---------------
function getRangeBySetting(aSheet, pos) {
  return aSheet.getRange(pos[0][0], pos[0][1], pos[0][2], pos[0][3]);
}

//---------------取得 "精準分析" 備註欄位 第N列 範圍---------------
function getAnalysisNoteRange(num) {
  var pos = getSetting(num*2+8);
  var aSheet = getLotteryAnalysisSheet();
  return getRangeBySetting(aSheet,pos);
}

//---------------取得 "精準分析" 備註資料 範圍---------------
function getAnalysisDataRange() {
  var pos = getSetting(6);
  var aSheet = getLotteryAnalysisSheet();
  return getRangeBySetting(aSheet,pos);
}

//---------------取得 "WIKI評分" 備註資料 範圍---------------
function getEvaluationsDataRange() {
  var pos = getSetting(6);
  var aSheet = getEvaluationsSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 精準分析名稱 範圍---------------
function getAnalysisPoolTitleRange() {
  var pos = getSetting(22);
  var aSheet = getLotteryAnalysisSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 精準分析池 範圍---------------
function getAnalysisPoolRange() {
  var pos = getSetting(16);
  var aSheet = getLotteryAnalysisSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 精準當期A池 範圍---------------
function getCurrentUpPoolARange() {
  var pos = getSetting(17);
  var aSheet = getUpPoolSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 精準當期B池 範圍---------------
function getCurrentUpPoolBRange() {
  var pos = getSetting(18);
  var aSheet = getUpPoolSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 精準上期A池 範圍---------------
function getPreviousUpPoolARange() {
  var pos = getSetting(19);
  var aSheet = getUpPoolSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 精準上期B池 範圍---------------
function getPreviousUpPoolBRange() {
  var pos = getSetting(20);
  var aSheet = getUpPoolSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 備份個人評分 範圍---------------
function getBackupPersonalEvaluationsRange() {
  var pos = getSetting(24);
  var aSheet = getSynchronousSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 個人評分 範圍---------------
function getPersonalEvaluationsRange() {
  var pos = getSetting(28);
  var aSheet = getPersonalEvaluationsSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 同步個人評分 範圍---------------
function getSynchronousPersonalEvaluationsRange() {
  var pos = getSetting(26);
  var aSheet = getSynchronousSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 "精準UP" 範圍---------------
function getUpPoolALLRange() {
  var pos = getSetting(30);
  var aSheet = getUpPoolSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 線上"精準UP" 範圍---------------
function getOnlineUpPoolALLRange() {
  var pos = getSetting(30);
  var aSheet = getOnlineUpPoolSheet();
  return getRangeBySetting(aSheet,pos);
}

//---------------取得 "WIKI評分" 範圍---------------
function getEvaluationsRange() {
  var pos = getSetting(4);
  var aSheet = getEvaluationsSheet();
  return getRangeBySetting(aSheet,pos);
}

//--------------取得 線上"WIKI評分" 範圍---------------
function getOnlineEvaluationsRange() {
  var pos = getSetting(4);
  var aSheet = getOnlineEvaluationsSheet();
  return getRangeBySetting(aSheet,pos);
}

