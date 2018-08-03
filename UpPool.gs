//---------------開啟觸發---------------
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  
  menuEntries.push({name: "切換精準AB池", functionName: "switchUpPoolAB"});
  menuEntries.push({name: "切換上期精準", functionName: "switchPreviousUpPool"});
  menuEntries.push({name: "切換當期精準", functionName: "switchCurrentUpPool"});
  menuEntries.push({name: "寫入精準UP（編）", functionName: "writeUpPool"});
  menuEntries.push(null);
  menuEntries.push({name: "重補分析註解", functionName: "addAnalysisNotes"});
  menuEntries.push({name: "清除分析註解", functionName: "clearAnalysisNotes"});
  menuEntries.push({name: "復原分析格式", functionName: "restoreAnalysisFormat"});
  menuEntries.push(null);
  menuEntries.push({name: "備分評分（編）", functionName: "backUpEvaluations"});
  menuEntries.push({name: "復原評分", functionName: "restoreEvaluations"});
  
  ss.addMenu("崩壞精準", menuEntries);
}

//---------------編輯觸發---------------
function onEdit() {
  
  var aSheet = SpreadsheetApp.getActiveSheet();
  var aCell = aSheet.getActiveCell();
  var aColumn = aCell.getColumn();
  var aRow = aCell.getRow();
  
  var editListRange = ( aColumn == 3 && aRow > 2 && aRow < 11 && aSheet.getName() == '精準分析' );
  if (editListRange) 
  {
    var range = aSheet.getRange(aCell.getRow(), aColumn + 1);
    var sourceRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(aCell.getValue());
    setDataValid_(range,sourceRange);
  }
  
  var editAnalysisRange = ( 4 <= aColumn && aColumn <=5 && 3 <= aRow && aRow <=10 && aSheet.getName() == '精準分析' );
  if (editAnalysisRange)
  {
    Utilities.sleep(1000);
    addAnalysisNotes();
  }
  
}

//---------------設定資料驗證清單---------------
function setDataValid_(range, sourceRange) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).build();
  range.setDataValidation(rule);
}

//---------------取得表格---------------
function getMySheetByName(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

//---------------取得"精準分析"表格---------------
function getLotteryAnalysisSheet() {
  return getMySheetByName('精準分析');
}

//---------------取得"評分"表格---------------
function getEvaluationsSheet() {
  return getMySheetByName('評分');
}

//---------------取得"評分BK"表格---------------
function getEvaluationsBackUpSheet() {
  return getMySheetByName('評分BK');
}

//---------------取得"評分BK"表格---------------
function getUpPoolSheet() {
  return getMySheetByName('精準UP');
}

//---------------取得GAS設定---------------
function getSetting(row) {
  return getMySheetByName('GAS設定').getRange(row,1,1,4).getValues();
}

function getSettingFormulas(row,column,height,width) {
  return getMySheetByName('GAS設定').getRange(row,column,height,width).getFormulas();
}

//---------------取得範圍資料---------------
function getRangeBySetting(aSheet, pos) {
  return aSheet.getRange(pos[0][0], pos[0][1], pos[0][2], pos[0][3]);
}

//---------------取得命名範圍---------------
function getNameRange(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getRangeByName(name);
}

function AddDay(now,days) {
  var aDay = 1000*60*60*24;
  var date = new Date(now);
  var newDate = new Date(date.getTime() + aDay * (days+1));
  var output = Utilities.formatDate(newDate,"GMT",'yyyy/MM/dd');
  return output;
}

//---------------切換精準池---------------
function switchUpPoolAB() {
  var pos = getSetting(14);
  var sheet = getLotteryAnalysisSheet();
  var range = getRangeBySetting(sheet,pos);
  var data = range.getValues();
  
  var tmp = data[0];
  data[0] = data[1];
  data[1] = tmp;
  range.setValues(data);
  Logger.log(data);
  
  var validRange = range.getCell(1,2);
  var sourceRange = getNameRange(data[0][0]);
  setDataValid_(validRange,sourceRange);
  
  validRange = range.getCell(2,2);
  sourceRange = getNameRange(data[1][0])
  setDataValid_(validRange,sourceRange);
  
  addAnalysisNotes();
}

//---------------切換上期精準---------------
function switchPreviousUpPool() {
  switchUpPool(19);
  addAnalysisNotes();
}

//---------------切換當期精準---------------
function switchCurrentUpPool() {
  switchUpPool(17);
  addAnalysisNotes();
}

//---------------切換一期精準---------------
function switchUpPool(setNum){
  var upPoolSheet = getUpPoolSheet();
  var lotteryAnalysisSheet = getLotteryAnalysisSheet();
  var setting;
  var copyRange
  var pasteRange;
  
  setting = getSetting(setNum);
  copyRange = getRangeBySetting(upPoolSheet,setting);
  setting = getSetting(16);
  pasteRange = getRangeBySetting(lotteryAnalysisSheet,setting);
  var data = copyRange.getValues();
  pasteRange.setValues(data);
  
  var validRange;
  var sourceRange;
  var skip = 2;
  var len = pasteRange.getValues().length - skip;
  for(var i=0; i<len ; i++){
    validRange = pasteRange.getCell(i+1+skip,2);
    sourceRange = getNameRange(data[i+skip][0]);
    setDataValid_(validRange,sourceRange);
  }
}

//---------------寫入精準UP---------------
function writeUpPool() {
  var upPoolSheet = getUpPoolSheet();
  var lotteryAnalysisSheet = getLotteryAnalysisSheet();
  var setting;
  var copyRange
  var pasteRange;
  
  copyPasteData(upPoolSheet,17,upPoolSheet,19);
  copyPasteData(upPoolSheet,18,upPoolSheet,20);
  var copyData = copyPasteData(lotteryAnalysisSheet,16,upPoolSheet,17);
  
  copyData[0][1] = AddDay(copyData[0][1],3);
  copyData[0][2] = AddDay(copyData[0][2],3);
  var tmp = copyData[2];
  copyData[2] = copyData[3];
  copyData[3] = tmp;
  
  setting = getSetting(18);
  pasteRange = getRangeBySetting(upPoolSheet,setting);
  pasteRange.setValues(copyData);
}

//---------------複製貼上數值---------------
function copyPasteData(copySheet,copyNum,pasteSheet,pasteNum) {
  var setting;
  var copyRange
  var pasteRange;
  
  setting = getSetting(copyNum);
  copyRange = getRangeBySetting(copySheet,setting);
  setting = getSetting(pasteNum);
  pasteRange = getRangeBySetting(pasteSheet,setting);
  var copyData = copyRange.getValues()
  pasteRange.setValues(copyRange.getValues());
  
  return copyData;
}

//---------------備份評分---------------
function backUpEvaluations() {
  var aSheet = getEvaluationsSheet();
  var bSheet = getEvaluationsBackUpSheet();
  
  CopyEvaluations(aSheet,bSheet);
}

//---------------復原評分---------------
function restoreEvaluations() {
  var aSheet = getEvaluationsBackUpSheet();
  var bSheet = getEvaluationsSheet();
  
  CopyEvaluations(aSheet,bSheet);
}

//---------------複製評分---------------
function CopyEvaluations(copySheet, pasteSheet) {
  
  var pos = getSetting(4);
  //Logger.log(setting);
  
  var copySheetRange = getRangeBySetting(copySheet,pos);
  var pasteSheetRange = getRangeBySetting(pasteSheet,pos);
  
  var copyValues = copySheetRange.getValues();
  pasteSheetRange.setValues(copyValues);
  
  copyValues = copySheetRange.getNotes();
  pasteSheetRange.setNotes(copyValues);
}

//---------------補上分析備註---------------
function addAnalysisNotes() {
  
  clearAnalysisNotes();
  
  var lotterySheet = getLotteryAnalysisSheet();
  var evaluationsSheet = getEvaluationsSheet();
  var notePos = getSetting(6);
  var posData = getRangeBySetting(lotterySheet,notePos).getValues();
  
  var notes = new Array(posData.length);
  
  for(var k=0; k<3; k++) {
    for(var i = 0; i < posData.length; i++) {
      notes[i] = [];
      if(posData[i][k*2] > 0) {
        var note = evaluationsSheet.getRange(posData[i][k*2],posData[i][k*2+1]).getNote();
        notes[i][0] = note;
      }
      else{
        notes[i][0] = "";
      }
    }
    var pos = getSetting(k*2+8);
    getRangeBySetting(lotterySheet,pos).setNotes(notes);
  }
  
}

//---------------清除分析備註---------------
function clearAnalysisNotes() {
  var aSheet = getLotteryAnalysisSheet();
  for(var i=0; i<3; i++) {
    var setting = getSetting(i*2+8);
    var editRange = getRangeBySetting(aSheet,setting);
    editRange.clearNote();
  }
}

//---------------復原分析格式---------------
function restoreAnalysisFormat() {
  var aSheet = getLotteryAnalysisSheet();
  
  var fontSizes = setValueLoop(12);
  var fontFamilies = setValueLoop('Microsoft JhengHei');
  
  var backgrounds = [];
  backgrounds[0] = setValueLoop('#d9ead3');
  backgrounds[1] = setValueLoop('#cfe2f3');
  backgrounds[2] = setValueLoop('#f9cb9c');
  
  for(var i=0; i<3; i++) {
    var setting = getSetting(i*2+8);
    var editRange = getRangeBySetting(aSheet,setting);
    var setText = getSettingFormulas(2,i+6,22,1);
    editRange.setFormulas(setText);
    editRange.setFontSizes(fontSizes);
    editRange.setFontFamilies(fontFamilies);
    editRange.setBackgrounds(backgrounds[i]);
  }
}

//---------------設置 陣列[22][0]---------------
function setValueLoop(value) {
  var array = [];
  for(var i=0; i<22; i++) {
    array[i] = [];
    array[i][0] = value ;
  }
  return array;
}

