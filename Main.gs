//---------------開啟觸發---------------
function onOpen() {
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('崩壞精準')
    .addSubMenu(
      ui.createMenu('切換精準池')
        .addItem('當期A池', 'switchCurrentUpPoolA')
        .addItem('當期B池', 'switchCurrentUpPoolB')
        .addItem('上期A池', 'switchPreviousUpPoolA')
        .addItem('上期B池', 'switchPreviousUpPoolB')
     )
    .addSeparator()
    .addItem('重補分析註解', 'addAnalysisNotes')
    .addItem('清除分析註解', 'clearAnalysisNotes')
    .addItem('復原分析格式', 'restoreAnalysisFormat')
    .addSeparator()
    .addItem('同步精準池（個）', 'synchronousUpPool')
    .addItem('同步WIKI評分（個）', 'synchronousEvaluations')
    .addItem('更新個人評分（個）', 'updateEvaluations')
    .addItem('還原個人評分（個）', 'restoreEvaluations')
    .addSeparator()
    .addItem('寫入精準UP（編）', 'writeUpPool')
    .addToUi();
}

//---------------編輯觸發---------------
function onEdit() {
  
  var aSheet = SpreadsheetApp.getActiveSheet();
  var aCell = aSheet.getActiveCell();
  var aColumn = aCell.getColumn();
  var aRow = aCell.getRow();
  
  //----------武器兩層下拉式選單----------
  var editListRange = ( aColumn == 3 && 2 < aRow && aRow < 11 && aSheet.getName() == '精準分析' );
  if (editListRange) 
  {
    var range = aSheet.getRange(aCell.getRow(), aColumn + 1);
    var sourceRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(aCell.getValue());
    setDataValid(range,sourceRange);
  }
  
  //----------精準池選擇----------
  var editPoolRange = ( aColumn == 2 && aRow == 2  && aSheet.getName() == '精準分析' );
  if(editPoolRange) {
    var poolName = getAnalysisPoolTitleRange().getValue();
    switch (poolName) {
      case '當期A':
        switchCurrentUpPoolA();
        break;
      case '當期B':
        switchCurrentUpPoolB();
        break;
      case '上期A':
        switchPreviousUpPoolA();
        break;
      case '上期B':
        switchPreviousUpPoolB();
        break;
      default:
        break;
    }
  }
  
  //----------編輯精準池修補註解----------
  var editAnalysisRange = ( 4 <= aColumn && aColumn <=5 && 3 <= aRow && aRow <=10 && aSheet.getName() == '精準分析' );
  if (editAnalysisRange)
  {
    setUpPoolTitle('被修改');
    sleep(500);
    addAnalysisNotes();
  }
  
}

//---------------設定資料驗證清單---------------
function setDataValid(range, sourceRange) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).build();
  range.setDataValidation(rule);
}

//---------------切換當期精準A池---------------
function switchCurrentUpPoolA() {
  var copyRange = getCurrentUpPoolARange();
  var title = '當期A';
  switchUpPoolProcess(copyRange, title);
}

//---------------切換當期精準B池---------------
function switchCurrentUpPoolB() {
  var copyRange = getCurrentUpPoolBRange();
  var title = '當期B';
  switchUpPoolProcess(copyRange, title);
}

//---------------切換上期精準A池---------------
function switchPreviousUpPoolA() {
  var copyRange = getPreviousUpPoolARange();
  var title = '上期A';
  switchUpPoolProcess(copyRange, title);
}

//---------------切換上期精準B池---------------
function switchPreviousUpPoolB() {
  var copyRange = getPreviousUpPoolBRange();
  var title = '上期B';
  switchUpPoolProcess(copyRange, title);
}

//---------------切換精準池流程---------------
function switchUpPoolProcess(copyRange,title) {
  setUpPoolTitle('修改中');
  switchUpPool(copyRange);
  addAnalysisNotes();
  setUpPoolTitle(title);
}

//---------------精準標題---------------
function setUpPoolTitle(name) {
  var analysisSheet = getLotteryAnalysisSheet();
  getAnalysisPoolTitleRange().setValue(name);
}

//---------------切換一期精準---------------
function switchUpPool(copyRange){
  var pasteRange = getAnalysisPoolRange();
  var data = copyRange.getValues();
  pasteRange.setValues(data);
  
  var validRange;
  var sourceRange;
  var skip = 2;
  var len = pasteRange.getValues().length - skip;
  for(var i=0; i<len ; i++){
    validRange = pasteRange.getCell(i+1+skip,2);
    sourceRange = getNameRange(data[i+skip][0]);
    setDataValid(validRange,sourceRange);
  }
}

//---------------同步WIKI評分---------------
function synchronousEvaluations() {
  var copyRange = getOnlineEvaluationsRange();
  var pasteRange = getEvaluationsRange();
  copyPasteData(copyRange,pasteRange);
  pasteRange.setNotes(copyRange.getNotes());
  
  addAnalysisNotes();
}

//---------------同步精準池---------------
function synchronousUpPool() {
  var copyRange = getOnlineUpPoolALLRange();
  var pasteRange = getUpPoolALLRange();
  copyPasteData(copyRange,pasteRange);
  switchCurrentUpPoolA();
}

//---------------更新個人評分資料---------------
function updateEvaluations() {
  var copyRange = getPersonalEvaluationsRange();
  var pasteRange = getBackupPersonalEvaluationsRange();
  copyPasteData(copyRange,pasteRange);
  sleep(250);
  pasteRange = copyRange;
  copyRange = getSynchronousPersonalEvaluationsRange();
  copyPasteData(copyRange,pasteRange);
}

//---------------還原個人評分---------------
function restoreEvaluations() {
  var copyRange = getBackupPersonalEvaluationsRange();
  var pasteRange = getPersonalEvaluationsRange();
  copyPasteData(copyRange,pasteRange);
}

//---------------補上分析備註---------------
function addAnalysisNotes() {
  
  clearAnalysisNotes();
  
  var evaluationsSheet = getEvaluationsSheet();
  var posData = getAnalysisDataRange().getValues();
  
  var notes = new Array(posData.length);
  
  for(var k=0; k<3; k++) {
    for(var i = 0; i < posData.length; i++) {
      notes[i] = [];
      if(posData[i][k*2] > 0) {
        var note = evaluationsSheet.getRange(posData[i][k*2], posData[i][k*2+1]).getNote();
        notes[i][0] = note;
      } else {
        notes[i][0] = "";
      }
    }
    getAnalysisNoteRange(k).setNotes(notes);
  }
  
}

//---------------清除分析備註---------------
function clearAnalysisNotes() {
  for(var i=0; i<3; i++) {
    var editRange = getAnalysisNoteRange(i);
    editRange.clearNote();
  }
}

//---------------復原分析格式---------------
function restoreAnalysisFormat() {
  var analysisSheet = getLotteryAnalysisSheet();
  
  var fontSizes = setValueLoop(12);
  var fontFamilies = setValueLoop('Microsoft JhengHei');
  
  var backgrounds = [];
  backgrounds[0] = setValueLoop('#d9ead3');
  backgrounds[1] = setValueLoop('#cfe2f3');
  backgrounds[2] = setValueLoop('#f9cb9c');
  
  for(var i=0; i<3; i++) {
    var pos = getSetting(i*2+8);
    var editRange = getRangeBySetting(analysisSheet,pos);
    var formulas = getSettingFormulas(2,i+6,22,1);
    editRange.setFormulas(formulas);
    editRange.setFontSizes(fontSizes);
    editRange.setFontFamilies(fontFamilies);
    editRange.setBackgrounds(backgrounds[i]);
  }
}

//---------------寫入精準UP---------------
function writeUpPool() {
  var analysisPoolRange = getAnalysisPoolRange();
  var currentUpPoolARange = getCurrentUpPoolARange();
  var PreviousUpPoolARange = getPreviousUpPoolARange();
  var currentUpPoolBRange = getCurrentUpPoolBRange();
  var PreviousUpPoolBRange = getPreviousUpPoolBRange();
  
  copyPasteData(currentUpPoolARange,PreviousUpPoolARange);
  copyPasteData(currentUpPoolBRange,PreviousUpPoolBRange);
  var copyData = copyPasteData(analysisPoolRange,currentUpPoolARange);
  
  copyData[0][1] = AddDay(copyData[0][1],3);
  copyData[0][2] = AddDay(copyData[0][2],3);
  var tmp = copyData[2];
  copyData[2] = copyData[3];
  copyData[3] = tmp;
  
  currentUpPoolBRange.setValues(copyData);
}




