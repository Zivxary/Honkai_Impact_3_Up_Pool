//---------------增加天數---------------
function AddDay(now,days) {
  var aDay = 1000*60*60*24;
  var date = new Date(now);
  var newDate = new Date(date.getTime() + aDay * (days+1));
  return Utilities.formatDate(newDate,"GMT",'yyyy/MM/dd');
}

//---------------設置 陣列[22][0] 都等於 value---------------
function setValueLoop(value) {
  var array = [];
  for(var i=0; i<22; i++) {
    array[i] = [];
    array[i][0] = value ;
  }
  return array;
}

//---------------複製貼上數值---------------
function copyPasteData(copyRange,pasteRange) {
  var copyData = copyRange.getValues()
  pasteRange.setValues(copyData);
  return copyData;
}

//---------------等待一段時間---------------
function sleep(time) {
  Utilities.sleep(time);
}