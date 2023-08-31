function verificaImpegni() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("foglio_impegni"); // Inserire il nome del foglio di lavoro
  var lastRow = sheet.getLastRow();
  var nomeImpegno, dataImpegno, oraImpegno, tipoPromemoria, mailReceiver;
  for (var row = 2; row <= lastRow; row++) {
    nomeImpegno = sheet.getRange(row, 1).getValue();
    if (nomeImpegno === "") {
      break;
    }
    dataImpegno = sheet.getRange(row, 2).getValue();
    oraImpegno = sheet.getRange(row, 3).getValue();
    tipoPromemoria = sheet.getRange(row, 4).getValue();
    mailReceiver = sheet.getRange(row, 5).getValue();
    
    var now = new Date();
    var oneHour = 60 * 60 * 1000;
    var tomorrow = new Date(now.getTime() + 24 * oneHour);

    if(tipoPromemoria==0){
      continue;
    }
    
    if (tipoPromemoria == 5 && isTomorrow(dataImpegno)) {
      invioMail(nomeImpegno, dataImpegno, oraImpegno, mailReceiver);
      sheet.getRange(row, 4).setValue(4);
    } 
    else if (((tipoPromemoria == 6)) && isNextWeek(dataImpegno)) {
      invioMail(nomeImpegno, dataImpegno, oraImpegno, mailReceiver);
      sheet.getRange(row, 4).setValue(tipoPromemoria-1);
    }
    else if (((tipoPromemoria == 4) || (tipoPromemoria == 3) || (tipoPromemoria == 2) || (tipoPromemoria == 1) ) && isToday(dataImpegno) && isWithinHours(oraImpegno, now, oneHour*4)) {
      invioMail(nomeImpegno, dataImpegno, oraImpegno, mailReceiver);
      sheet.getRange(row, 4).setValue(tipoPromemoria-1);
  }
}
}

function isTomorrow(date) {
  var now = new Date();
  var tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  return isSameDate(date, tomorrow);
}

function isNextWeek(date) {
  var now = new Date();
  var nextWeek = new Date(now.getTime() + 7* 24 * 60 * 60 * 1000);
  return isSameDateOrBelow(date, nextWeek);
}


function isToday(date) {
  var now = new Date();
  return isSameDate(date, now);
}

function isWithinHours(timeString, referenceDate, numMilliseconds) {
  //se la data oraImpegno di ora è tra meno 
  var nowPlusMillis = new Date(referenceDate.getTime() + numMilliseconds);
  var time = new Date(nowPlusMillis.toDateString() + timeString); //il bug è qui: la data non viene definita
  return nowPlusMillis >= time;
}

function isWithinMinutes(timeString, referenceDate, numMinutes) {
  var now = new Date(referenceDate.getTime() + numMinutes * 60 * 1000);
  var time = new Date(now.toDateString() + " " + timeString);
  return now >= time;
}

function isSameDate(date1, date2) {
  return date1.getFullYear() == date2.getFullYear() && 
         date1.getMonth() == date2.getMonth() && 
         date1.getDay() == date2.getDay();
}

function isSameDateOrBelow(date1, date2) {
  return ((date1.getFullYear() == date2.getFullYear() && 
         date1.getMonth() == date2.getMonth() && 
         date1.getDate() == date2.getDate() )
         || 
         (date1.getFullYear() == date2.getFullYear() && 
         date1.getMonth() == date2.getMonth() && 
         date1.getDate() < date2.getDate() )
         ||
         (date1.getFullYear() == date2.getFullYear() && 
         date1.getMonth() < date2.getMonth()));
}


function invioMail(nomeImpegno, dataImpegno, oraImpegno, mailReceiver) {
  var oggetto = nomeImpegno;
  var descrizione = "Ricordati che in data " + Utilities.formatDate(dataImpegno, "GMT+2", "dd/MM/yyyy") + " alle ore " + oraImpegno + " hai il seguente impegno: " + nomeImpegno;

    MailApp.sendEmail(mailReceiver, oggetto, descrizione);
}


