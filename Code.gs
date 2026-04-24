function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('보드게임 대여 신청')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAllGames() {
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get('allGames_v8'); 
  
  if (cachedData) {
    return JSON.parse(cachedData);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var zones = ['1구역', '2구역', '3구역', '4구역'];
  var allGames = {};

  zones.forEach(function(zone) {
    var sheet = ss.getSheetByName(zone);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      data.shift(); 

      var gameList = data.map(function(row) {
        return {
          id: row[0],         
          location: row[1],   
          maker: row[2],      
          name: row[3],       
          total: row[4],      
          available: (row[5] !== "" && row[5] !== undefined) ? row[5] : row[4] 
        };
      }).filter(function(game) { return game.name && game.name !== ""; });
      
      allGames[zone] = gameList;
    }
  });
  
  cache.put('allGames_v8', JSON.stringify(allGames), 900);
  return allGames;
}

function submitRequest(school, teacherName, zone, gameName, count, returnDate, type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("대여현황");
  if(!sheet) throw new Error("'대여현황' 시트가 없습니다.");

  var reqId = "REQ_" + new Date().getTime();
  
  // 📌 대여일에도 텍스트 강제 방어막(') 추가!
  var todayDate = "'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var exactTime = "'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
  
  var applicantInfo = school + " / " + teacherName;
  
  sheet.insertRowBefore(2);
  var newRowData = [[reqId, type, applicantInfo, zone, gameName, count, todayDate, exactTime, returnDate]];
  sheet.getRange(2, 1, 1, 9).setValues(newRowData);
  
  if (sheet.getLastRow() > 2) {
    var sourceRange = sheet.getRange(3, 1, 1, sheet.getLastColumn()); 
    var targetRange = sheet.getRange(2, 1, 1, sheet.getLastColumn()); 
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  
  CacheService.getScriptCache().remove('allGames_v8');
  return "신청이 완료되었습니다! 😆";
}

function getCurrentlyRented() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("대여현황");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  data.shift(); 

  var rentedList = data.filter(function(row) {
    return row[1] === '대여 대기' || row[1] === '대여 중' || row[1] === '반납 대기'; 
  }).map(function(row) {
    
    var ts = 0;
    if (row[0] && row[0].toString().indexOf('REQ_') !== -1) {
      ts = Number(row[0].toString().split('_')[1]);
    }

    return {
      status: row[1], 
      gameName: row[4],
      applicant: row[2],
      count: row[5],
      returnDate: Utilities.formatDate(new Date(row[8]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      timestamp: ts 
    };
  });

  rentedList.sort(function(a, b) {
    return b.timestamp - a.timestamp;
  });

  return rentedList; 
}

function onEdit(e) {
  if (!e) return;
  
  var ss = e.source;
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== "대여현황") return; 
  
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  if (row <= 1) return; 
  
  if (col === 2) {
    var newValue = e.value;       
    var oldValue = e.oldValue;    
    
    var zone = sheet.getRange(row, 4).getValue(); 
    var game = sheet.getRange(row, 5).getValue(); 
    var count = sheet.getRange(row, 6).getValue(); 
    var applicantInfo = sheet.getRange(row, 3).getValue(); 
    
    if (!zone || !game) return; 
    
    if (newValue === '대여 중' && oldValue !== '대여 중') {
      updateInventoryForEdit(zone, game, count, '대여 중');
      CacheService.getScriptCache().remove('allGames_v8');
      ss.toast("✅ [" + game + "] 대여 승인! 재고가 차감되었습니다.", "승인 완료");
    }
    else if (newValue === '반납 완료' && oldValue !== '반납 완료') {
      updateInventoryForEdit(zone, game, count, '반납 완료');
      
      // 📌 반납일과 반납 시간 모두 텍스트 강제 방어막(') 추가!
      var returnTodayDate = "'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      var returnExactTime = "'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
      
      var targetSheet = ss.getSheetByName("반납기록");
      if (!targetSheet) {
        targetSheet = ss.insertSheet("반납기록");
        targetSheet.appendRow(["상태", "학교/이름", "구역", "보드게임명", "대여개수", "반납일", "반납 시간"]);
      }
      
      var newRowData = [
        '반납 완료', 
        applicantInfo,
        zone, 
        game, 
        count, 
        returnTodayDate, 
        returnExactTime
      ];
      
      targetSheet.appendRow(newRowData); 
      
      var targetLastRow = targetSheet.getLastRow();
      if (targetLastRow > 2) {
        var tSource = targetSheet.getRange(targetLastRow - 1, 1, 1, targetSheet.getLastColumn());
        var tTarget = targetSheet.getRange(targetLastRow, 1, 1, targetSheet.getLastColumn());
        tSource.copyTo(tTarget, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
        tSource.copyTo(tTarget, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
      
      sheet.deleteRow(row);
      
      CacheService.getScriptCache().remove('allGames_v8');
      ss.toast("🔄 [" + game + "] 반납 처리 완료! [반납기록] 시트로 저장되었습니다.", "반납 완료");
    }
  }
}

function updateInventoryForEdit(zone, gameName, count, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(zone);
  if(!sheet) return;
  
  var tf = sheet.createTextFinder(gameName).matchEntireCell(true).findNext();
  if (!tf) return;
  
  var row = tf.getRow();
  
  var currentAvailable = sheet.getRange(row, 6).getValue();
  if (currentAvailable === "") {
    currentAvailable = sheet.getRange(row, 5).getValue(); 
  }
  currentAvailable = Number(currentAvailable);
  
  if (newStatus === '대여 중') {
    sheet.getRange(row, 6).setValue(currentAvailable - Number(count));
  } else if (newStatus === '반납 완료') {
    sheet.getRange(row, 6).setValue(currentAvailable + Number(count));
  }
}
