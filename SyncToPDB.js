function syncToPDB() {
  var srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시세수집');
  var lastRow = srcSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('반영할 데이터가 없습니다.');
    return;
  }

  // 시세수집 전체 데이터 가져오기 (A열부터 K열까지: 총 11개 열)
  var allData = srcSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var processedCount = 0;
  var skippedNoModel = 0;
  var skippedPass = 0;

  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    var rowIndex = i + 2;

    var brand = String(row[1]).trim();
    var model = String(row[2]).trim();
    var year = String(row[3]).trim();
    var type = String(row[4]).trim();
    var size = String(row[5]).trim();
    var alreadyDone = row[8]; // I열 (반영 여부/시간)
    var memo = String(row[10] || '').trim(); // K열 (주석 컬럼)

    // 이미 반영된 항목은 스킵
    if (alreadyDone) continue;

    // K열 주석에 '패스'가 포함되어 있으면 PDB 반영 건너뜀
    if (memo.indexOf('패스') !== -1) {
      skippedPass++;
      continue;
    }

    if (!model || model.indexOf('미등록추정') !== -1) {
      skippedNoModel++;
      continue;
    }

    var isSpin = type.indexOf('스피닝') !== -1;
    var targetYear = year || '';

    // 1. '시세수집' 전체 데이터 중에서 현재 품목과 스펙이 완전히 일치하는 항목들의 가격 수집
    var matchedPrices = [];
    for (var k = 0; k < allData.length; k++) {
      var kRow = allData[k];
      var kBrand = String(kRow[1]).trim();
      var kModel = String(kRow[2]).trim();
      var kYear = String(kRow[3]).trim();
      var kType = String(kRow[4]).trim();
      var kSize = String(kRow[5]).trim();
      var kPrice = Number(kRow[6]);

      var kIsSpin = kType.indexOf('스피닝') !== -1;
      var manuMatch = kBrand === brand;
      var modelMatch = kModel === model;
      var yearMatch = (kYear === targetYear);
      var sizeMatch = isSpin ? (kSize === size) : true;
      var typeMatch = (isSpin === kIsSpin);

      if (manuMatch && modelMatch && yearMatch && sizeMatch && typeMatch && !isNaN(kPrice) && kPrice > 0) {
        matchedPrices.push(kPrice);
      }
    }

    if (matchedPrices.length === 0) continue;

    // 2. 최근 10개만 추출
    var recentPrices = matchedPrices.slice(-10);
    var countUsed = recentPrices.length;

    // 3. 최저가, 평균가(1천원 단위 반올림 후 만원 맞춤), 최고가 산출
    var minPriceRaw = Math.min.apply(null, recentPrices);
    var maxPriceRaw = Math.max.apply(null, recentPrices);
    var sum = recentPrices.reduce(function(a, b) { return a + b; }, 0);
    var rawAvg = sum / countUsed;

    var minPrice = Math.round(minPriceRaw / 10000) * 10000;
    var avgPrice = Math.round(rawAvg / 10000) * 10000;
    var maxPrice = Math.round(maxPriceRaw / 10000) * 10000;

    // 4. PDB 시트에 반영 
    // [스피닝 SpinPDB 구조 변경 예시]: A:제조사, B:모델명, C:연식, D:사이즈, E:최저가, F:평균가, G:최고가, H:평균 자료 갯수
    // [베이트 BaitPDB 구조 변경 예시]: A:제조사, B:모델명, C:연식, D:최저가, E:평균가, F:최고가, G:평균 자료 갯수
    var pdbSheetName = isSpin ? 'SpinPDB' : 'BaitPDB';
    var pdbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pdbSheetName);
    if (!pdbSheet) continue;

    var pdbLastRow = pdbSheet.getLastRow();
    var pdbData = pdbLastRow >= 2 ? pdbSheet.getRange(2, 1, pdbLastRow - 1, isSpin ? 4 : 3).getValues() : [];

    var matchedRow = -1;
    for (var j = 0; j < pdbData.length; j++) {
      var pRow = pdbData[j];
      var manuMatch = String(pRow[0]).trim() === brand;
      var modelMatch = String(pRow[1]).trim() === model;
      var yearMatch = String(pRow[2]).trim() === targetYear;
      var sizeMatch = isSpin ? String(pRow[3]).trim() === size : true;

      if (manuMatch && modelMatch && yearMatch && sizeMatch) {
        matchedRow = j + 2;
        break;
      }
    }

    if (matchedRow !== -1) {
      if (isSpin) {
        // SpinPDB 열 매핑: E(최저), F(평균), G(최고), H(개수)
        pdbSheet.getRange(matchedRow, 5).setValue(minPrice);
        pdbSheet.getRange(matchedRow, 6).setValue(avgPrice);
        pdbSheet.getRange(matchedRow, 7).setValue(maxPrice);
        pdbSheet.getRange(matchedRow, 8).setValue(countUsed + '개');
      } else {
        // BaitPDB 열 매핑: D(최저), E(평균), F(최고), G(개수)
        pdbSheet.getRange(matchedRow, 4).setValue(minPrice);
        pdbSheet.getRange(matchedRow, 5).setValue(avgPrice);
        pdbSheet.getRange(matchedRow, 6).setValue(maxPrice);
        pdbSheet.getRange(matchedRow, 7).setValue(countUsed + '개');
      }
    } else {
      var newRow = pdbSheet.getLastRow() + 1;

      if (isSpin) {
        pdbSheet.getRange(newRow, 1, 1, 8).setValues([[brand, model, targetYear, size, minPrice, avgPrice, maxPrice, countUsed + '개']]);
      } else {
        pdbSheet.getRange(newRow, 1, 1, 7).setValues([[brand, model, targetYear, minPrice, avgPrice, maxPrice, countUsed + '개']]);
      }
    }

    // 5. 시세수집 I열에 처리 일시 및 건수 기록
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var logText = formattedDate + ' (' + countUsed + '개)';
    
    srcSheet.getRange(rowIndex, 9).setValue(logText);
    processedCount++;
  }

  SpreadsheetApp.getUi().alert(
    '반영 완료 (최저/평균/최고가 제안 적용): ' + processedCount + '건\n' +
    '주석 패스 건너뜀: ' + skippedPass + '건\n' +
    '모델명 미확정 건너뜀: ' + skippedNoModel + '건'
  );
}