function sortPDB() {
  sortSheet('SpinPDB', true);
  sortSheet('BaitPDB', false);
  SpreadsheetApp.getUi().alert('정렬 완료');
}

function sortSheet(sheetName, isSpin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // 데이터 1건 이하면 정렬 불필요

  // ★ 변경된 열 범위 반영 (SpinPDB: 8열 / BaitPDB: 7열)
  // [SpinPDB]: 제조사(0), 모델명(1), 연식(2), 사이즈(3), 최저가(4), 평균가(5), 최고가(6), 개수(7)
  // [BaitPDB]: 제조사(0), 모델명(1), 연식(2), 최저가(3), 평균가(4), 최고가(5), 개수(6)
  var numCols = isSpin ? 8 : 7; 
  var range = sheet.getRange(2, 1, lastRow - 1, numCols);
  var data = range.getValues();

  data.sort(function (a, b) {
    return compareRows(a, b, isSpin);
  });

  range.setValues(data);
}

function compareRows(a, b, isSpin) {
  // 1. 제조사 우선순위 및 가나다순 정렬 (최우선)
  var brandA = String(a[0]).trim();
  var brandB = String(b[0]).trim();
  
  var priorityBrands = ['다이와', '시마노', '도요', '아부가르시아'];
  var indexA = priorityBrands.indexOf(brandA);
  var indexB = priorityBrands.indexOf(brandB);

  if (indexA !== -1 && indexB !== -1) {
    var cmp = indexA - indexB;
    if (cmp !== 0) return cmp;
  } else if (indexA !== -1) {
    return -1;
  } else if (indexB !== -1) {
    return 1;
  } else {
    var cmp = brandA.localeCompare(brandB, 'ko');
    if (cmp !== 0) return cmp;
  }

  // 2. 평균 시세 가격 내림차순 (SpinPDB는 평균가 F열=인덱스 5, BaitPDB는 E열=인덱스 4)
  var priceIdx = isSpin ? 5 : 4; 
  var priceA = Number(a[priceIdx]) || 0;
  var priceB = Number(b[priceIdx]) || 0;
  var cmp = priceB - priceA;
  if (cmp !== 0) return cmp;

  // 3. 모델명 가나다순 (가격이 같을 경우 모델명순)
  var modelA = String(a[1]).trim();
  var modelB = String(b[1]).trim();
  cmp = modelA.localeCompare(modelB, 'ko');
  if (cmp !== 0) return cmp;

  // 4. 연식 최신순 (연식없음은 맨 뒤)
  var yearA = parseYear(a[2]);
  var yearB = parseYear(b[2]);
  if (yearA === null && yearB === null) cmp = 0;
  else if (yearA === null) cmp = 1;
  else if (yearB === null) cmp = -1;
  else cmp = yearB - yearA;
  if (cmp !== 0) return cmp;

  // 5. 사이즈 오름차순 (스피닝만, D열=인덱스 3)
  if (isSpin) {
    var sizeA = parseSizeNum(a[3]);
    var sizeB = parseSizeNum(b[3]);
    cmp = sizeA - sizeB;
  }
  return cmp;
}

function parseYear(val) {
  var s = String(val).trim();
  if (!s || s.indexOf('연식없음') !== -1) return null;
  var n = parseInt(s, 10);
  return isNaN(n) ? null : n;
}

function parseSizeNum(val) {
  var s = String(val).trim();
  var m = s.match(/\d+/);
  if (!m) return Infinity;
  return parseInt(m[0], 10);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('릴케어팜 도구')
    .addItem('PDB에 반영', 'syncToPDB')
    .addItem('PDB 정렬', 'sortPDB')
    .addToUi();
}