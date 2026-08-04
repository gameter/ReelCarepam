// [낚시릴 품번 및 모델명 주요 알파벳 약어 정리]
// 1. [공통 (시마노 & 다이와)]
//    - 바디/핸들: C(컴팩트 바디), DH(더블 핸들)
//    - 스풀규격: S(쉘로우스풀), SS(슈퍼 쉘로우 스풀)
// 
// 2. [시마노(Shimano) 고유]
//    - 바디/핸들: SW(솔트워터 빅게임용)
//    - 기어비: PG(파워기어), HG(하이기어), XG(엑스트라 하이기어)
//    - 스풀규격: M(미디움스풀), SSS(스페셜 슈퍼 쉘로우)
//    - 핵심기술: DC(디지털 컨트롤), MGL(매그넘라이트스플), CI4+(카본 소재), XD(내구성 강화)
//    - 스타일: BFS(베이트피네스스타일)   
// 
// 3. [다이와(Daiwa) 고유]
//    - 구동/컨셉: TW(T-Wing), SV(백래시 방지), LT(Light & Tough), SF(슈퍼 피네스), FC(피네스 커스텀), PC(파워 커스텀), HLC(장타), AIR(초경량), HD(헤비듀티), ST(송어계류용 맥실드 없이 무중력릴링)
//    - 기어비: P(파워기어), H(하이기어), XH(엑스트라 하이기어)
//    - 스풀규격: D(딥스풀)
//    - 부속: QD(퀵드랙)
//    - 바디: A(알리미늄)
//    - 세대: 숫자(1~3)

const SUFFIX_WHITELIST = [
  // 공통 약어
  'C', 'DH', 'S', 'SS',
  
  // 시마노(Shimano) 고유 약어
  'SW', 'PG', 'HG', 'XG', 'M', 'SSS', 'DC', 'MGL', 'CI4+', 'XD','BFS',
  
  // 다이와(Daiwa) 고유 약어
  'TW', 'SV', 'LT', 'SF', 'FC', 'PC', 'HLC', 'AIR', 'HD', 'P', 'H', 'XH', 'D', 'QD', 'A', 'ST'
];

function getSheetNameByType(type) {
  return (type.indexOf('스피닝') !== -1) ? 'SpinPDB' : 'BaitPDB';
}

function getBaseModelDict(brand, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
  let models = [];
  data.forEach(function(r) {
    if (String(r[0]).trim() === String(brand).trim() && r[1]) {
      models.push(String(r[1]).trim());
    }
  });
  models = [...new Set(models)].sort(function(a, b) { return b.length - a.length; });
  Logger.log('brand=%s, sheetName=%s, found models=%s', brand, sheetName, JSON.stringify(models));
  return models;
}


function findModelAcrossAllBrands(title, sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return null;
  var data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();

  var candidates = [];
  data.forEach(function(r) {
    if (r[0] && r[1]) {
      candidates.push({ brand: String(r[0]).trim(), model: String(r[1]).trim() });
    }
  });
  // 중복 제거 + 긴 모델명 우선 (짧은 이름이 잘못 걸리는 것 방지)
  var seen = {};
  candidates = candidates.filter(function(c) {
    var key = c.brand + '|' + c.model;
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
  candidates.sort(function(a, b) { return b.model.length - a.model.length; });

  for (var i = 0; i < candidates.length; i++) {
    if (title.indexOf(candidates[i].model) !== -1) {
      return candidates[i]; // { brand: '시마노', model: '알데바란' }
    }
  }
  return null;
}


function extractFullModelName(title, brand, sheetName) {
  var models = getBaseModelDict(brand, sheetName);
  var baseModel = '';
  var restIndex = -1;
  var correctedBrand = brand;

  for (var i = 0; i < models.length; i++) {
    var idx = title.indexOf(models[i]);
    if (idx !== -1) { baseModel = models[i]; restIndex = idx + models[i].length; break; }
  }

  // 원래 브랜드로 못 찾았으면, 전체 브랜드에서 역검색
  if (!baseModel) {
    var found = findModelAcrossAllBrands(title, sheetName);
    if (found) {
      baseModel = found.model;
      correctedBrand = found.brand;
      restIndex = title.indexOf(found.model) + found.model.length;
    }
  }

  if (!baseModel) return { fullName: '', correctedBrand: brand };

  var restTokens = title.slice(restIndex).trim().split(/\s+/);
  var fullName = baseModel;

  for (var j = 0; j < restTokens.length; j++) {
    var token = restTokens[j];
    var isWhitelisted = SUFFIX_WHITELIST.some(function(s) {
      return token.toUpperCase().startsWith(s.replace(' ', ''));
    });
    var isSpecCombo = /^\d+[A-Za-z]+\+?$/.test(token) || (/^[A-Za-z]+\+?$/.test(token) && token.length <= 4);

    if (isWhitelisted || isSpecCombo) {
      fullName += ' ' + token;
    } else {
      break;
    }
  }

  return { fullName: fullName, correctedBrand: correctedBrand };
}



function extractArticleId(url) {
  if (!url) return '';
  // 네이버 카페 URL에서 아티클 번호(숫자) 추출 (예: .../articles/123456)
  var match = String(url).match(/\/articles\/(\d+)/);
  if (match) return match[1];
  // 만약 다른 형태의 URL이라면 숫자로 된 마지막 경로를 추출하거나 그대로 반환
  return String(url).trim();
}

// 기존 isAlreadyCollected 함수를 대체합니다 (URL과 제목 모두 검사)
function checkDuplicate(url, title, sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { isDuplicate: false, reason: null };
  
  // J열(URL)과 H열(제목) 데이터를 한 번에 가져옴 (범위: H2:J)
  var dataRange = sheet.getRange("H2:J" + lastRow).getValues(); 
  var targetId = extractArticleId(url);
  var trimmedTitle = title ? String(title).trim() : '';

  for (var i = 0; i < dataRange.length; i++) {
    var savedTitle = dataRange[i][0]; // H열 (Index 0)
    var savedUrl = dataRange[i][2];   // J열 (Index 2)

    // 1. URL(게시글 번호) 기준 중복 체크
    if (savedUrl && targetId) {
      var savedId = extractArticleId(savedUrl);
      if (savedId === targetId) {
        return { isDuplicate: true, reason: "DUPLICATE_URL" };
      }
    }

    // 2. 제목 기준 중복 체크 (공백 제거 후 비교)
    if (savedTitle && trimmedTitle) {
      if (String(savedTitle).trim() === trimmedTitle) {
        return { isDuplicate: true, reason: "DUPLICATE_TITLE" };
      }
    }
  }
  
  return { isDuplicate: false, reason: null }; // ✅ 여기가 false여야 정상 수집됩니다!
}

function extractModelPositionally(title, brand, year) {
  var remaining = title;

  if (brand) remaining = remaining.split(brand).join('').trim();

  if (year) {
    var re = new RegExp('(?<!\\d)' + year + '(?!\\d)');
    remaining = remaining.replace(re, '').trim();
  }

  var tokens = remaining.split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return '';

  var modelName = tokens[0];
  for (var i = 1; i < tokens.length; i++) {
    var token = tokens[i];
    var isWhitelisted = SUFFIX_WHITELIST.some(function(s) {
      return token.toUpperCase().startsWith(s.replace(' ', ''));
    });
    var isSpecCombo = /^\d+[A-Za-z]+\+?$/.test(token) || (/^[A-Za-z]+\+?$/.test(token) && token.length <= 4);
    if (isWhitelisted || isSpecCombo) {
      modelName += ' ' + token;
    } else {
      break;
    }
  }
  return modelName;
}


// doPost 내부의 중복 검사 로직 수정
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시세수집');
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({result: "ERROR_NO_SHEET"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // URL 및 제목 중복 검사 실행
    var dupCheck = checkDuplicate(data.url, data.title, sheet);
    if (dupCheck.isDuplicate) {
      return ContentService.createTextOutput(JSON.stringify({result: dupCheck.reason}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var pdbSheetName = getSheetNameByType(data.type);
    var result = extractFullModelName(data.title, data.brand, pdbSheetName);

    if (!result.fullName) {
      var guessed = extractModelPositionally(data.title, data.brand, data.year);
      if (guessed) {
        result.fullName = guessed + ' (미등록추정)';
      }
    }

    var upperModelName = result.fullName ? result.fullName.toUpperCase() : '';
    var upperSize = data.size ? String(data.size).toUpperCase() : '';
    var rawYear = data.year ? String(data.year).replace(/[^0-9]/g, '') : '';
    var nextRow = sheet.getLastRow() + 1;

    sheet.getRange(nextRow, 1, 1, 10).setValues([[
      new Date(),
      sanitizeForSheet(result.correctedBrand),
      sanitizeForSheet(upperModelName),
      '',
      sanitizeForSheet(data.type),
      sanitizeForSheet(upperSize),
      sanitizeForSheet(data.price),
      sanitizeForSheet(data.title),
      '',
      sanitizeForSheet(data.url)
    ]]);

    if (rawYear) {
      sheet.getRange(nextRow, 4).setFormula('="' + rawYear + '"');
    }

    return ContentService.createTextOutput(JSON.stringify({result: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({result: "ERROR", message: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}