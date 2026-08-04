// 새로운 CustomerDB 컬럼 순서 매핑 (1번부터 34번까지)
const COL = {
  DATE: 1,             // 위탁 접수일
  NAME: 2,             // 고객 이름
  PHONE: 3,            // 전화번호
  TX_COUNT: 4,         // 거래횟수
  ID: 5,               // 고유번호
  PW: 6,               // 조회 비밀번호
  ITEM: 7,             // 위탁상품
  STATUS: 8,           // 현재상태
  COMMENT: 9,          // 진행 코멘트
  SERVICE: 10,         // 서비스 등급 (베이직 판매 / 릴케어 인증 판매)
  LINK: 11,            // 판매링크
  ACCOUNT: 12,         // 정산계좌
  DELIVERY_FEE: 13,    // 위탁배송비
  FAST_PRICE: 14,      // 빠른 매도가
  OPT_PRICE: 15,       // 적정 매도가
  HIGH_PRICE: 16,      // 높은 매도가
  REPAIR_COST: 17,     // 정비비
  PART_COST: 18,       // 부품비
  FINAL_PRICE: 19,     // 최종 매도가
  COMMISSION: 20,      // 위탁수수료 (17%)
  SETTLEMENT: 21,      // 최종정산금
  ADDRESS: 22,         // 수거주소
  PICKUP_DATE: 23,     // 희망픽업일
  NAVER_FEE: 24,       // 네이버 수수료 (5%)
  SHIPPING_FEE: 25,    // 배송비
  ACTUAL_PROFIT: 26,   // 실제수익
  USER_INQUIRY: 27,    // 위탁자 문의
  SELLER_ANSWER: 28,   // 판매자 답변
  INSPECTION_ID: 29,   // 검수번호 (검수DB 연결키)
  INSPECTION_RES: 30,  // 검수결과
  INSPECTION_DATE: 31, // 검수완료일
  SALE_REG_DATE: 32,   // 판매등록일
  SALE_END_DATE: 33,   // 판매완료일
  ADMIN_MEMO: 34       // 관리자 메모
};

const IDX = {};
Object.keys(COL).forEach(function(key) {
  IDX[key] = COL[key] - 1;
});

function sanitizeForSheet(val) {
  var s = (val === null || val === undefined) ? "" : String(val);
  if (/^[=+\-@]/.test(s)) { s = "'" + s; }
  return s;
}

function forceTextValue(val) {
  var s = (val === null || val === undefined) ? "" : String(val);
  return "'" + s;
}

function normalizePhone(phone) {
  var p = String(phone || "").replace(/[^0-9]/g, '');
  if (p.length === 10 && p.startsWith("10")) { p = "0" + p; }
  else if (p.length === 8) { p = "010" + p; }
  return p;
}

// 새로운 시트 이름 적용 ('베이트시세표', '스피닝시세표')
function getSheetName(type) { 
  return (type === 'Spin') ? '스피닝시세표' : '베이트시세표'; 
}

function getManufacturers(type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName(type));
  if (!sheet) return [];
  var data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
  var list = [];
  data.forEach(function(r) { if (r[0] && list.indexOf(r[0]) === -1) list.push(r[0]); });
  return list;
}

// 모델 목록 및 바디컨셉(D열), 간단소개(H열) 정보를 함께 구조화하여 전달
function getModelsAndDetails(type, manu) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName(type));
  if (!sheet) return { list: [], map: {} };
  
  // 시트 구조: A:제조사, B:모델명, C:연식, D:바디컨셉/옵션, E:정가, F:중고최소가, G:중고최대가, H:간단소개
  var data = sheet.getRange("A2:H" + sheet.getLastRow()).getValues();
  var modelList = [], modelMap = {};
  
  data.forEach(function(row) {
    if (row[0] === manu && row[1]) {
      var model = String(row[1]).trim();
      var rawYear = String(row[2]).trim();
      var year = (rawYear === "") ? "연식없음" : rawYear;
      var concept = String(row[3] || "").trim(); // D열 바디컨셉/옵션
      var intro = String(row[7] || "").trim();   // H열 간단소개
      
      if (modelList.indexOf(model) === -1) modelList.push(model);
      if (!modelMap[model]) modelMap[model] = [];
      
      modelMap[model].push({
        year: year,
        concept: concept,
        intro: intro
      });
    }
  });
  return { list: modelList, map: modelMap };
}

// 검색 시 최소가 ~ 최대가 반환 함수
function getPriceRange(type, manu, model, year) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName(type));
  if (!sheet) return null;
  
  var data = sheet.getRange("A2:H" + sheet.getLastRow()).getValues();

  for (var i = 0; i < data.length; i++) {
    var sheetManu = String(data[i][0]).trim();
    var sheetModel = String(data[i][1]).trim();
    var sheetRowYear = String(data[i][2]).trim();
    var normalizedSheetYear = (sheetRowYear === "") ? "연식없음" : sheetRowYear;
    var normalizedTargetYear = (String(year).trim() === "") ? "연식없음" : String(year).trim();
    
    if (sheetManu === manu && sheetModel === model && normalizedSheetYear === normalizedTargetYear) {
      var parseVal = function(val) {
        if (typeof val === 'string') return parseInt(val.replace(/[^0-9]/g, ''), 10) || 0;
        return val || 0;
      };
      
      return {
        min: parseVal(data[i][5]), // F열: 중고 최소가
        max: parseVal(data[i][6])  // G열: 중고 최대가
      };
    }
  }
  return null;
}

// 1. 위탁 신청 데이터 저장 (초기 상태: 픽업/입고대기)
function saveCustomer(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerDB');
  if (!sheet) return "ERROR_NO_SHEET";

  var rawPhone = normalizePhone(data.p);
  if (!data.n || !rawPhone || rawPhone.length < 8 || !data.pw || String(data.pw).length !== 4) {
    return "ERROR_INVALID_INPUT";
  }

  var now = new Date();
  var yy = String(now.getFullYear()).substring(2, 4);
  var mm = String(now.getMonth() + 1).padStart(2, '0');
  var dd = String(now.getDate()).padStart(2, '0');
  var formattedDate = yy + mm + dd;

  var phoneData = sheet.getRange("C2:C" + sheet.getLastRow()).getValues();
  var txCount = 1;
  for (var i = 0; i < phoneData.length; i++) {
    if (normalizePhone(phoneData[i][0]) === rawPhone) { txCount++; }
  }

  var idCount = 1;
  var idData = sheet.getRange("E2:E" + sheet.getLastRow()).getValues();
  idData.forEach(function(row){
    if(String(row[0]).startsWith("RC" + formattedDate)){ idCount++; }
  });

  var uniqueId = "RC" + formattedDate + "-" + String(idCount).padStart(3,'0');
  var initialStatus = "픽업/입고대기";
  var initialComment = getDefaultCommentForStatus(initialStatus);

  var newRow = [];
  newRow[IDX.DATE] = formattedDate;
  newRow[IDX.NAME] = sanitizeForSheet(data.n);
  newRow[IDX.PHONE] = forceTextValue(rawPhone);
  newRow[IDX.TX_COUNT] = txCount;
  newRow[IDX.ID] = uniqueId;
  newRow[IDX.PW] = forceTextValue(data.pw);
  newRow[IDX.ITEM] = sanitizeForSheet(data.item);
  newRow[IDX.STATUS] = initialStatus;
  newRow[IDX.COMMENT] = initialComment;
  newRow[IDX.SERVICE] = "";
  newRow[IDX.LINK] = "";
  newRow[IDX.ACCOUNT] = "";
  newRow[IDX.DELIVERY_FEE] = 4000;
  newRow[IDX.FAST_PRICE] = "";
  newRow[IDX.OPT_PRICE] = "";
  newRow[IDX.HIGH_PRICE] = "";
  newRow[IDX.REPAIR_COST] = "";
  newRow[IDX.PART_COST] = "";
  newRow[IDX.FINAL_PRICE] = "";
  newRow[IDX.COMMISSION] = "";
  newRow[IDX.SETTLEMENT] = "";
  newRow[IDX.ADDRESS] = sanitizeForSheet(data.addr);
  newRow[IDX.PICKUP_DATE] = forceTextValue(data.pickupDate);
  newRow[IDX.NAVER_FEE] = "";
  newRow[IDX.SHIPPING_FEE] = "";
  newRow[IDX.ACTUAL_PROFIT] = "";
  newRow[IDX.USER_INQUIRY] = "";
  newRow[IDX.SELLER_ANSWER] = "";
  newRow[IDX.INSPECTION_ID] = "";
  newRow[IDX.INSPECTION_RES] = "";
  newRow[IDX.INSPECTION_DATE] = "";
  newRow[IDX.SALE_REG_DATE] = "";
  newRow[IDX.SALE_END_DATE] = "";
  newRow[IDX.ADMIN_MEMO] = "";

  sheet.appendRow(newRow);

  var newRowIndex = sheet.getLastRow();
  var statusCell = sheet.getRange(newRowIndex, COL.STATUS);
  var statuses = [
    "픽업/입고대기", "입고/검수대기", "전략협의", "베이직 판매", 
    "릴케어 인증 판매", "판매중", "판매정지", "계좌정보제공", 
    "송금대기", "거래완료", "진행보류"
  ];
  var rule = SpreadsheetApp.newDataValidation().explicitValueInList(statuses).build();
  statusCell.setDataValidation(rule);

  return "OK";
}

// [수정된 서버 코드] 검수DB 파싱 함수
function getInspectionData(targetInspectionId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('검수DB');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues(); // 전체 표 데이터 가져오기 (행 x 열)
  var targetIdStr = String(targetInspectionId || "").trim();
  
  // 1단계: 일치하는 검수번호가 들어있는 열(Column) 번호 찾기
  var targetColIdx = -1;
  for (var c = 3; c < data[0].length; c++) { // D열(인덱스 3)부터 검수 데이터 시작
    var cellVal = String(data[0][c] || "").trim(); // 1행이 검수번호임
    if (cellVal === targetIdStr) {
      targetColIdx = c;
      break;
    }
  }

  if (targetColIdx === -1) {
    for (var r = 0; r < data.length; r++) {
      for (var c = 3; c < data[r].length; c++) {
        if (String(data[r][c] || "").trim() === targetIdStr) {
          targetColIdx = c;
          break;
        }
      }
      if (targetColIdx !== -1) break;
    }
  }

  if (targetColIdx === -1) return null;

  var inspection = {
    meta: {},
    items: {}
  };

  // 2단계: 해당 열의 데이터를 읽어오면서 A열(항목명)을 키로 하여 items에 저장
  for (var r = 0; r < data.length; r++) {
    var itemKey = String(data[r][0] || "").trim(); // A열: 항목명 (예: 검수자, 제조사, 외관 평가 등)
    var maxScore = data[r][1];                     // B열: 배점
    var stepsText = data[r][2];                    // C열: 단계 텍스트
    var actualVal = data[r][targetColIdx];         // 해당 상품의 실데이터 값

    if (!itemKey) continue;

    // 상단 메타 정보 영역도 items에 같이 담아주어야 프론트엔드에서 items["검수자"]로 읽을 수 있습니다.
    inspection.items[itemKey] = {
      maxScore: maxScore,
      steps: stepsText,
      value: actualVal
    };
  }

  return inspection;
}

// 3. 위탁품 상세 조회
function checkProduct(name, phone, pw) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerDB');
  if (!sheet) return "NO_DATA";
  
  var data = sheet.getRange("B2:AJ" + sheet.getLastRow()).getValues();
  var results = [];
  var isNamePhoneMatch = false;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (String(row[IDX.NAME - 1]).trim() === String(name).trim() && normalizePhone(row[IDX.PHONE - 1]) === normalizePhone(phone)) {
      isNamePhoneMatch = true;
      if (String(row[IDX.PW - 1]).trim() === String(pw).trim()) {
        
        var inspectionId = String(row[IDX.INSPECTION_ID - 1] || "").trim();
        var inspectionDetails = null;
        if (inspectionId) {
          inspectionDetails = getInspectionData(inspectionId);
        }

        results.push({
          id: row[IDX.ID - 1],
          item: row[IDX.ITEM - 1],
          status: row[IDX.STATUS - 1],
          name: row[IDX.NAME - 1] || "",
          comment: row[IDX.COMMENT - 1] || "",
          service: row[IDX.SERVICE - 1] || "",
          link: row[IDX.LINK - 1] || "",
          settlementAccount: row[IDX.ACCOUNT - 1] || "",
          fastPrice: row[IDX.FAST_PRICE - 1] || 0,
          optPrice: row[IDX.OPT_PRICE - 1] || 0,
          highPrice: row[IDX.HIGH_PRICE - 1] || 0,
          repairCost: row[IDX.REPAIR_COST - 1] || 0,
          partCost: row[IDX.PART_COST - 1] || 0,
          finalPrice: row[IDX.FINAL_PRICE - 1] || 0,
          commission: row[IDX.COMMISSION - 1] || 0,
          settlement: row[IDX.SETTLEMENT - 1] || 0,
          userInquiry: row[IDX.USER_INQUIRY - 1] || "",
          sellerAnswer: row[IDX.SELLER_ANSWER - 1] || "",
          inspectionId: inspectionId,
          inspectionDetails: inspectionDetails
        });
      }
    }
  }

  if (!isNamePhoneMatch) return "NO_DATA";
  if (results.length === 0) return "PW_ERROR";
  return results;
}

// 4. 위탁자의 전략 선택 저장
function saveStrategyChoice(name, phone, pw, id, strategyType, chosenPrice) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerDB');
  if (!sheet) return "ERROR_NO_SHEET";

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowMatch =
      String(row[IDX.NAME]).trim() === String(name).trim() &&
      normalizePhone(row[IDX.PHONE]) === normalizePhone(phone) &&
      String(row[IDX.PW]).trim() === String(pw).trim() &&
      String(row[IDX.ID]).trim() === String(id).trim();

    if (rowMatch) {
      var rowNum = i + 1;
      
      var serviceName = (strategyType === '릴케어 인증 판매' || strategyType === '프리미엄') ? '릴케어 인증 판매' : '베이직 판매';
      var targetStatus = serviceName;
      
      sheet.getRange(rowNum, COL.SERVICE).setValue(serviceName);
      
      var finalPrice = (serviceName === '릴케어 인증 판매') ? (Number(chosenPrice) + 20000) : Number(chosenPrice);
      sheet.getRange(rowNum, COL.FINAL_PRICE).setValue(finalPrice);
      
      sheet.getRange(rowNum, COL.STATUS).setValue(targetStatus);
      sheet.getRange(rowNum, COL.COMMENT).setValue(getDefaultCommentForStatus(targetStatus));

      calculateFinancials(sheet, rowNum, finalPrice, serviceName);
      return "OK";
    }
  }
  return "ERROR_NOT_FOUND";
}

// 5. 정산 계좌 정보 저장
function saveAccountInfo(name, phone, pw, id, bank, account) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerDB');
  if (!sheet) return "ERROR_NO_SHEET";

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowMatch =
      String(row[IDX.NAME]).trim() === String(name).trim() &&
      normalizePhone(row[IDX.PHONE]) === normalizePhone(phone) &&
      String(row[IDX.PW]).trim() === String(pw).trim() &&
      String(row[IDX.ID]).trim() === String(id).trim();

    if (rowMatch) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, COL.ACCOUNT).setValue(sanitizeForSheet(bank + " " + account));
      sheet.getRange(rowNum, COL.STATUS).setValue("송금대기");
      sheet.getRange(rowNum, COL.COMMENT).setValue(getDefaultCommentForStatus("송금대기"));
      return "OK";
    }
  }
  return "ERROR_NOT_FOUND";
}

// 6. 1:1 한 줄 문의 저장
function saveUserInquiry(name, phone, pw, id, inquiryText){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerDB');
  if(!sheet) return "ERROR_NO_SHEET";

  var data = sheet.getDataRange().getValues();
  for(var i = 1; i < data.length; i++){
    var row = data[i];
    var rowMatch =
      String(row[IDX.NAME]).trim() === String(name).trim() &&
      normalizePhone(row[IDX.PHONE]) === normalizePhone(phone) &&
      String(row[IDX.PW]).trim() === String(pw).trim() &&
      String(row[IDX.ID]).trim() === String(id).trim();

    if(rowMatch){
      var rowNum = i + 1;
      sheet.getRange(rowNum, COL.USER_INQUIRY).setValue(sanitizeForSheet(inquiryText));
      return "OK";
    }
  }
  return "ERROR_NOT_FOUND";
}

function getDefaultCommentForStatus(statusName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var commentSheet = ss.getSheetByName('DB컬럼설명');
  if (!commentSheet) return "";
  
  var data = commentSheet.getRange("A2:B" + commentSheet.getLastRow()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(statusName).trim()) {
      return data[i][1];
    }
  }
  return "";
}

function calculateFinancials(sheet, rowNum, finalPrice, serviceType) {
  var deliveryFee = Number(sheet.getRange(rowNum, COL.DELIVERY_FEE).getValue()) || 4000;
  var repairCost = Number(sheet.getRange(rowNum, COL.REPAIR_COST).getValue()) || 0;
  var partCost = Number(sheet.getRange(rowNum, COL.PART_COST).getValue()) || 0;

  var isCertified = (serviceType === '릴케어 인증 판매' || serviceType === '프리미엄');
  var pureProductPrice = isCertified ? (finalPrice - 20000) : finalPrice;
  var commission = pureProductPrice <= 120000 ? 20000 : Math.round(pureProductPrice * 0.17);
  var settlement = pureProductPrice - commission - deliveryFee - repairCost - partCost;
  if (settlement < 0) settlement = 0;

  var naverFee = Math.round(finalPrice * 0.05);
  var actualProfit = commission - naverFee + (isCertified ? 20000 : 0);

  sheet.getRange(rowNum, COL.COMMISSION).setValue(commission);
  sheet.getRange(rowNum, COL.SETTLEMENT).setValue(settlement);
  sheet.getRange(rowNum, COL.NAVER_FEE).setValue(naverFee);
  sheet.getRange(rowNum, COL.ACTUAL_PROFIT).setValue(actualProfit);
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'CustomerDB') return;

  var range = e.range;
  var row = range.rowStart;
  var col = range.columnStart;
  if (row < 2) return;

  var now = new Date();
  var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

  if (col === COL.LINK) {
    var linkVal = sheet.getRange(row, COL.LINK).getValue();
    if (linkVal && String(linkVal).trim() !== "") {
      sheet.getRange(row, COL.SALE_REG_DATE).setValue(formattedDate);
    }
    return;
  }

  if (col === COL.STATUS) {
    var statusCell = sheet.getRange(row, COL.STATUS);
    var newStatus = statusCell.getValue();

    if (newStatus === "전략협의") {
      var inspectionId = sheet.getRange(row, COL.INSPECTION_ID).getValue();
      var fastPrice = Number(sheet.getRange(row, COL.FAST_PRICE).getValue()) || 0;
      var optPrice = Number(sheet.getRange(row, COL.OPT_PRICE).getValue()) || 0;
      var highPrice = Number(sheet.getRange(row, COL.HIGH_PRICE).getValue()) || 0;

      if (!inspectionId || fastPrice === 0 || optPrice === 0 || highPrice === 0) {
        SpreadsheetApp.getUi().alert(
          "⚠️ [입력 누락 경고]\n\n" +
          "검수번호가 입력되지 않았거나, 빠른/적정/높은 매도가 중 빈 항목이 있습니다.\n" +
          "정보를 모두 입력한 후 '전략협의' 상태로 변경해 주세요."
        );
        statusCell.setValue("입고/검수대기");
        sheet.getRange(row, COL.COMMENT).setValue(getDefaultCommentForStatus("입고/검수대기"));
        return;
      }
    }

    if (newStatus === "판매중") {
      var linkVal = sheet.getRange(row, COL.LINK).getValue();
      if (!linkVal || String(linkVal).trim() === "") {
        SpreadsheetApp.getUi().alert(
          "⚠️ [판매 링크 누락 경고]\n\n" +
          "판매 링크(K열)가 입력되지 않았습니다. 링크를 입력한 후 '판매중'으로 변경해 주세요."
        );
        statusCell.setValue("베이직 판매");
        return;
      }
      sheet.getRange(row, COL.SALE_REG_DATE).setValue(formattedDate);
    }

    if (newStatus === "거래완료") {
      var finalPrice = Number(sheet.getRange(row, COL.FINAL_PRICE).getValue()) || 0;
      if (finalPrice > 0) {
        var serviceType = sheet.getRange(row, COL.SERVICE).getValue();
        calculateFinancials(sheet, row, finalPrice, serviceType);
      }
      sheet.getRange(row, COL.SALE_END_DATE).setValue(formattedDate);
    }

    var comment = getDefaultCommentForStatus(newStatus);
    sheet.getRange(row, COL.COMMENT).setValue(comment);
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('릴케어팜 시세조회 및 위탁 서비스')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}