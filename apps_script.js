// 시골손맛 주문 접수 - Google Apps Script
// 사용법: 구글시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣고
// "배포 > 새 배포"까지 진행하면 됨 (기존에 배포한 적 있으면 "관리" 톱니에서
// 버전을 "새 버전"으로 올려서 배포해야 코드가 실제로 반영됨).
//
// 컬럼 순서는 파이썬 쪽 image_order.py의 "로젠택배양식" 시트랑 똑같이 맞춰뒀음:
// 구매자명, 수취인명, 옵션정보, 수량, 연락처1, 연락처2, 배송주소, 상세주소,
// 우편번호, 송하인번호, 배송메모, 발송예정일  (+ 접수시각/합계금액/광고동의는 참고용으로 뒤에 추가)
// "입금확인"(맨 마지막 칸)은 손님이 채우는 게 아니라 사장님이 입금 확인 후 직접 체크하는 칸.
// pyorder 쪽에서 이 칸이 체크된 행만 가져가게 되어있음 - 계좌이체라 자동 확인이 안 돼서 그럼.

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("주문");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("주문");
    sheet.appendRow([
      "구매자명", "수취인명", "옵션정보", "수량", "연락처1", "연락처2",
      "배송주소", "상세주소", "우편번호", "송하인번호", "배송메모", "발송예정일",
      "접수시각", "합계금액", "광고수신동의", "동의시각", "입금확인", "주문번호"
    ]);
  }
  ensureTextColumns(sheet); // appendRow 쓰기 전에 미리 텍스트 서식 걸어둬야 010... 앞자리 0이 안 날아감

  var data = JSON.parse(e.postData.contents);
  var sameParty = data.sameParty !== false;

  // 구매자명/송하인번호: "다른 분이 받으신다"고 체크했으면 보내는사람(주문자) 정보를 씀.
  // 같은 분이면 수취인 정보를 그대로 구매자로도 씀.
  var buyerName = sameParty ? (data.name || "") : (data.buyerName || "");
  var senderPhone = sameParty ? (data.phone || "") : (data.buyerPhone || "");

  sheet.appendRow([
    buyerName,                 // 구매자명
    data.name || "",           // 수취인명
    data.productText || "",    // 옵션정보
    1,                         // 수량 (박스 개수 - 상품 개수는 옵션정보 텍스트 안에 이미 포함됨)
    data.phone || "",          // 연락처1 (수취인 연락처)
    "",                        // 연락처2 (이 폼에서는 안 받음)
    data.address || "",        // 배송주소 (다음 우편번호 서비스로 검색된 정확한 주소)
    data.addressDetail || "",  // 상세주소 (동/호수 등, 고객이 직접 입력)
    data.zipcode || "",        // 우편번호 (다음 우편번호 서비스에서 자동으로 받아옴)
    senderPhone,                // 송하인번호 (보내는사람 자기 번호)
    data.note || "",           // 배송메모
    data.shipDate || "",       // 발송예정일 (고객이 직접 지정 안 했으면 빈칸 - 나중에 직접 정함)
    new Date(),                 // 접수시각
    data.totalAmount || 0,      // 합계금액
    data.marketingConsent ? "동의" : "미동의",
    data.marketingConsentAt || "",
    false,                      // 입금확인 (새 주문은 항상 미확인으로 시작, 사장님이 나중에 체크)
    data.orderNumber || "",     // 주문번호 (손님 화면에 뜨는 6자리 번호, 나중에 주문조회용)
  ]);

  formatOrderSheet(sheet); // 주문 들어올 때마다 열 너비를 새 내용에 맞게 다시 조정

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 브라우저에서 이 URL로 그냥 접속했을 때(GET) 확인용 응답 + 입금 자동확인(?action=confirmPayment)
// + 손님 주문조회(?action=lookupOrder)
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "confirmPayment") {
    return confirmPaymentByAmount(p);
  }
  if (p.action === "lookupOrder") {
    return lookupOrderByNumber(p);
  }
  return ContentService
    .createTextOutput("시골손맛 주문 접수 서버가 정상 작동 중입니다.")
    .setMimeType(ContentService.MimeType.TEXT);
}

// 폰 자동화 앱(매크로드로이드 등)이 입금 문자를 감지했을 때 호출하는 창구.
// 금액이 일치하는 "미확인" 주문을 찾아서 자동으로 입금확인 체크함.
// 같은 금액의 미확인 주문이 여러 건이면(구분 불가) 자동 처리 안 하고 사람이 확인하도록 남겨둠.
//
// 비밀키는 코드에 적지 않고 "스크립트 속성"에 따로 저장함(이 파일은 깃허브에 공개되어 있어서,
// 코드에 직접 적으면 누구나 보고 악용할 수 있음). 설정 방법:
// Apps Script 편집기 좌측 톱니바퀴(프로젝트 설정) → 맨 아래 "스크립트 속성" → 속성 추가
// → 속성: CONFIRM_SECRET, 값: 본인만 아는 긴 랜덤 문자열 (예: 32자 이상 영문+숫자)
function confirmPaymentByAmount(p) {
  var secret = PropertiesService.getScriptProperties().getProperty("CONFIRM_SECRET");
  var result = { ok: false, message: "" };

  if (!secret) {
    result.message = "서버에 CONFIRM_SECRET이 설정되지 않음";
    return jsonOut(result);
  }
  if (p.key !== secret) {
    result.message = "인증 실패";
    return jsonOut(result);
  }

  // 폰 쪽에서 "66,000"처럼 콤마가 낀 채로 보내도 되게, 숫자 아닌 문자는 다 제거하고 읽음
  var amount = Number(String(p.amount || "").replace(/[^\d]/g, ""));
  if (!amount) {
    result.message = "금액이 없거나 잘못됨";
    return jsonOut(result);
  }
  var name = (p.name || "").trim();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("주문");
  if (!sheet) {
    result.message = "주문 시트 없음";
    return jsonOut(result);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    result.message = "주문 데이터 없음";
    return jsonOut(result);
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
  var matches = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowAmount = row[13];   // N열: 합계금액
    var rowConfirmed = row[16]; // Q열: 입금확인
    var rowReceiver = String(row[1] || ""); // B열: 수취인명
    if (rowConfirmed === true) continue;
    if (Number(rowAmount) !== amount) continue;
    matches.push({ sheetRow: i + 2, receiver: rowReceiver });
  }

  if (name && matches.length > 1) {
    var nameMatches = matches.filter(function (m) {
      return m.receiver.indexOf(name) !== -1 || name.indexOf(m.receiver) !== -1;
    });
    if (nameMatches.length >= 1) matches = nameMatches;
  }

  if (matches.length === 0) {
    result.message = "금액 " + amount + "원과 일치하는 미확인 주문 없음";
  } else if (matches.length > 1) {
    result.message = "같은 금액(" + amount + "원) 미확인 주문이 " + matches.length + "건이라 자동확인 보류 - 직접 확인 필요";
  } else {
    sheet.getRange(matches[0].sheetRow, 17).setValue(true);
    result.ok = true;
    result.message = "입금확인 처리됨: " + matches[0].receiver + " / " + amount + "원";
  }
  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 손님이 "주문조회" 화면에서 주문번호(6자리)만 입력하면 그 주문 하나만 찾아서 알려줌.
// 시트 전체를 브라우저로 내려보내면 다른 손님 정보까지 다 노출되니까, 여기서 딱 그 한 건만
// 골라서 돌려줌. 주소도 전체 다 보여주지 않고 앞부분(시/군/구 정도)만 마스킹해서 줌.
function lookupOrderByNumber(p) {
  var result = { found: false };
  var orderNumber = String(p.orderNumber || "").trim();
  if (!/^\d{6}$/.test(orderNumber)) {
    result.message = "주문번호는 숫자 6자리예요";
    return jsonOut(result);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("주문");
  if (!sheet) {
    result.message = "주문 시트 없음";
    return jsonOut(result);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    result.message = "해당 번호의 주문을 찾을 수 없어요";
    return jsonOut(result);
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (String(row[17]) === orderNumber) { // R열: 주문번호
      result.found = true;
      result.receiver = row[1];                          // 수취인명
      result.product = row[2];                            // 옵션정보
      result.amount = row[13];                            // 합계금액
      result.shipDate = row[11] || "";                    // 발송예정일(고객 지정, 문자열)
      result.receivedAt = formatDateVal_(row[12]);         // 접수시각
      result.paid = row[16] === true;                      // 입금확인
      result.addressMasked = maskAddress_(row[6]);          // 배송주소(일부만)
      return jsonOut(result);
    }
  }
  result.message = "해당 번호의 주문을 찾을 수 없어요";
  return jsonOut(result);
}

function maskAddress_(addr) {
  var parts = String(addr || "").trim().split(/\s+/);
  return parts.slice(0, 2).join(" ") + (parts.length > 2 ? " ..." : "");
}

function formatDateVal_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd HH:mm");
  }
  return String(v || "");
}

// 시트를 보기 편하게 서식 적용 (헤더 색칠+고정, 열 너비, 줄바꿈, 금액/날짜 서식, 줄무늬 배경).
// 주문이 들어올 때마다(doPost) 자동으로 다시 적용됨. 지금 당장 기존 데이터에 적용하려면:
// Apps Script 편집기 상단에서 함수 선택 드롭다운을 "formatOrderSheet"로 바꾸고 ▶ 실행 버튼 한 번 누르면 됨.
// 연락처1(5)/연락처2(6)/우편번호(9)/송하인번호(10)/주문번호(18): 숫자로만 이루어진 문자열이
// 그대로 appendRow되면 시트가 "숫자"로 자동 인식해서 앞자리 0을 없애버림(01012341234 → 1012341234).
// 열 서식을 미리 "일반 텍스트(@)"로 걸어두면 그 뒤로 들어오는 값은 숫자 변환 없이 문자열 그대로 저장됨.
// 미래에 추가될 행까지 미리 걸어둬야 하므로 getMaxRows()로 열 전체에 적용.
function ensureTextColumns(sheet) {
  [5, 6, 9, 10, 18].forEach(function (col) {
    sheet.getRange(1, col, sheet.getMaxRows(), 1).setNumberFormat("@");
  });
}

function formatOrderSheet(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("주문");
  if (!sheet) return;
  ensureTextColumns(sheet);

  var COLS = 18;
  sheet.getRange(1, 1, 1, COLS)
    .setFontWeight("bold")
    .setBackground("#4B5D34")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  // autoResizeColumns는 Apps Script 자체 버그로 안정적으로 안 먹혀서(구글 이슈트래커에
  // 등록된 미해결 버그) 대신 실제 데이터 길이에 맞춘 고정폭을 씀. 정말 자동으로 맞추고
  // 싶으면 시트에서 열 전체 선택 → 우클릭 → "열 크기 조정" → "데이터에 맞추기"(이건 정상 작동함).
  var widths = {
    1: 100, 2: 100, 3: 260, 4: 55, 5: 110, 6: 110, 7: 220, 8: 140,
    9: 80, 10: 110, 11: 180, 12: 100, 13: 140, 14: 100, 15: 90, 16: 160, 17: 90, 18: 90,
  };
  Object.keys(widths).forEach(function (col) {
    sheet.setColumnWidth(Number(col), widths[col]);
  });

  // getMaxRows()는 시트 전체 용량(새로 만들면 보통 1000행)이라 그걸 쓰면 매번
  // 빈 행 수백~수천 개까지 서식/체크박스를 입히느라 느려짐(주문 접수가 멈춘 것처럼 보임).
  // getLastRow()는 실제 데이터가 있는 마지막 행이라 이걸 써야 함.
  var lastRow = Math.max(sheet.getLastRow(), 2);
  // 옵션정보(3): 상품 여러 개 담기면 내용이 매우 길어질 수 있어서 줄바꿈으로 처리
  sheet.getRange(2, 3, lastRow - 1, 1).setWrap(true);
  sheet.getRange(2, 14, lastRow - 1, 1).setNumberFormat('#,##0"원"'); // 합계금액
  sheet.getRange(2, 13, lastRow - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm"); // 접수시각
  sheet.getRange(2, 17, lastRow - 1, 1).insertCheckboxes(); // 입금확인 - 사장님이 입금 확인 후 직접 체크

  sheet.getBandings().forEach(function (b) { b.remove(); });
  sheet.getRange(1, 1, lastRow, COLS)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN, true, false);
}

// "주문" 시트를 완전히 지우고 헤더+서식이 딱 맞는 깨끗한 상태로 새로 만듦.
// 행만 지우면 시트 자체는 남아있어서 헤더가 다시 안 만들어짐 - 그래서 탭 삭제까지 자동으로 처리함.
// 쓰는 법: 편집기 상단 함수 선택 드롭다운에서 "resetOrderSheet" 선택 → ▶ 실행.
// ⚠️ 지금까지 쌓인 주문 데이터가 전부 지워짐 - 테스트 데이터만 있을 때만 실행할 것.
function resetOrderSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName("주문");
  if (old) ss.deleteSheet(old);

  var sheet = ss.insertSheet("주문");
  sheet.appendRow([
    "구매자명", "수취인명", "옵션정보", "수량", "연락처1", "연락처2",
    "배송주소", "상세주소", "우편번호", "송하인번호", "배송메모", "발송예정일",
    "접수시각", "합계금액", "광고수신동의", "동의시각", "입금확인", "주문번호"
  ]);
  formatOrderSheet(sheet);
}
