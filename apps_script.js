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

// 연락처1/우편번호/송하인번호/주문번호: 숫자로만 이루어진 문자열을 그냥 appendRow하면
// 시트가 "숫자"로 자동 인식해서 앞자리 0을 없애버림(01099998888 -> 1099998888).
// 열 서식을 미리 텍스트로 걸어두는 방식은 시도해봤지만 실제로 효과가 없었음(재배포 후에도
// 재현됨) - 대신 값 앞에 작은따옴표(')를 붙여서 쓰면 시트가 그 값을 문자로 강제 인식하고
// 작은따옴표 자체는 저장되지 않음. Apps Script로 값을 쓸 때 텍스트를 강제하는 표준적인 방법.
//
// 손님이 입력하는 모든 텍스트 칸(이름/주소/메모 등)에도 똑같이 적용함 - 이유는 하나 더 있는데,
// 값이 "="나 "+"로 시작하면 나중에 사장님이 시트에서 그 셀을 열 때 수식으로 실행될 수 있음
// (스프레드시트 수식 인젝션). 작은따옴표를 앞에 붙이면 이 문제도 같이 막아줌.
function forceText_(v) {
  v = (v === null || v === undefined) ? "" : String(v);
  return v === "" ? "" : ("'" + v);
}

function doPost(e) {
  // 1분에 30건 넘게 들어오면 스크립트로 도배하는 걸로 보고 막음. no-cors라 손님 화면엔 어차피
  // 응답 성공 여부가 안 보이고(3초 지나면 그냥 완료 화면으로 넘어가는 구조라) 여기서 조용히
  // 막아도 됨 - 진짜 손님이 몰려도 1분에 30건이면 거의 걸릴 일 없는 넉넉한 기준.
  if (!checkRateLimit_("submit_rl", 30)) {
    return ContentService.createTextOutput(JSON.stringify({ result: "rate_limited" })).setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("주문");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("주문");
    sheet.appendRow([
      "구매자명", "수취인명", "옵션정보", "수량", "연락처1", "연락처2",
      "배송주소", "상세주소", "우편번호", "송하인번호", "배송메모", "발송예정일",
      "접수시각", "합계금액", "광고수신동의", "동의시각", "입금확인", "주문번호"
    ]);
  }
  var data = JSON.parse(e.postData.contents);

  // 손님이 "최근 1시간 내 같은 정보로 주문한 기록이 있어요" 안내를 보고 "기존 주문 수정"을
  // 선택했거나, 주문조회에서 "수정하기"로 들어온 경우 - 새 행을 또 만들지 않고 기존 행을 고침.
  if (data.updateOrderNumber) {
    return updateExistingOrder_(sheet, data);
  }

  var sameParty = data.sameParty !== false;

  // 구매자명/송하인번호: "다른 분이 받으신다"고 체크했으면 보내는사람(주문자) 정보를 씀.
  // 같은 분이면 수취인 정보를 그대로 구매자로도 씀.
  var buyerName = sameParty ? (data.name || "") : (data.buyerName || "");
  var senderPhone = sameParty ? (data.phone || "") : (data.buyerPhone || "");

  sheet.appendRow([
    forceText_(buyerName),        // 구매자명
    forceText_(data.name),        // 수취인명
    forceText_(data.productText), // 옵션정보
    1,                         // 수량 (박스 개수 - 상품 개수는 옵션정보 텍스트 안에 이미 포함됨)
    forceText_(data.phone),    // 연락처1 (수취인 연락처) - 앞자리 0 보존
    "",                        // 연락처2 (이 폼에서는 안 받음)
    forceText_(data.address),       // 배송주소 (다음 우편번호 서비스로 검색된 정확한 주소)
    forceText_(data.addressDetail), // 상세주소 (동/호수 등, 고객이 직접 입력)
    forceText_(data.zipcode),  // 우편번호 (다음 우편번호 서비스에서 자동으로 받아옴) - 앞자리 0 보존
    forceText_(senderPhone),   // 송하인번호 (보내는사람 자기 번호) - 앞자리 0 보존
    forceText_(data.note),     // 배송메모
    data.shipDate || "",       // 발송예정일 (고객이 직접 지정 안 했으면 빈칸 - 나중에 직접 정함)
    new Date(),                 // 접수시각
    data.totalAmount || 0,      // 합계금액
    data.marketingConsent ? "동의" : "미동의",
    data.marketingConsentAt || "",
    false,                      // 입금확인 (새 주문은 항상 미확인으로 시작, 사장님이 나중에 체크)
    forceText_(data.orderNumber), // 주문번호 (손님 화면에 뜨는 6자리 번호, 나중에 주문조회용) - 앞자리 0 보존
  ]);

  formatNewOrderRow(sheet, sheet.getLastRow()); // 방금 추가된 행 1개만 서식 (전체 재적용은 느려서 안 함)

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
  if (p.action === "checkRecent") {
    return checkRecentOrder_(p);
  }
  if (p.action === "editLookup") {
    return getOrderForEdit_(p);
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
    // 실제로 계좌에서 돈을 보내는 사람은 항상 "구매자명"(A열)임. "다른 분이 받으신다"를
    // 체크한 주문은 수취인(B열)과 구매자가 다른 사람이라, 입금자명 대조는 구매자명으로
    // 해야 함(같은 분이면 doPost에서 구매자명=수취인명으로 이미 같게 저장되니 문제 없음).
    var rowBuyer = String(row[0] || ""); // A열: 구매자명
    if (rowConfirmed === true) continue;
    if (Number(rowAmount) !== amount) continue;
    matches.push({ sheetRow: i + 2, buyer: rowBuyer, productText: String(row[2] || "") });
  }

  if (name && matches.length > 1) {
    var nameMatches = matches.filter(function (m) {
      return m.buyer.indexOf(name) !== -1 || name.indexOf(m.buyer) !== -1;
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
    result.message = "입금확인 처리됨: " + matches[0].buyer + " / " + amount + "원";
    sendAlimtalkNotification_(matches[0].buyer, matches[0].productText, amount);
  }
  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 1분에 몇 번까지 허용할지를 CacheService로 셈. Apps Script 웹앱은 요청자 IP를 알 방법이
// 없어서 "누가" 많이 하는지는 구분 못 하고 "전체 합쳐서" 얼마나 왔는지만 셀 수 있음 - 그래도
// 주문번호(6자리, 백만 가지) 전수조사처럼 짧은 시간에 수천 번씩 두드리는 건 확실히 막아줌.
// 실제 손님 트래픽(한 명이 한 번 조회/주문)은 이 정도 기준에 절대 안 걸림.
function checkRateLimit_(key, maxPerMinute) {
  var cache = CacheService.getScriptCache();
  var count = Number(cache.get(key) || 0);
  if (count >= maxPerMinute) return false;
  cache.put(key, String(count + 1), 60); // 60초 지나면 자동으로 초기화됨
  return true;
}

// 손님이 "주문조회" 화면에서 주문번호(6자리)만 입력하면 그 주문 하나만 찾아서 알려줌.
// 시트 전체를 브라우저로 내려보내면 다른 손님 정보까지 다 노출되니까, 여기서 딱 그 한 건만
// 골라서 돌려줌. 주소도 전체 다 보여주지 않고 앞부분(시/군/구 정도)만 마스킹해서 줌.
function lookupOrderByNumber(p) {
  var result = { found: false };

  // 1분에 60번 이상 조회되면 주문번호 전수조사(브루트포스) 시도로 보고 잠깐 막음.
  if (!checkRateLimit_("lookup_rl", 60)) {
    result.message = "너무 많은 조회가 있었어요. 잠시 후 다시 시도해주세요.";
    return jsonOut(result);
  }

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

// 손님이 실수로 잘못 주문했을 때, 새 행을 또 만들지 않고 기존 주문 행을 고쳐씀.
// 주문번호만으로는 아무나 남의 주문을 고칠 수 있으니, 그 주문에 저장된 연락처와
// 손님이 지금 입력한 연락처가 같은지 한 번 더 확인하고 나서만 수정함.
function updateExistingOrder_(sheet, data) {
  var orderNumber = String(data.updateOrderNumber || "").trim();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonOut({ result: "error", message: "주문 데이터 없음" });
  }
  var values = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
  var phone = String(data.phone || "").trim();

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][17]) !== orderNumber) continue;
    var storedPhone = String(values[i][4] || "").trim();
    if (storedPhone !== phone) {
      return jsonOut({ result: "error", message: "연락처가 일치하지 않아 수정할 수 없음" });
    }
    var row = i + 2;
    sheet.getRange(row, 1).setValue(forceText_(data.name));           // 구매자명 (수정은 항상 본인 기준으로 단순화)
    sheet.getRange(row, 2).setValue(forceText_(data.name));           // 수취인명
    sheet.getRange(row, 3).setValue(forceText_(data.productText));    // 옵션정보
    sheet.getRange(row, 5).setValue(forceText_(data.phone));          // 연락처1
    sheet.getRange(row, 7).setValue(forceText_(data.address));        // 배송주소
    sheet.getRange(row, 8).setValue(forceText_(data.addressDetail));  // 상세주소
    sheet.getRange(row, 9).setValue(forceText_(data.zipcode));        // 우편번호
    sheet.getRange(row, 10).setValue(forceText_(data.phone));         // 송하인번호 (수정은 항상 본인 기준)
    sheet.getRange(row, 11).setValue(forceText_(data.note));          // 배송메모
    sheet.getRange(row, 12).setValue(data.shipDate || "");            // 발송예정일
    sheet.getRange(row, 14).setValue(data.totalAmount || 0);          // 합계금액
    formatNewOrderRow(sheet, row);
    return jsonOut({ result: "success", updated: true });
  }
  return jsonOut({ result: "error", message: "해당 주문번호를 찾을 수 없음" });
}

// 같은 사람이 실수로 또 주문했을 때를 대비한 중복 감지. 이름+연락처+주소가 전부 같고
// 접수시각이 1시간 이내인 주문이 있으면 알려줌 - 손님이 "그 주문 수정할지" 고를 수 있게.
function checkRecentOrder_(p) {
  var result = { found: false };
  if (!checkRateLimit_("checkrecent_rl", 60)) {
    return jsonOut(result); // 너무 잦으면 그냥 "없음"으로 취급 (새 주문 진행에는 지장 없음)
  }

  var name = (p.name || "").trim();
  var phone = (p.phone || "").trim();
  var address = (p.address || "").trim();
  if (!name || !phone) return jsonOut(result);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("주문");
  if (!sheet) return jsonOut(result);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOut(result);

  var data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
  var oneHourAgo = new Date(Date.now() - 60 * 60 * 1000);
  for (var i = data.length - 1; i >= 0; i--) { // 최근 행부터 확인
    var row = data[i];
    var rowTime = row[12];
    if (!(rowTime instanceof Date) || rowTime < oneHourAgo) continue;
    if (String(row[1] || "").trim() !== name) continue;
    if (String(row[4] || "").trim() !== phone) continue;
    if (String(row[6] || "").trim() !== address) continue;
    result.found = true;
    result.orderNumber = String(row[17] || "");
    result.productText = row[2];
    result.receivedAt = formatDateVal_(row[12]);
    return jsonOut(result);
  }
  return jsonOut(result);
}

// 손님이 주문조회에서 "수정하기"를 누르면, 주문번호+연락처가 같이 맞아야 전체 정보를
// 내려줌 (조회 화면(lookupOrder)은 주소를 마스킹해서 보여주지만, 수정하려면 원본 폼을
// 채워야 해서 원본이 필요함 - 그래서 조회보다 한 단계 더 강한 인증을 요구함).
function getOrderForEdit_(p) {
  var result = { found: false };
  if (!checkRateLimit_("editlookup_rl", 30)) {
    result.message = "너무 많은 요청이 있었어요. 잠시 후 다시 시도해주세요.";
    return jsonOut(result);
  }

  var orderNumber = String(p.orderNumber || "").trim();
  var phone = String(p.phone || "").trim();
  if (!/^\d{6}$/.test(orderNumber) || !phone) {
    result.message = "주문번호와 연락처를 정확히 입력해주세요";
    return jsonOut(result);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("주문");
  if (!sheet) { result.message = "주문 시트 없음"; return jsonOut(result); }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { result.message = "해당 번호의 주문을 찾을 수 없어요"; return jsonOut(result); }

  var data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (String(row[17]) !== orderNumber) continue;
    if (String(row[4] || "").trim() !== phone) {
      result.message = "주문번호와 연락처가 일치하지 않아요";
      return jsonOut(result);
    }
    result.found = true;
    result.name = row[1];
    result.phone = row[4];
    result.address = row[6];
    result.addressDetail = row[7];
    result.zipcode = row[8];
    result.note = row[10];
    result.shipDate = row[11] || "";
    return jsonOut(result);
  }
  result.message = "해당 번호의 주문을 찾을 수 없어요";
  return jsonOut(result);
}

// 시트에서 "입금확인" 체크박스를 사장님이 직접 손으로 클릭했을 때도 알림톡이 가게 하는
// 트리거. Apps Script의 onEdit(e)는 특별한 이름이라 시트를 편집할 때마다 자동 실행됨 -
// 별도로 트리거 등록 안 해도 됨(단순 트리거). 다만 단순 트리거는 UrlFetchApp 같은 외부
// 호출 권한이 제한적일 수 있어서, 안 되면 편집기에서 "트리거" 메뉴로 설치형 트리거를
// 이 함수 이름으로 하나 등록해줘야 함(스프레드시트 수정 시 실행).
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "주문") return;
  if (e.range.getColumn() !== 17) return; // Q열: 입금확인
  if (e.value !== "TRUE") return; // 체크(true)로 바뀐 경우만
  var row = e.range.getRow();
  if (row < 2) return;
  var rowData = sheet.getRange(row, 1, 1, 14).getValues()[0];
  sendAlimtalkNotification_(String(rowData[0] || ""), String(rowData[2] || ""), rowData[13]);
}

// SOLAPI를 통해 카카오 알림톡으로 사장님 폰에 "입금확인된 주문" 알림을 보냄.
// 자동확인(confirmPaymentByAmount)이든 사장님이 시트에서 직접 체크박스를 누르든
// 둘 다 이 함수를 거쳐서 알림이 감.
//
// 필요한 스크립트 속성(프로젝트 설정 > 스크립트 속성에서 추가):
// SOLAPI_API_KEY, SOLAPI_API_SECRET, SOLAPI_PFID, SOLAPI_TEMPLATE_ID,
// SOLAPI_SENDER_PHONE(솔라피에 사전등록된 발신번호), SOLAPI_RECEIVER_PHONE(알림 받을 사장님 번호)
//
// 템플릿 변수 이름은 #{이름}/#{상품}/#{금액}으로 등록했다고 가정함 - 실제 등록한 변수명이
// 다르면 아래 variables 객체의 키를 그거에 맞게 고쳐야 함.
function sendAlimtalkNotification_(buyerName, productText, amount) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty("SOLAPI_API_KEY");
  var apiSecret = props.getProperty("SOLAPI_API_SECRET");
  var pfId = props.getProperty("SOLAPI_PFID");
  var templateId = props.getProperty("SOLAPI_TEMPLATE_ID");
  var senderPhone = props.getProperty("SOLAPI_SENDER_PHONE");
  var receiverPhone = props.getProperty("SOLAPI_RECEIVER_PHONE");
  // 설정이 아직 안 끝났으면 조용히 건너뜀 - 알림톡 실패 때문에 입금확인 처리 자체가
  // 실패한 것처럼 보이면 안 되니까(이미 입금확인은 이 함수 호출 전에 끝난 상태).
  if (!apiKey || !apiSecret || !pfId || !templateId || !senderPhone || !receiverPhone) return;

  var date = new Date().toISOString();
  var salt = Utilities.getUuid();
  var rawSignature = Utilities.computeHmacSha256Signature(date + salt, apiSecret);
  var signature = rawSignature.map(function (b) {
    return ("0" + (b & 0xFF).toString(16)).slice(-2);
  }).join("");

  var payload = {
    message: {
      to: receiverPhone,
      from: senderPhone,
      kakaoOptions: {
        pfId: pfId,
        templateId: templateId,
        variables: {
          "#{이름}": buyerName || "",
          "#{상품}": productText || "",
          "#{금액}": String(amount || "")
        }
      }
    }
  };

  try {
    var response = UrlFetchApp.fetch("https://api.solapi.com/messages/v4/send", {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "HMAC-SHA256 apiKey=" + apiKey + ", date=" + date + ", salt=" + salt + ", signature=" + signature
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true, // 실패해도 예외로 안 터지게(입금확인 처리에 영향 없도록)
    });
    // 실패 원인 확인용 - Apps Script 편집기 "실행 기록"(왼쪽 시계 아이콘)에서 응답 내용을
    // 볼 수 있음. 성공해도 어차피 로그는 남기고, 문제 생겼을 때만 여기서 확인하면 됨.
    Logger.log("SOLAPI 알림톡 응답 " + response.getResponseCode() + ": " + response.getContentText());
  } catch (err) {
    // 알림톡 발송 실패는 무시함 - 입금확인 자체는 이미 정상 처리된 뒤라 손님/사장님 업무에 지장 없음
    Logger.log("SOLAPI 알림톡 호출 중 예외: " + err);
  }
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

// 새로 추가된 주문 행 "1개"에만 필요한 서식(줄바꿈/금액·날짜 서식/체크박스)을 입힘.
// 예전엔 주문이 들어올 때마다 formatOrderSheet()로 전체 행 범위(줄무늬 배경, 열너비 18개
// 재설정 포함)를 통째로 다시 서식 입혔는데, 이게 구글시트 API 호출을 데이터 전체 크기만큼
// 반복해서 주문이 쌓일수록 "접수하는 중..."이 점점 느려지는 원인이었음(시트에 값 자체는
// appendRow로 바로 반영되지만, doPost 응답이 이 전체 재서식이 끝나야 돌아가서 화면이 안 넘어감).
// 새 행 1개만 처리하면 주문이 몇 건이 쌓여도 매번 빠름.
function formatNewOrderRow(sheet, row) {
  sheet.getRange(row, 3).setWrap(true); // 옵션정보 - 길어질 수 있어서 줄바꿈
  sheet.getRange(row, 13).setNumberFormat("yyyy-mm-dd hh:mm"); // 접수시각
  sheet.getRange(row, 14).setNumberFormat('#,##0"원"'); // 합계금액
  sheet.getRange(row, 17).insertCheckboxes(); // 입금확인
}

// 시트 전체를 보기 편하게 서식 적용 (헤더 색칠+고정, 열 너비, 줄무늬 배경 등).
// 이제 주문이 들어올 때 자동으로는 안 돌고(느려서 formatNewOrderRow로 대체함),
// 줄무늬 배경이나 열 너비가 틀어져 보일 때 가끔 수동으로만 돌리면 됨:
// Apps Script 편집기 상단에서 함수 선택 드롭다운을 "formatOrderSheet"로 바꾸고 ▶ 실행 버튼 한 번 누르면 됨.
function formatOrderSheet(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("주문");
  if (!sheet) return;

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
