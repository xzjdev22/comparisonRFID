// ==========================================
// 1. 코드 최상단 상수 정의
// ==========================================
const SHEET_REALTIME = "실시간전송내용";
const SHEET_TARGET_ALL = "청구대상건";
const SHEET_TARGET = "청구대상건(요양)";
const SHEET_SERVICE = "서비스내용 입력(요양)";
const SHEET_RESULT = "실전송대비 청구기록 비교(요양)";
const COLOR_POSITIVE = "#d9ead3";
const COLOR_NEGATIVE = "#f4cccc";

/**
 * 실전송대비 청구기록 비교 시트 생성 및 서식/조건부서식 적용 함수
 */
function createComparisonSheet() {
  const startTime = new Date().getTime();
  Logger.log(
    "🚀 [시작] 실전송대비 청구기록 비교 스크립트 시작 (고차함수/초고속 버전)",
  );

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 대상 시트 가져오기
  const targetSheet = ss.getSheetByName(SHEET_TARGET);
  const serviceSheet = ss.getSheetByName(SHEET_SERVICE);

  if (!targetSheet || !serviceSheet) {
    Logger.log("❌ [에러] 필수 시트를 찾을 수 없습니다.");
    SpreadsheetApp.getUi().alert(
      `'${SHEET_TARGET}' 또는 '${SHEET_SERVICE}' 시트를 찾을 수 없습니다.`,
    );
    return;
  }

  Logger.log("1️⃣ [시트 생성] 기존 결과 시트 삭제 및 새 시트 생성 중...");
  let resultSheet = ss.getSheetByName(SHEET_RESULT);
  if (resultSheet) {
    ss.deleteSheet(resultSheet);
  }
  resultSheet = ss.insertSheet(SHEET_RESULT);

  // ----------------------------------------------------
  // 2. 청구대상건 원본 전체 복사 (서식/멘션/링크 보존)
  // ----------------------------------------------------
  const targetRange = targetSheet.getDataRange();
  const lastRow = targetRange.getLastRow();
  const lastCol = targetRange.getLastColumn();

  if (lastRow < 1) {
    Logger.log("⚠️ [경고] 청구대상건 시트에 데이터가 없습니다.");
    return;
  }

  Logger.log(
    `2️⃣ [데이터 복사] 원본 데이터 전체 복사 중... (총 ${lastRow}행 / ${lastCol}열)`,
  );
  targetRange.copyTo(
    resultSheet.getRange(1, 1),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false,
  );

  // ----------------------------------------------------
  // 3. C열(인정번호) 및 E열(핸드폰번호) 제거
  // ----------------------------------------------------
  Logger.log("3️⃣ [열 삭제] C열(인정번호), E열(핸드폰번호) 삭제 중...");
  resultSheet.deleteColumn(5); // 원본 E열 삭제
  resultSheet.deleteColumn(3); // 원본 C열 삭제

  // ----------------------------------------------------
  // 4. 서비스내용 입력 시트 데이터 맵핑 (PK: 수급자명 + YYYY-MM-DD)
  // ----------------------------------------------------
  Logger.log("4️⃣ [서비스내용 입력] PK 데이터 맵핑 생성 중...");
  const serviceValues = serviceSheet.getDataRange().getValues();

  // 헤더(1행) 제외 후 reduce로 Map 구성
  const serviceMap = serviceValues.slice(1).reduce((map, rowData) => {
    const recipientName = String(rowData[15] || "").trim(); // P열 (0-index 15)
    const rawDate = rowData[4]; // E열 (0-index 4)
    const dateStr = formatDateToYYYYMMDD(rawDate);

    if (recipientName && dateStr) {
      const pk = `${recipientName}_${dateStr}`;
      map.set(pk, {
        provideTime: rowData[12], // M열 (0-index 12)
        startTime: rowData[9], // J열 (0-index 9)
        endTime: rowData[11], // L열 (0-index 11)
      });
    }
    return map;
  }, new Map());

  Logger.log(`   └─ 매칭 가능 수급자 PK 수: ${serviceMap.size}개`);

  // ----------------------------------------------------
  // 5. 추가 열 (제공시간, 시작, 종료) 헤더 및 데이터 세팅
  // ----------------------------------------------------
  Logger.log("5️⃣ [데이터 연동] 데이터 map 처리 및 일괄 쓰기 중...");
  const newColStart = 8; // H열 (C,E열 삭제 후 G열 오른쪽에 추가)

  // 헤더 추가
  resultSheet
    .getRange(1, newColStart, 1, 3)
    .setValues([["제공시간", "시작", "종료"]]);

  // 열 너비 설정
  resultSheet.setColumnWidth(newColStart, serviceSheet.getColumnWidth(13));
  resultSheet.setColumnWidth(newColStart + 1, serviceSheet.getColumnWidth(10));
  resultSheet.setColumnWidth(newColStart + 2, serviceSheet.getColumnWidth(12));

  // 헤더 서식 복사
  resultSheet
    .getRange(1, 7)
    .copyTo(
      resultSheet.getRange(1, newColStart, 1, 3),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false,
    );

  const targetValues = targetSheet.getDataRange().getValues();

  // map 함수를 통해 매칭 데이터 배열 생성 (헤더 제외)
  const appendedData = targetValues.slice(1).map((rowData) => {
    const recipientName = String(rowData[1] || "").trim(); // B열 (수급자성명, 0-index 1)
    const rawDate = rowData[7]; // H열 (시작시간, 0-index 7)

    const dateStr = formatDateToYYYYMMDD(rawDate);
    const pk = `${recipientName}_${dateStr}`;

    const match = serviceMap.get(pk);
    return match
      ? [match.provideTime, match.startTime, match.endTime]
      : ["", "", ""];
  });

  // 매칭 결과 데이터를 단 1회의 setValues로 일괄 반영
  if (appendedData.length > 0) {
    resultSheet
      .getRange(2, newColStart, appendedData.length, 3)
      .setValues(appendedData);
  }

  // ----------------------------------------------------
  // 6. 동적 조건부 서식 (Conditional Formatting Rules) 적용
  // ----------------------------------------------------
  Logger.log("6️⃣ [조건부 서식] 동적 색상 조건부 서식 적용 중...");
  const rules = [];

  // 1) 총시간(E열) vs 제공시간(H열)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($E2<>"", $H2<>"", INT($E2)=INT($H2))')
      .setBackground(COLOR_POSITIVE)
      .setRanges([
        resultSheet.getRange(`E2:E${lastRow}`),
        resultSheet.getRange(`H2:H${lastRow}`),
      ])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($E2<>"", $H2<>"", INT($E2)<>INT($H2))')
      .setBackground(COLOR_NEGATIVE)
      .setRanges([
        resultSheet.getRange(`E2:E${lastRow}`),
        resultSheet.getRange(`H2:H${lastRow}`),
      ])
      .build(),
  );

  // 2) 시작시간(F열) vs 시작(I열) [hh:mm 비교]
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($F2<>"", $I2<>"", TEXT($F2,"hh:mm")=TEXT($I2,"hh:mm"))',
      )
      .setBackground(COLOR_POSITIVE)
      .setRanges([
        resultSheet.getRange(`F2:F${lastRow}`),
        resultSheet.getRange(`I2:I${lastRow}`),
      ])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($F2<>"", $I2<>"", TEXT($F2,"hh:mm")<>TEXT($I2,"hh:mm"))',
      )
      .setBackground(COLOR_NEGATIVE)
      .setRanges([
        resultSheet.getRange(`F2:F${lastRow}`),
        resultSheet.getRange(`I2:I${lastRow}`),
      ])
      .build(),
  );

  // 3) 종료시간(G열) vs 종료(J열) [hh:mm 비교]
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($G2<>"", $J2<>"", TEXT($G2,"hh:mm")=TEXT($J2,"hh:mm"))',
      )
      .setBackground(COLOR_POSITIVE)
      .setRanges([
        resultSheet.getRange(`G2:G${lastRow}`),
        resultSheet.getRange(`J2:J${lastRow}`),
      ])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($G2<>"", $J2<>"", TEXT($G2,"hh:mm")<>TEXT($J2,"hh:mm"))',
      )
      .setBackground(COLOR_NEGATIVE)
      .setRanges([
        resultSheet.getRange(`G2:G${lastRow}`),
        resultSheet.getRange(`J2:J${lastRow}`),
      ])
      .build(),
  );

  resultSheet.setConditionalFormatRules(rules);

  // ----------------------------------------------------
  // 7. 시트 순서 정렬
  // ----------------------------------------------------
  Logger.log("7️⃣ [시트 정렬] 지정된 순서대로 시트 배치 중...");
  const orderList = [
    SHEET_REALTIME,
    SHEET_TARGET_ALL,
    SHEET_TARGET,
    SHEET_SERVICE,
    SHEET_RESULT,
  ];

  orderList.forEach((sheetName, index) => {
    const sh = ss.getSheetByName(sheetName);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(index + 1);
    }
  });

  ss.setActiveSheet(resultSheet);

  const endTime = new Date().getTime();
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  Logger.log(`🏁 [완료] 소요 시간: ${duration}초`);

  SpreadsheetApp.getUi().alert(
    `실전송대비 청구기록 비교 시트 생성이 완료되었습니다. (${duration}초 소요)`,
  );
}

/**
 * 날짜 데이터를 YYYY-MM-DD 문자열 형태로 정규화하는 헬퍼 함수
 */
function formatDateToYYYYMMDD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const str = String(val).trim();
  const match = str.match(/\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : "";
}
