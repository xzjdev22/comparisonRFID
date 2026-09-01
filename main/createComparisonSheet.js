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
  Logger.log("🚀 [시작] 실전송대비 청구기록 비교 스크립트 시작");

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 대상 시트 가져오기
  const targetSheet = ss.getSheetByName(SHEET_TARGET);
  const serviceSheet = ss.getSheetByName(SHEET_SERVICE);
  const targetAllSheet = ss.getSheetByName(SHEET_TARGET_ALL);

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
  // 2. 청구대상건 원본 전체 복사 (서식/멘션/링크 완벽 보존)
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
  // 3. 서비스내용 입력 시트 데이터 맵핑 (M, J, L열 매칭)
  // ----------------------------------------------------
  Logger.log("3️⃣ [서비스내용 입력] PK 데이터 맵핑 생성 중...");
  const serviceValues = serviceSheet.getDataRange().getValues();

  // 헤더 제외 후 Map 구성 (PK: 수급자성명 + YYYY-MM-DD)
  const serviceMap = serviceValues.slice(1).reduce((map, rowData) => {
    const recipientName = String(rowData[15] || "").trim(); // P열 (0-index 15)
    const rawDate = rowData[4]; // E열 (0-index 4)
    const dateStr = formatDateToYYYYMMDD(rawDate);

    if (recipientName && dateStr) {
      const pk = `${recipientName}_${dateStr}`;
      map.set(pk, {
        provideTime: rowData[12], // M열 (0-index 12) - 제공시간
        startTime: rowData[9], // J열 (0-index 9) - 시작
        endTime: rowData[11], // L열 (0-index 11) - 종료
      });
    }
    return map;
  }, new Map());

  // ----------------------------------------------------
  // 4. I열 오른쪽(J, K, L열)에 연동 열 추가 배치 및 서식 설정
  // ----------------------------------------------------
  Logger.log("4️⃣ [데이터 연동] J, K, L열에 매칭 데이터 배치 중...");
  const newColStart = 10; // J열 (10번째 열)

  // 헤더 입력 (J열: 제공시간, K열: 시작, L열: 종료)
  resultSheet
    .getRange(1, newColStart, 1, 3)
    .setValues([["제공시간", "시작", "종료"]]);

  // 헤더 서식 복사 (I열 서식을 J~L열에 복사)
  resultSheet
    .getRange(1, 9)
    .copyTo(
      resultSheet.getRange(1, newColStart, 1, 3),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false,
    );

  // 연동 열 너비 설정 (서비스내용 입력 시트 원본 너비 적용)
  resultSheet.setColumnWidth(10, serviceSheet.getColumnWidth(13)); // M열 너비 -> J열(제공시간) 적용
  resultSheet.setColumnWidth(11, serviceSheet.getColumnWidth(10)); // J열 너비 -> K열(시작) 적용
  resultSheet.setColumnWidth(12, serviceSheet.getColumnWidth(12)); // L열 너비 -> L열(종료) 적용

  const targetValues = targetSheet.getDataRange().getValues();

  // 매칭 데이터 배열 생성 (B열: 수급자성명, H열: 시작시간/일자)
  const appendedData = targetValues.slice(1).map((rowData) => {
    const recipientName = String(rowData[1] || "").trim(); // B열
    const rawDate = rowData[7]; // H열

    const dateStr = formatDateToYYYYMMDD(rawDate);
    const pk = `${recipientName}_${dateStr}`;

    const match = serviceMap.get(pk);
    return match
      ? [match.provideTime, match.startTime, match.endTime]
      : ["", "", ""];
  });

  if (appendedData.length > 0) {
    resultSheet
      .getRange(2, newColStart, appendedData.length, 3)
      .setValues(appendedData);
  }

  // J, K, L열 데이터 유효성 검사 제거
  resultSheet.getRange(1, newColStart, lastRow, 3).clearDataValidations();

  // ----------------------------------------------------
  // 5. 조건부 서식 적용 (시작시간 vs 시작, 종료시간 vs 종료)
  // ----------------------------------------------------
  Logger.log("5️⃣ [조건부 서식] 동적 색상 조건부 서식 적용 중...");
  const rules = [];

  // 1) 시작시간 vs 시작: H열 vs K열 (hh:mm 비교)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($H2<>"", $K2<>"", TEXT($H2,"hh:mm")=TEXT($K2,"hh:mm"))',
      )
      .setBackground(COLOR_POSITIVE)
      .setRanges([
        resultSheet.getRange(`H2:H${lastRow}`),
        resultSheet.getRange(`K2:K${lastRow}`),
      ])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($H2<>"", $K2<>"", TEXT($H2,"hh:mm")<>TEXT($K2,"hh:mm"))',
      )
      .setBackground(COLOR_NEGATIVE)
      .setRanges([
        resultSheet.getRange(`H2:H${lastRow}`),
        resultSheet.getRange(`K2:K${lastRow}`),
      ])
      .build(),
  );

  // 2) 종료시간 vs 종료: I열 vs L열 (hh:mm 비교)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($I2<>"", $L2<>"", TEXT($I2,"hh:mm")=TEXT($L2,"hh:mm"))',
      )
      .setBackground(COLOR_POSITIVE)
      .setRanges([
        resultSheet.getRange(`I2:I${lastRow}`),
        resultSheet.getRange(`L2:L${lastRow}`),
      ])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND($I2<>"", $L2<>"", TEXT($I2,"hh:mm")<>TEXT($L2,"hh:mm"))',
      )
      .setBackground(COLOR_NEGATIVE)
      .setRanges([
        resultSheet.getRange(`I2:I${lastRow}`),
        resultSheet.getRange(`L2:L${lastRow}`),
      ])
      .build(),
  );

  resultSheet.setConditionalFormatRules(rules);

  // ----------------------------------------------------
  // 6. [후순위 작업] SHEET_TARGET_ALL 서식 적용
  // ----------------------------------------------------
  if (targetAllSheet) {
    Logger.log("6️⃣ [후순위 작업] SHEET_TARGET_ALL 서식 복사 적용 중...");
    targetAllSheet
      .getDataRange()
      .copyTo(
        resultSheet.getRange(1, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false,
      );
  }

  // ----------------------------------------------------
  // 7. 시트 순서 정렬 및 활성화
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
