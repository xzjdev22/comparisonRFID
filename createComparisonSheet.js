// ==========================================
// 1. 코드 최상단 상수 정의
// ==========================================
const SHEET_REALTIME = "실시간전송내용";
const SHEET_TARGET = "청구대상건";
const SHEET_SERVICE = "서비스내용 입력";
const SHEET_RESULT = "실전송대비 청구기록 비교";
const COLOR_POSITIVE = "#d9ead3";
const COLOR_NEGATIVE = "#f4cccc";

/**
 * 실전송대비 청구기록 비교 시트 생성 및 서식/조건부서식 적용 함수
 */
function createComparisonSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 대상 시트 가져오기
  const targetSheet = ss.getSheetByName(SHEET_TARGET);
  const serviceSheet = ss.getSheetByName(SHEET_SERVICE);

  if (!targetSheet || !serviceSheet) {
    SpreadsheetApp.getUi().alert(
      `'${SHEET_TARGET}' 또는 '${SHEET_SERVICE}' 시트를 찾을 수 없습니다.`,
    );
    return;
  }

  // 1. 기존 결과 시트 제거 후 새로 생성
  let resultSheet = ss.getSheetByName(SHEET_RESULT);
  if (resultSheet) {
    ss.deleteSheet(resultSheet);
  }
  resultSheet = ss.insertSheet(SHEET_RESULT);

  // ----------------------------------------------------
  // 2. 청구대상건 원본 전체 복사 (서식, 폰트, 멘션, 링크, 행/열 너비 보존)
  // ----------------------------------------------------
  const targetRange = targetSheet.getDataRange();
  const lastRow = targetRange.getLastRow();
  const lastCol = targetRange.getLastColumn();

  if (lastRow < 1) return;

  // 전체 복사 (값, 서식, 링크, 스마트 멘션 포함)
  targetRange.copyTo(
    resultSheet.getRange(1, 1),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false,
  );

  // 열 너비 및 행 높이 복사
  for (let c = 1; c <= lastCol; c++) {
    resultSheet.setColumnWidth(c, targetSheet.getColumnWidth(c));
  }
  for (let r = 1; r <= lastRow; r++) {
    resultSheet.setRowHeight(r, targetSheet.getRowHeight(r));
  }

  // ----------------------------------------------------
  // 3. C열(인정번호) 및 E열(핸드폰번호) 제거
  // ----------------------------------------------------
  // 오른쪽 열(E열 = 5)부터 먼저 삭제해야 열 인덱스가 밀리지 않음
  resultSheet.deleteColumn(5); // 원본 E열 삭제
  resultSheet.deleteColumn(3); // 원본 C열 삭제

  // ----------------------------------------------------
  // 4. 서비스내용 입력 시트 데이터 맵핑 (PK: 수급자명 + YYYY-MM-DD)
  // ----------------------------------------------------
  // 서비스내용 입력 시트 필터 확인
  const serviceFilter = serviceSheet.getFilter();
  const serviceValues = serviceSheet.getDataRange().getValues();
  const serviceMap = new Map();

  // 서비스내용 입력 데이터 읽기 (2행부터 헤더 제외)
  for (let r = 2; r <= serviceValues.length; r++) {
    // isRowHiddenByFilter 활용
    if (serviceFilter && serviceFilter.isRowHiddenByFilter(r)) continue;

    const rowData = serviceValues[r - 1];
    const recipientName = String(rowData[15] || "").trim(); // P열 (0-index 15)
    const rawDate = rowData[4]; // E열 (0-index 4)

    const dateStr = formatDateToYYYYMMDD(rawDate);
    if (!recipientName || !dateStr) continue;

    const pk = `${recipientName}_${dateStr}`;

    // 매칭 데이터 저장 (M열: 제공시간, J열: 시작, L열: 종료)
    // 원본 가져오기 (시간 서식 지키기 위해)
    const serviceRange = serviceSheet.getRange(
      r,
      1,
      1,
      serviceSheet.getLastColumn(),
    );

    serviceMap.set(pk, {
      row: r,
      provideTime: rowData[12], // M열 (0-index 12)
      startTime: rowData[9], // J열 (0-index 9)
      endTime: rowData[11], // L열 (0-index 11)
      range: serviceRange,
    });
  }

  // ----------------------------------------------------
  // 5. 추가 열 (제공시간, 시작, 종료) 헤더 및 데이터 세팅
  // ----------------------------------------------------
  // C, E열 제거로 인해 종료시간은 G열(7)에 위치함 -> H, I, J열에 추가
  const newColStart = 8; // H열

  // 헤더 추가
  resultSheet.getRange(1, newColStart).setValue("제공시간");
  resultSheet.getRange(1, newColStart + 1).setValue("시작");
  resultSheet.getRange(1, newColStart + 2).setValue("종료");

  // 서비스내용 입력 시트의 원본 열 너비 복사 (M열: 13, J열: 10, L열: 12)
  resultSheet.setColumnWidth(newColStart, serviceSheet.getColumnWidth(13));
  resultSheet.setColumnWidth(newColStart + 1, serviceSheet.getColumnWidth(10));
  resultSheet.setColumnWidth(newColStart + 2, serviceSheet.getColumnWidth(12));

  // 헤더 서식 복사 (기존 G열 헤더 서식 적용)
  resultSheet
    .getRange(1, 7)
    .copyTo(
      resultSheet.getRange(1, newColStart, 1, 3),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false,
    );

  // 청구대상건 데이터 읽기 (필터 반영 및 PK 매칭)
  const targetFilter = targetSheet.getFilter();
  const targetValues = targetSheet.getDataRange().getValues();

  // 매칭된 데이터 결과를 담을 배열
  const appendedData = [];

  for (let r = 2; r <= lastRow; r++) {
    // isRowHiddenByFilter 활용하여 숨겨진 행 제외
    if (targetFilter && targetFilter.isRowHiddenByFilter(r)) {
      appendedData.push(["", "", ""]);
      continue;
    }

    const rowData = targetValues[r - 1];
    const recipientName = String(rowData[1] || "").trim(); // B열 (수급자성명, 0-index 1)
    const rawDate = rowData[7]; // H열 (시작시간, 0-index 7)

    const dateStr = formatDateToYYYYMMDD(rawDate);
    const pk = `${recipientName}_${dateStr}`;

    if (serviceMap.has(pk)) {
      const match = serviceMap.get(pk);

      // 원본 서식 복사 (M열 -> H열, J열 -> I열, L열 -> J열)
      serviceSheet
        .getRange(match.row, 13)
        .copyTo(
          resultSheet.getRange(r, newColStart),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false,
        );
      serviceSheet
        .getRange(match.row, 10)
        .copyTo(
          resultSheet.getRange(r, newColStart + 1),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false,
        );
      serviceSheet
        .getRange(match.row, 12)
        .copyTo(
          resultSheet.getRange(r, newColStart + 2),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false,
        );

      appendedData.push([match.provideTime, match.startTime, match.endTime]);
    } else {
      appendedData.push(["", "", ""]);
    }
  }

  // 데이터 한 번에 입력
  if (appendedData.length > 0) {
    resultSheet
      .getRange(2, newColStart, appendedData.length, 3)
      .setValues(appendedData);
  }

  // ----------------------------------------------------
  // 6. 동적 조건부 서식 (Conditional Formatting Rules) 적용
  // ----------------------------------------------------
  /*
    결과 시트 열 구조:
    - E열: 총시간 (기존 G열)
    - F열: 시작시간 (기존 H열)
    - G열: 종료시간 (기존 I열)
    - H열: 제공시간 (추가)
    - I열: 시작 (추가)
    - J열: 종료 (추가)
  */
  const rules = [];

  // 1) 총시간(E열) vs 제공시간(H열)
  // 긍정: 정수값 일치
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($E2<>"", $H2<>"", INT($E2)=INT($H2))')
      .setBackground(COLOR_POSITIVE)
      .setRanges([
        resultSheet.getRange(`E2:E${lastRow}`),
        resultSheet.getRange(`H2:H${lastRow}`),
      ])
      .build(),
  );
  // 부정: 다름
  rules.push(
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
  );
  rules.push(
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
  );
  rules.push(
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
  // 순서: 실시간전송내용 -> 청구대상건 -> 서비스내용 입력 -> 실전송대비 청구기록 비교
  // ----------------------------------------------------
  const orderList = [SHEET_REALTIME, SHEET_TARGET, SHEET_SERVICE, SHEET_RESULT];
  orderList.forEach((sheetName, index) => {
    const sh = ss.getSheetByName(sheetName);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(index + 1);
    }
  });

  // 작업 후 결과 시트 활성화
  ss.setActiveSheet(resultSheet);
  SpreadsheetApp.getUi().alert(
    "실전송대비 청구기록 비교 시트 생성이 완료되었습니다.",
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
