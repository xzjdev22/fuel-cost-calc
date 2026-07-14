/**
 * 교통비 계산 및 전월 대비 분석 시스템 (Final)
 * * 주요 기능:
 * 1. 카카오 API를 이용한 구간별 최단 거리 산출
 * 2. 작업자별 개별 교통비 상세 시트 자동 생성
 * 3. 최종결과 시트에서 전월 데이터와 실시간 비교 분석
 */

// [설정] 카카오 REST API 키 (스크립트 속성에서 로드)
const KAKAO_REST_API_KEY =
  PropertiesService.getScriptProperties().getProperty("KAKAO_REST_API_KEY");

/**
 * 메인 실행 함수: 모든 프로세스를 제어합니다.
 */
function calculateAllWorkersRoute() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originSheet = ss.getSheetByName("1_출발지");
  const destSheet = ss.getSheetByName("2_도착지");

  console.log("--- 시스템 가동 시작 ---");

  // 기초 시트 및 API 키 검증
  if (!originSheet || !destSheet) {
    console.error("오류: '1_출발지' 또는 '2_도착지' 시트가 없습니다.");
    SpreadsheetApp.getUi().alert(
      "'1_출발지' 또는 '2_도착지' 시트를 찾을 수 없습니다."
    );
    return;
  }
  if (!KAKAO_REST_API_KEY) {
    SpreadsheetApp.getUi().alert(
      "스크립트 속성에 'KAKAO_REST_API_KEY'를 설정해주세요."
    );
    return;
  }

  // 1. 출발지 데이터 맵핑 (이름: B열[1], 주소: C열[2])
  const originMap = {};
  const originValues = originSheet.getDataRange().getValues();
  for (let i = 1; i < originValues.length; i++) {
    const name = originValues[i][1]; // B열
    if (name) {
      originMap[name] = {
        address: cleanAddress(originValues[i][2]), // C열
        remoteTargets: (originValues[i][3] || "")
          .toString()
          .split(",")
          .map((t) => t.trim()),
        status: originValues[i][4],
        rowIdx: i + 1,
      };
    }
  }
  console.log("출발지 로드 완료: " + Object.keys(originMap).length + "명");

  // 2. 도착지 데이터 맵핑 (출발지와 동일하게 이름: B열[1], 주소: C열[2] 방식 적용)
  const destMap = {};
  const destValues = destSheet.getDataRange().getValues();
  for (let i = 1; i < destValues.length; i++) {
    const placeName = destValues[i][1]
      ? destValues[i][1].toString().trim()
      : ""; // B열[1]
    const placeAddr = destValues[i][2]; // C열[2]
    if (placeName) {
      destMap[placeName] = cleanAddress(placeAddr);
    }
  }
  console.log("도착지 로드 완료: " + Object.keys(destMap).length + "곳");
  // 디버깅: 로드된 장소명 일부 출력
  console.log(
    "로드된 장소 예시: " + Object.keys(destMap).slice(0, 3).join(", ")
  );

  // 3. 최종결과 시트 초기 세팅
  let summarySheet =
    ss.getSheetByName("0_최종결과") || ss.insertSheet("0_최종결과", 0);
  const prevUrl = summarySheet.getRange("E2").getValue();

  const existingSupportFunds = {};
  const lastRowOfSummary = summarySheet.getLastRow();
  if (lastRowOfSummary >= 5) {
    const data = summarySheet
      .getRange(5, 2, lastRowOfSummary - 4, 5)
      .getValues();
    data.forEach((row) => {
      if (row[0]) existingSupportFunds[row[0]] = row[4] || 0;
    });
  }

  summarySheet.clear().clearFormats();
  if (prevUrl) summarySheet.getRange("E2").setValue(prevUrl);

  // 상단 디자인 설정
  summarySheet
    .getRange("B1")
    .setValue("제어값 ▶")
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#444444")
    .setHorizontalAlignment("center");
  summarySheet
    .getRange("C1:E1")
    .setValues([
      [
        formatHeader("적용금액"),
        formatHeader("원거리 교통비"),
        formatHeader("전월 교통비 시트(URL)"),
      ],
    ]);
  if (!summarySheet.getRange("C2").getValue())
    summarySheet.getRange("C2").setValue(1000);
  if (!summarySheet.getRange("D2").getValue())
    summarySheet.getRange("D2").setValue(5100);
  summarySheet.getRange("C1:E1").setBackground("#fff2cc").setFontWeight("bold");
  summarySheet.getRange("C2:D2").setNumberFormat('#,##0"원"');
  summarySheet
    .getRange("B1:E2")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#cccccc",
      SpreadsheetApp.BorderStyle.SOLID
    );

  const mainHeaders = [
    "No",
    "이름",
    "총 거리",
    "구간교통비",
    "원거리교통비",
    "기타지원금",
    "총 교통비",
    "전월 교통비",
    "전월 대비",
  ].map(formatHeader);
  summarySheet
    .getRange(4, 1, 1, 9)
    .setValues([mainHeaders])
    .setBackground("#eeeeee")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#cccccc",
      SpreadsheetApp.BorderStyle.SOLID
    );
  summarySheet.getRange(4, 7).setBackground("#fff2cc");

  const distanceCache = {};
  const startTime = new Date().getTime();
  const processedWorkers = [];

  // 4. 대상 시트 순회 (9_...의 사본)
  const allSheets = ss.getSheets();
  allSheets.forEach((sheet) => {
    const sName = sheet.getName();
    if (sName.startsWith("9_") && sName.endsWith("의 사본")) {
      const workerName = sName.replace("9_", "").replace("의 사본", "").trim();
      const workerInfo = originMap[workerName];

      if (workerInfo && workerInfo.status !== "완료") {
        if (new Date().getTime() - startTime > 300000) return;

        console.log(">>> 처리 중: " + workerName);
        processIndividualWorker(
          ss,
          workerName,
          workerInfo.address,
          sheet,
          workerInfo.remoteTargets,
          distanceCache,
          destMap
        );

        originSheet.getRange(workerInfo.rowIdx, 5).setValue("완료");
        processedWorkers.push(workerName);
        SpreadsheetApp.flush();
      }
    }
  });

  updateSummaryAndCompare(
    ss,
    summarySheet,
    processedWorkers,
    existingSupportFunds
  );
  fixAllColumnWidths(ss);
  sortSheetsByName(ss);

  console.log("--- 모든 작업 완료 ---");
  SpreadsheetApp.getUi().alert("✅ 작업이 완료되었습니다.");
}

/**
 * 개인별 상세 시트 생성
 */
function processIndividualWorker(
  ss,
  workerName,
  homeAddress,
  dataSheet,
  remoteTargetList,
  distanceCache,
  destMap
) {
  const dataValues = dataSheet.getDataRange().getValues();
  let rawEvents = [];

  for (let j = 1; j < dataValues.length; j++) {
    const placeName = dataValues[j][0]
      ? dataValues[j][0].toString().trim()
      : "";
    const timeVal = dataValues[j][1];
    if (!placeName || !timeVal) continue;

    const targetAddr = destMap[placeName];
    if (!targetAddr) {
      console.warn(
        "  [매칭 실패] 시트의 장소명: '" +
          placeName +
          "' (2_도착지 시트에 이 이름이 있는지 확인하세요)"
      );
      continue;
    }

    const startTime = new Date(timeVal);
    rawEvents.push({
      time: startTime,
      date: Utilities.formatDate(startTime, "GMT+9", "yyyy-MM-dd"),
      name: placeName,
      address: targetAddr,
    });
  }

  if (rawEvents.length === 0) {
    console.error(
      "  [중단] " + workerName + "님은 매칭된 주소 데이터가 하나도 없습니다."
    );
    return;
  }
  rawEvents.sort((a, b) => a.time - b.time);

  const resultSheetName = `3_교통비_${workerName}`;
  let resultSheet =
    ss.getSheetByName(resultSheetName) || ss.insertSheet(resultSheetName);
  resultSheet.clear();

  resultSheet.setFrozenRows(1);
  const rowHeaders = [
    "일자",
    "구간",
    "출발지",
    "도착지",
    "거리",
    "교통비",
    "원거리교통비",
  ].map(formatHeader);
  resultSheet.appendRow(rowHeaders);
  resultSheet
    .getRange(1, 1, 1, 7)
    .setBackground("#eeeeee")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#cccccc",
      SpreadsheetApp.BorderStyle.SOLID
    );

  const groupedByDate = rawEvents.reduce((acc, obj) => {
    (acc[obj.date] = acc[obj.date] || []).push(obj);
    return acc;
  }, {});

  for (const date in groupedByDate) {
    const dailyEvents = groupedByDate[date];
    let prevAddr = homeAddress;

    dailyEvents.forEach((event, idx) => {
      const fromName = idx === 0 ? "자택" : dailyEvents[idx - 1].name;
      const toAddr = event.address;
      const cacheKey = `${prevAddr}|${toAddr}`;

      let dist =
        prevAddr === toAddr
          ? 0
          : distanceCache[cacheKey] ?? getKakaoDistance(prevAddr, toAddr);
      distanceCache[cacheKey] = dist;

      let isRemote = remoteTargetList.includes(event.name);
      let remoteFormula = isRemote ? "='0_최종결과'!$D$2" : 0;

      resultSheet.appendRow([
        date,
        `구간 ${idx + 1}`,
        fromName,
        event.name,
        dist.toFixed(2),
        "",
        remoteFormula,
      ]);
      prevAddr = toAddr;
    });
    resultSheet.appendRow(["", "", "", "", "", "", ""]);
  }

  const lastRow = resultSheet.getLastRow() + 1;
  const fareFormula = `=IF(E${lastRow}=0, 0, '0_최종결과'!$C$2 * CEILING(E${lastRow}/5))`;

  resultSheet.appendRow([
    "",
    "",
    "",
    "월별 총계",
    `=SUM(E2:E${lastRow - 1})`,
    fareFormula,
    `=SUM(G2:G${lastRow - 1})`,
  ]);
  resultSheet.appendRow([
    "",
    "",
    "",
    "총 합계금액",
    "",
    "",
    `=F${lastRow}+G${lastRow}`,
  ]);

  const dataRange = resultSheet.getRange(2, 1, lastRow - 1, 7);
  dataRange.applyRowBanding(
    SpreadsheetApp.BandingTheme.LIGHT_GREY,
    false,
    false
  );
  resultSheet.getRange(2, 1, lastRow, 4).setHorizontalAlignment("center");
  resultSheet.getRange(2, 5, lastRow, 3).setHorizontalAlignment("right");
  resultSheet
    .getRange(2, 1, lastRow, 7)
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#cccccc",
      SpreadsheetApp.BorderStyle.SOLID
    );
  resultSheet
    .getRange(lastRow, 4, 2, 4)
    .setBackground("#d9ead3")
    .setFontWeight("bold");
  resultSheet.getRange(lastRow, 5).setNumberFormat('#,##0.00"km"');
  resultSheet.getRange(lastRow, 6, 2, 2).setNumberFormat('#,##0"원"');
}

/**
 * 최종결과 업데이트 (서식 버그 수정 버전)
 */
function updateSummaryAndCompare(
  ss,
  summarySheet,
  processedWorkers,
  existingSupportFunds
) {
  let rowIdx = 5;
  processedWorkers.forEach((workerName, i) => {
    const resSheetName = `3_교통비_${workerName}`;
    const resSheet = ss.getSheetByName(resSheetName);
    if (!resSheet) return;

    const lastRow = resSheet.getLastRow();
    const monthlyTotalRow = lastRow - 1;
    const supportFund = existingSupportFunds[workerName] || 0;

    summarySheet
      .getRange(rowIdx, 1, 1, 9)
      .setFormulas([
        [
          i + 1,
          `="${workerName}"`,
          `=ROUND('${resSheetName}'!E${monthlyTotalRow}, 0)`,
          `='${resSheetName}'!F${monthlyTotalRow}`,
          `='${resSheetName}'!G${monthlyTotalRow}`,
          supportFund,
          `='${resSheetName}'!G${lastRow} + F${rowIdx}`,
          `=VLOOKUP(B${rowIdx}, IMPORTRANGE($E$2, "'0_최종결과'!$B$5:$G$100"), 6, FALSE)`,
          `=G${rowIdx}-H${rowIdx}`,
        ],
      ]);
    rowIdx++;
  });

  const finalDataRow = rowIdx - 1;
  if (finalDataRow >= 5) {
    // 1. 데이터 영역 줄무늬 적용 (총계행 제외)
    const dataRange = summarySheet.getRange(5, 1, finalDataRow - 4, 9);
    dataRange.applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY,
      false,
      false
    );

    // 2. 총계 행 설정
    summarySheet.getRange(rowIdx, 2).setValue("총계");
    for (let col = 3; col <= 9; col++) {
      let colLetter = String.fromCharCode(64 + col);
      summarySheet
        .getRange(rowIdx, col)
        .setFormula(`=SUM(${colLetter}5:${colLetter}${finalDataRow})`);
    }

    // 3. 총계 행 디자인 적용 (배경 #34495e, 글자 흰색)
    const totalRowRange = summarySheet.getRange(rowIdx, 1, 1, 8);
    totalRowRange
      .setBackground("#34495e")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#cccccc",
        SpreadsheetApp.BorderStyle.SOLID
      );

    // 4. 개별 열 강조 (총 교통비 열 노란색 배경)
    summarySheet
      .getRange(5, 7, finalDataRow - 4, 1)
      .setBackground("#fff2cc")
      .setFontWeight("bold");
  }

  summarySheet.getRange("C5:C" + rowIdx).setNumberFormat('#,##0"km"');
  summarySheet.getRange("D5:H" + rowIdx).setNumberFormat('#,##0"원"');
  summarySheet
    .getRange("I5:I" + rowIdx)
    .setNumberFormat('[Red]+#,##0"원";[Blue]-#,##0"원";0"원"');
}

/** 텍스트 포맷팅 */
function formatHeader(text) {
  return text ? text.replace("(", "\n(") : "";
}

/** 시트 정렬 */
function sortSheetsByName(ss) {
  const sheets = ss.getSheets();
  const sheetNames = sheets.map((s) => s.getName()).sort();
  for (let i = 0; i < sheetNames.length; i++) {
    ss.getSheetByName(sheetNames[i]).activate();
    ss.moveActiveSheet(i + 1);
  }
}

/** 컬럼 폭 조정 */
function fixAllColumnWidths(ss) {
  ss.getSheets().forEach((sheet) => {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const name = sheet.getName();
    if (lastCol > 0) {
      sheet.setRowHeight(1, 45);
      if (lastRow > 1) sheet.setRowHeights(2, lastRow - 1, 28);
      if (name === "0_최종결과") {
        sheet.setRowHeight(4, 45);
        const widths = [30, 80, 80, 90, 100, 100, 120, 110, 110];
        widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
      } else if (name.startsWith("3_교통비_")) {
        const widths = [110, 80, 130, 130, 90, 100, 120];
        widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
      } else {
        sheet.autoResizeColumns(1, lastCol);
      }
    }
  });
}

/** 카카오 API 거리 측정 */
function getKakaoDistance(origin, destination) {
  if (!origin || !destination || origin === destination) return 0;
  try {
    const start = getCoords(origin);
    const end = getCoords(destination);
    if (!start || !end) return 0;
    const url = `https://apis-navi.kakaomobility.com/v1/directions?origin=${start.lng},${start.lat}&destination=${end.lng},${end.lat}&priority=DISTANCE&summary=true`;
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "KakaoAK " + KAKAO_REST_API_KEY },
      muteHttpExceptions: true,
    });
    const data = JSON.parse(response.getContentText());
    return data.routes && data.routes[0].result_code === 0
      ? data.routes[0].summary.distance / 1000
      : 0;
  } catch (e) {
    return 0;
  }
}

/** 위경도 변환 */
function getCoords(address) {
  try {
    const url = `https://dapi.kakao.com/v2/local/search/address.json?query=${encodeURIComponent(
      address
    )}`;
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "KakaoAK " + KAKAO_REST_API_KEY },
      muteHttpExceptions: true,
    });
    const data = JSON.parse(response.getContentText());
    return data.documents && data.documents.length > 0
      ? { lat: data.documents[0].y, lng: data.documents[0].x }
      : null;
  } catch (e) {
    return null;
  }
}

/** 주소 정규화 (괄호 제거) */
function cleanAddress(addr) {
  return addr ? addr.toString().split("(")[0].trim() : "";
}
