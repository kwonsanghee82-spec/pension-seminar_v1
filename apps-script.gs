/**
 * 미래에셋증권 퇴직연금 세미나 - 응답 저장용 Google Apps Script
 *
 * 설치 방법
 * 1) 새 Google Sheets 문서를 만들고, 시트(탭) 이름을 "설문", "출석" 으로 각각 만들어 둡니다.
 *    (헤더 행은 자동으로 채워지므로 비워 두어도 됩니다)
 * 2) 문서 메뉴에서 확장 프로그램 > Apps Script 를 엽니다.
 * 3) 기본 Code.gs 내용을 지우고 이 파일 내용 전체를 붙여넣습니다.
 * 4) 오른쪽 위 "배포" > "새 배포" > 유형 선택(⚙️) > "웹 앱" 선택
 *      - 실행 계정: 나
 *      - 액세스 권한이 있는 사용자: 전체(익명 포함)
 *    를 선택한 뒤 배포합니다.
 * 5) 배포 완료 후 나오는 "웹 앱 URL"을 복사해서, index.html과 함께 있는
 *    app.jsx 파일 상단의 GAS_URL 상수에 붙여넣습니다.
 */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const sheetName = body.sheet;
    const data = body.data || {};
    const sheet = getOrCreateSheet_(sheetName);

    // 첫 행(헤더)이 비어있으면 data의 키로 헤더를 만든다.
    const headers = getOrInitHeaders_(sheet, Object.keys(data));

    const row = headers.map((h) => {
      const v = data[h];
      if (v === undefined || v === null) return "";
      if (typeof v === "object") return JSON.stringify(v); // ratings, topics 등
      return v;
    });
    sheet.appendRow(row);

    return jsonResponse_({ ok: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    const sheetName = e.parameter.sheet;
    const sheet = getOrCreateSheet_(sheetName);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return jsonResponse_({ ok: true, rows: [] });

    const headers = values[0];
    const rows = values.slice(1).map((r) => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
    return jsonResponse_({ ok: true, rows });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function getOrInitHeaders_(sheet, keys) {
  const firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const hasHeaders = firstRow.some((v) => v !== "");
  if (hasHeaders) return firstRow.filter((v) => v !== "");
  sheet.getRange(1, 1, 1, keys.length).setValues([keys]);
  return keys;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
