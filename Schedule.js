/**
 * ============================================================
 *  일정관리 동기화 — 공휴일·마감예정일 수식, 일정관리 시트 서식/구글표
 * ============================================================
 *
 *  건이 등록될 때(SheetWriter.js의 _Sheets에등록 → _일정관리행추가) 자동
 *  호출되는 동기화 로직과, "공휴일·마감예정일 갱신" 메뉴로 독립 실행되는
 *  재계산 로직을 함께 둡니다. 대한민국 공휴일은 공개 캘린더에서 자동 동기화합니다.
 * ============================================================
 */

const 일정관리_헤더색 = '#1f3a5f';
const 일정관리_특이사항헤더색 = '#dbe7f3';
const 일정관리_특이사항헤더글자색 = '#1f3a5f';
const 공휴일시트명 = '공휴일';
const 대한민국공휴일캘린더URL = 'https://calendar.google.com/calendar/ical/ko.south_korea%23holiday%40group.v.calendar.google.com/public/basic.ics';
const 일정관리_접수대장참조맵 = {
  '접수일자': '접수일자',
  '심사접수일': '심사접수일',
  '기업명': '기업명',
  '담당자명': '담당자명',
  '담당자직급': '담당자직급',
  '담당자전화': '담당자전화',
  '담당자휴대전화': '담당자휴대전화',
  '이메일': '이메일',
  '소재지': '소재지',
  '제품명': '제품명',
  '제품수': '제품수',
  '개요': '개요',
  '제공형태': '제공형태',
  '제공형태기타': '제공형태기타',
  '제품분류': '제품분류',
  '인공지능적용목적': '인공지능적용목적',
  '인공지능적용범위': '인공지능적용범위',
  '명세서작성방식': '명세서작성방식',
  '보유인증': '보유인증',
  '인공지능기능수': '인공지능기능수',
};

function _일정관리헤더색적용_(시트) {
  const lastCol = Math.max(1, 시트.getLastColumn());
  const 헤더 = 시트.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

  // 기존 시트도 갱신 시 전체 헤더를 남색으로 통일하고, 특이사항만 연한 톤으로 구분합니다.
  시트.getRange(1, 1, 1, lastCol)
    .setBackground(일정관리_헤더색)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  const 특이사항열 = 헤더.indexOf('특이사항') + 1;
  if (특이사항열 > 0) {
    시트.getRange(1, 특이사항열)
      .setBackground(일정관리_특이사항헤더색)
      .setFontColor(일정관리_특이사항헤더글자색)
      .setFontWeight('bold');
  }
}

function _일정관리서식적용_(시트, 요약뷰) {
  const lastCol = Math.max(1, 시트.getLastColumn());
  const 헤더 = 시트.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

  // 일정관리 날짜는 실제 날짜값을 유지하고 화면에는 두 자리 연도로 간결하게 표시
  ['접수일자', '심사접수일', '마감예정일', '보완요청일', '연장마감일'].forEach(날짜헤더 => {
    const 열 = 헤더.indexOf(날짜헤더) + 1;
    if (열 > 0) 시트.getRange(2, 열, Math.max(1, 시트.getMaxRows() - 1), 1).setNumberFormat('yy-mm-dd');
  });

  시트.setHiddenGridlines(false);
  시트.setFrozenRows(1);
  시트.setFrozenColumns(Math.min(2, lastCol));

  시트.setRowHeight(1, 34);
  시트.getRange(1, 1, Math.max(1, 시트.getMaxRows()), lastCol)
    .setVerticalAlignment('middle');
  시트.getRange(2, 1, Math.max(1, 시트.getMaxRows() - 1), lastCol)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  try {
    _일정관리구글표적용_(시트, 헤더);
  } catch (e) {
    Logger.log('일정관리 Google Sheets 표 적용 실패: ' + e.message);
  }

  _일정관리헤더색적용_(시트);

  const 너비맵 = {
    '순번': 45, '접수번호': 115,
    '접수일자': 90, '심사접수일': 95, '마감예정일': 95,
    '상태': 75, '보완요청일': 95, '연장마감일': 95, '담당심사원': 95, '특이사항': 260,
    '기업명': 155, '담당자명': 85, '담당자직급': 80, '담당자전화': 115, '담당자휴대전화': 115, '이메일': 180, '소재지': 200,
    '제품명': 190, '제품수': 65, '개요': 260,
    '제공형태': 110, '제품분류': 100,
    '인공지능적용목적': 260, '인공지능적용범위': 260,
    '명세서작성방식': 110, '기타제출서류여부': 120, '보유인증': 130,
    '인공지능기능수': 90,
  };
  헤더.forEach((h, idx) => {
    if (너비맵[h]) 시트.setColumnWidth(idx + 1, 너비맵[h]);
  });

  시트.showColumns(1, lastCol);
  if (요약뷰) {
    const 표시컬럼 = new Set([
      '순번', '접수번호',
      '접수일자', '심사접수일', '마감예정일',
      '상태', '보완요청일', '연장마감일', '담당심사원', '특이사항',
      '기업명', '담당자명', '담당자직급', '담당자전화', '담당자휴대전화', '소재지',
      '제품명', '제품수', '제공형태', '제품분류',
      '인공지능기능수',
    ]);
    헤더.forEach((h, idx) => {
      if (h && !표시컬럼.has(h)) 시트.hideColumns(idx + 1);
    });
  }
}

function _일정관리구글표적용_(시트, 헤더) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const sheetId = 시트.getSheetId();
  const tableName = '일정관리_표';
  const lastCol = 헤더.length;
  const lastRow = Math.max(2, 시트.getLastRow());
  const 기존표목록 = _시트표목록조회_(spreadsheetId, sheetId);

  const requests = 기존표목록
    .filter(t => t.name === tableName || (t.range && t.range.sheetId === sheetId))
    .map(t => ({ deleteTable: { tableId: t.tableId } }));

  requests.push({
    addTable: {
      table: {
        name: tableName,
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: lastRow,
          startColumnIndex: 0,
          endColumnIndex: lastCol,
        },
        rowsProperties: {
          headerColorStyle: { rgbColor: _hexToRgb_(일정관리_헤더색) },
          firstBandColorStyle: { rgbColor: _hexToRgb_('#ffffff') },
          secondBandColorStyle: { rgbColor: _hexToRgb_('#f8fbfb') },
        },
        columnProperties: 헤더.map((h, idx) => ({
          columnIndex: idx,
          columnName: h || `Column ${idx + 1}`,
          columnType: _일정관리표컬럼타입_(h),
        })),
      },
    },
  });

  _sheetsBatchUpdate_(spreadsheetId, requests);
}

function _일정관리표컬럼타입_(헤더명) {
  if (['순번', '제품수', '인공지능기능수'].indexOf(헤더명) >= 0) return 'DOUBLE';
  if (['접수일자', '심사접수일', '마감예정일', '보완요청일', '연장마감일'].indexOf(헤더명) >= 0) return 'DATE';
  return 'TEXT';
}

function _시트표목록조회_(spreadsheetId, sheetId) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties(sheetId),tables(tableId,name,range))`;
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Sheets API 표 조회 실패 (${code}): ${res.getContentText()}`);
  }
  const data = JSON.parse(res.getContentText());
  const sheet = (data.sheets || []).find(s => s.properties && s.properties.sheetId === sheetId);
  return sheet && sheet.tables ? sheet.tables : [];
}

function _sheetsBatchUpdate_(spreadsheetId, requests) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ requests: requests }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Sheets API batchUpdate 실패 (${code}): ${res.getContentText()}`);
  }
}

function _hexToRgb_(hex) {
  const value = String(hex || '').replace('#', '');
  const n = parseInt(value, 16);
  return {
    red: ((n >> 16) & 255) / 255,
    green: ((n >> 8) & 255) / 255,
    blue: (n & 255) / 255,
  };
}

/** 인공지능제품모델 시트에 등록 (재파싱 시 같은 접수번호 기존 행 삭제 후 재기록) */
/**
 * 일정관리 시트에 신규 행 추가 (하이브리드 싱크)
 *  · 접수대장이 원본인 컬럼 → 접수대장 INDEX+MATCH 수식으로 자동 참조
 *  · 일정관리가 원본인 컬럼(상태·담당심사원) → 직접 입력값 (편집 가능)
 *  · 마감예정일 → 접수일자로부터 15 WD(주말·공휴일 제외) 수식
 *
 * 접수번호와 대상 컬럼의 실제 헤더 위치를 각각 찾아 수식을 만들므로,
 * 순번을 A열로 옮기거나 컬럼 순서를 바꿔도 정상 참조합니다.
 */
function _일정관리행추가(ss, 접수번호, 직접값) {
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 기존데이터 = 일정시트.getDataRange().getDisplayValues();
  const 기존헤더 = 기존데이터[0].map(v => String(v).trim());
  const 기존접수번호열 = 기존헤더.indexOf('접수번호');
  if (기존접수번호열 >= 0) {
    const 기존행 = 기존데이터.slice(1).findIndex(행 =>
      String(행[기존접수번호열]).trim() === String(접수번호).trim()
    );
    if (기존행 >= 0) {
      Logger.log(`일정관리 기존 행 유지 - append 생략: ${접수번호} (${기존행 + 2}행)`);
      return { 신규: false, 행번호: 기존행 + 2 };
    }
  }
  const 신청연도 = _날짜연도_(직접값.접수일자) || new Date().getFullYear();
  _공휴일연도확보_(ss, [신청연도, 신청연도 + 1]);
  // ⚠️ 실제 시트의 라이브 헤더를 읽음 (정적 정의가 아니라).
  // 심사원이 컬럼을 수동 추가·삽입해도 값이 밀리지 않고 이름 기준으로 배치됨.
  const 일정H = 일정시트.getRange(1, 1, 1, 일정시트.getLastColumn())
    .getValues()[0].map(v => String(v).trim());
  const 대장시트 = ss.getSheetByName(SHEET.접수대장);
  const 대장H = 대장시트.getRange(1, 1, 1, 대장시트.getLastColumn())
    .getValues()[0].map(v => String(v).trim());

  // 접수대장을 원본으로 참조할 컬럼 (일정관리 컬럼명 → 접수대장 컬럼명)
  // 대부분 이름이 같지만, 명세서/기타제출서류는 가공이 필요해 별도 처리
  const 새행번호 = 일정시트.getLastRow() + 1;
  const iD접수 = 일정H.indexOf('접수번호') + 1;
  const 접수열문자 = columnLetter(iD접수);
  const 일정접수셀 = `$${접수열문자}${새행번호}`;

  // 각 컬럼별 값/수식 생성
  const 행값 = 일정H.map(h => {
    // 0) 순번 = 헤더 제외한 현재 행 위치 (새행번호 - 1)
    if (h === '순번') return 새행번호 - 1;

    // 1) 직접 입력값 (상태·담당심사원). 날짜는 접수대장 참조.
    if (h === '접수번호') return 접수번호;
    if (h === '상태') return '대기';
    if (h === '담당심사원' && Object.prototype.hasOwnProperty.call(직접값, h)) return 직접값[h];

    // 2) 마감예정일 = 접수일자로부터 15 WD (토·일·공휴일 제외)
    if (h === '마감예정일') {
      const iD신청 = 일정H.indexOf('접수일자') + 1;
      const 신청셀 = `${columnLetter(iD신청)}${새행번호}`;
      return `=IF(${신청셀}="","",WORKDAY(${신청셀},15,'${공휴일시트명}'!$A$2:$A))`;
    }

    // 2-1) 연장마감일 = 보완요청일로부터 30 WD (토·일·공휴일 제외)
    if (h === '연장마감일') {
      const iD보완 = 일정H.indexOf('보완요청일') + 1;
      const 보완셀 = `${columnLetter(iD보완)}${새행번호}`;
      return `=IF(${보완셀}="","",WORKDAY(${보완셀},30,'${공휴일시트명}'!$A$2:$A))`;
    }

    // 3) 기타제출서류여부 = 접수대장의 파일명 있으면 Y
    if (h === '기타제출서류여부') {
      return _접수대장조회수식_(대장H, '기타제출서류파일명', 일정접수셀, true);
    }

    // 4) 접수대장 참조 컬럼 → 컬럼 순서와 무관한 INDEX+MATCH
    if (일정관리_접수대장참조맵[h]) {
      return _접수대장조회수식_(대장H, 일정관리_접수대장참조맵[h], 일정접수셀, false);
    }

    return '';
  });

  일정시트.appendRow(행값);
  try {
    _일정관리서식적용_(일정시트, true);
  } catch (e) {
    Logger.log('일정관리 표 갱신 실패: ' + e.message);
  }
  return { 신규: true, 행번호: 새행번호 };
}

function _접수대장조회수식_(대장H, 대상헤더, 일정접수셀, 여부표시) {
  const 접수번호열 = 대장H.indexOf('접수번호') + 1;
  const 대상열 = 대장H.indexOf(대상헤더) + 1;
  if (접수번호열 < 1 || 대상열 < 1) return '=""';
  const 접수범위 = `'${SHEET.접수대장}'!$${columnLetter(접수번호열)}:$${columnLetter(접수번호열)}`;
  const 대상범위 = `'${SHEET.접수대장}'!$${columnLetter(대상열)}:$${columnLetter(대상열)}`;
  const 조회식 = `INDEX(${대상범위},MATCH(${일정접수셀},${접수범위},0))`;
  return 여부표시
    ? `=IFERROR(IF(${조회식}="","","Y"),"")`
    : `=IFERROR(IF(${조회식}="","",${조회식}),"")`;
}

/** 기존 일정관리 행의 접수대장 참조 수식만 헤더 기반 수식으로 복구한다. */
function 일정관리참조수식복구() {
  const ui = SpreadsheetApp.getUi();
  if (!_관리자여부()) {
    ui.alert('관리자만 일정관리 참조 수식을 복구할 수 있습니다.');
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 대장시트 = ss.getSheetByName(SHEET.접수대장);
  if (!일정시트 || !대장시트 || 일정시트.getLastRow() < 2) {
    ui.alert('복구할 일정관리 데이터가 없습니다.');
    return;
  }
  const 일정H = 일정시트.getRange(1, 1, 1, 일정시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 대장H = 대장시트.getRange(1, 1, 1, 대장시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 일정접수열 = 일정H.indexOf('접수번호') + 1;
  const 대장접수열 = 대장H.indexOf('접수번호') + 1;
  if (일정접수열 < 1 || 대장접수열 < 1) {
    ui.alert('접수번호 헤더를 찾지 못해 복구를 중단했습니다.');
    return;
  }
  const 행수 = 일정시트.getLastRow() - 1;
  const 접수값 = 일정시트.getRange(2, 일정접수열, 행수, 1).getDisplayValues();
  const 대상행수 = 접수값.filter(행 => String(행[0]).trim()).length;
  if (!대상행수) {
    ui.alert('접수번호가 입력된 일정관리 행이 없습니다.');
    return;
  }
  const 확인 = ui.alert(
    '일정관리 참조 수식 복구',
    `접수번호가 있는 ${대상행수}개 행의 접수대장 참조 컬럼만 INDEX+MATCH 수식으로 다시 설정합니다.\n상태·담당심사원·특이사항은 변경하지 않습니다. 계속할까요?`,
    ui.ButtonSet.OK_CANCEL
  );
  if (확인 !== ui.Button.OK) return;

  Object.keys(일정관리_접수대장참조맵).forEach(일정헤더 => {
    const 일정열 = 일정H.indexOf(일정헤더) + 1;
    if (일정열 < 1) return;
    접수값.forEach((행, idx) => {
      if (!String(행[0]).trim()) return;
      const 접수셀 = `$${columnLetter(일정접수열)}${idx + 2}`;
      일정시트.getRange(idx + 2, 일정열)
        .setFormula(_접수대장조회수식_(대장H, 일정관리_접수대장참조맵[일정헤더], 접수셀, false));
    });
  });
  const 기타열 = 일정H.indexOf('기타제출서류여부') + 1;
  if (기타열 > 0) {
    접수값.forEach((행, idx) => {
      if (!String(행[0]).trim()) return;
      const 접수셀 = `$${columnLetter(일정접수열)}${idx + 2}`;
      일정시트.getRange(idx + 2, 기타열)
        .setFormula(_접수대장조회수식_(대장H, '기타제출서류파일명', 접수셀, true));
    });
  }
  _마감예정일수식갱신_(ss);
  ui.alert(`일정관리 참조 수식 복구 완료: ${대상행수}개 행`);
}

/** 날짜 값(Date 또는 문자열)에서 연도를 추출 */
function _날짜연도_(값) {
  if (값 instanceof Date && !isNaN(값.getTime())) return 값.getFullYear();
  const m = String(값 || '').match(/(20\d{2})/);
  return m ? Number(m[1]) : null;
}

/** 대한민국 공휴일을 자동 보관하는 보조 시트 확보 */
function _공휴일시트확보_(ss) {
  let 시트 = ss.getSheetByName(공휴일시트명);
  if (!시트) {
    시트 = ss.insertSheet(공휴일시트명);
    시트.getRange(1, 1, 1, 2).setValues([['날짜', '공휴일명']])
      .setBackground('#5f6368').setFontColor('#ffffff').setFontWeight('bold');
    시트.setFrozenRows(1);
    시트.setColumnWidth(1, 100);
    시트.setColumnWidth(2, 180);
  }
  return 시트;
}

/**
 * 공개 대한민국 공휴일 캘린더에서 필요한 연도의 법정·대체·임시 공휴일을 자동 동기화한다.
 * 기념일(어버이날·식목일 등)은 WORKDAY 제외일이 아니므로 포함하지 않는다.
 */
function _공휴일연도확보_(ss, 연도목록) {
  const 시트 = _공휴일시트확보_(ss);
  const 필요연도 = Array.from(new Set(연도목록.map(Number).filter(y => y >= 2000 && y <= 2100)));
  if (!필요연도.length) return 시트;

  const 기존값 = 시트.getLastRow() > 1
    ? 시트.getRange(2, 1, 시트.getLastRow() - 1, 2).getValues()
    : [];
  const 기존연도 = new Set(기존값.map(r => _날짜연도_(r[0])).filter(Boolean));
  const 누락연도 = 필요연도.filter(y => !기존연도.has(y));
  if (!누락연도.length) return 시트;

  try {
    const ics = UrlFetchApp.fetch(대한민국공휴일캘린더URL, { muteHttpExceptions: false })
      .getContentText('UTF-8').replace(/\r?\n[ \t]/g, '');
    const 공휴일명패턴 = /(새해|신정|설날|삼일절|3·1절|어린이날|부처님오신날|석가탄신일|현충일|광복절|추석|개천절|한글날|성탄절|크리스마스|선거일|임시공휴일|대체공휴일|쉬는 날)/;
    const 추가값 = [];

    ics.split('BEGIN:VEVENT').slice(1).forEach(block => {
      const 날짜매치 = block.match(/DTSTART;VALUE=DATE:(\d{4})(\d{2})(\d{2})/);
      const 이름매치 = block.match(/(?:^|\n)SUMMARY(?:;[^:]*)?:(.*)/);
      if (!날짜매치 || !이름매치) return;
      const 연도 = Number(날짜매치[1]);
      const 이름 = 이름매치[1].trim().replace(/\\,/g, ',');
      if (누락연도.indexOf(연도) < 0 || !공휴일명패턴.test(이름)) return;
      추가값.push([new Date(연도, Number(날짜매치[2]) - 1, Number(날짜매치[3])), 이름]);
    });

    const 날짜맵 = {};
    기존값.concat(추가값).forEach(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (isNaN(d.getTime())) return;
      const key = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
      날짜맵[key] = [d, r[1]];
    });
    const 전체값 = Object.keys(날짜맵).sort().map(k => 날짜맵[k]);
    if (시트.getLastRow() > 1) 시트.getRange(2, 1, 시트.getLastRow() - 1, 2).clearContent();
    if (전체값.length) {
      시트.getRange(2, 1, 전체값.length, 2).setValues(전체값);
      시트.getRange(2, 1, 전체값.length, 1).setNumberFormat('yy-mm-dd');
    }
  } catch (e) {
    Logger.log('공휴일 자동 동기화 실패(기존 공휴일 목록으로 계산): ' + e.message);
  }
  return 시트;
}

/** 일정관리의 기본 마감(15 WD)과 보완 연장마감(30 WD) 수식을 모두 갱신 */
function _마감예정일수식갱신_(ss) {
  const 시트 = ss.getSheetByName(SHEET.일정관리);
  if (!시트 || 시트.getLastRow() < 2) {
    _공휴일연도확보_(ss, [new Date().getFullYear(), new Date().getFullYear() + 1]);
    return 0;
  }
  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 신청열 = 헤더.indexOf('접수일자') + 1;
  const 마감열 = 헤더.indexOf('마감예정일') + 1;
  const 보완요청열 = 헤더.indexOf('보완요청일') + 1;
  const 연장마감열 = 헤더.indexOf('연장마감일') + 1;
  if (신청열 < 1 || 마감열 < 1) return 0;

  const 행수 = 시트.getLastRow() - 1;
  const 신청값 = 시트.getRange(2, 신청열, 행수, 1).getValues().flat();
  const 보완요청값 = 보완요청열 > 0
    ? 시트.getRange(2, 보완요청열, 행수, 1).getValues().flat()
    : [];
  // 기본 마감뿐 아니라 연장마감 계산에 필요한 연도의 공휴일도 확보한다.
  // 보완요청일이 접수일자의 다음 해 이후인 경우 접수일자만 보면 해당 공휴일이 누락될 수 있다.
  const 연도목록 = 신청값.concat(보완요청값).map(_날짜연도_).filter(Boolean);
  const 현재연도 = new Date().getFullYear();
  연도목록.push(현재연도, 현재연도 + 1);
  _공휴일연도확보_(ss, 연도목록.concat(연도목록.map(y => y + 1)));

  const 신청열문자 = columnLetter(신청열);
  const 수식 = 신청값.map((_, i) => {
    const 행 = i + 2;
    const 신청셀 = `${신청열문자}${행}`;
    return [`=IF(${신청셀}="","",WORKDAY(${신청셀},15,'${공휴일시트명}'!$A$2:$A))`];
  });
  시트.getRange(2, 마감열, 행수, 1).setFormulas(수식).setNumberFormat('yy-mm-dd');
  if (보완요청열 > 0 && 연장마감열 > 0) {
    const 보완요청열문자 = columnLetter(보완요청열);
    const 연장수식 = 신청값.map((_, i) => {
      const 행 = i + 2;
      const 보완셀 = `${보완요청열문자}${행}`;
      return [`=IF(${보완셀}="","",WORKDAY(${보완셀},30,'${공휴일시트명}'!$A$2:$A))`];
    });
    // 구버전에서 남아 있을 수 있는 수기 날짜 입력 검사를 먼저 제거해야
    // 자동 수식 입력 시 "날짜를 직접 선택" 유효성 검사 예외가 발생하지 않는다.
    시트.getRange(2, 연장마감열, 행수, 1)
      .clearDataValidations()
      .setFormulas(연장수식)
      .setNumberFormat('yy-mm-dd');
  }
  return 행수;
}

/** 관리자 수동 실행용: 공휴일과 기존 마감예정일 수식을 즉시 갱신 */
function 마감예정일갱신() {
  const 갱신건수 = _마감예정일수식갱신_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert(`마감일 갱신 완료: ${갱신건수}건\n기본: 접수일자 + 15 WD\n보완: 보완요청일 + 30 WD\n(주말·공휴일 제외)`);
}
