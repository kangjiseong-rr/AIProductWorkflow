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

/** 더 이상 쓰지 않는 옛 컬럼(심사링크, 보고서생성)이 남아있으면 완전히 삭제 */
function _일정관리레거시컬럼삭제_(시트) {
  ['심사링크', '보고서생성'].forEach(옛컬럼명 => {
    const lastCol = 시트.getLastColumn();
    if (lastCol < 1) return;
    const 헤더 = 시트.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
    const iOld = 헤더.indexOf(옛컬럼명);
    if (iOld >= 0) 시트.deleteColumn(iOld + 1);
  });
}

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
  _일정관리레거시컬럼삭제_(시트);
  const lastCol = Math.max(1, 시트.getLastColumn());
  const 헤더 = 시트.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

  // 일정관리 날짜는 실제 날짜값을 유지하고 화면에는 두 자리 연도로 간결하게 표시
  ['신청일', '심사접수일', '마감예정일', '보완요청일', '연장마감일'].forEach(날짜헤더 => {
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
    '신청일': 90, '심사접수일': 95, '마감예정일': 95,
    '상태': 75, '보완요청일': 95, '연장마감일': 95, '담당심사원': 95, '특이사항': 260,
    '기업명': 155, '담당자명': 85, '연락처': 115, '이메일': 180, '소재지': 200,
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
      '신청일', '심사접수일', '마감예정일',
      '상태', '보완요청일', '연장마감일', '담당심사원', '특이사항',
      '기업명', '담당자명', '연락처', '소재지',
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
  if (['신청일', '심사접수일', '마감예정일', '보완요청일', '연장마감일'].indexOf(헤더명) >= 0) return 'DATE';
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
 *  · 접수대장이 원본인 컬럼 → 접수대장 VLOOKUP 수식으로 자동 참조
 *  · 일정관리가 원본인 컬럼(상태·담당심사원) → 직접 입력값 (편집 가능)
 *  · 마감예정일 → 신청일로부터 15 WD(주말·공휴일 제외) 수식
 *
 * VLOOKUP은 접수번호(A열)를 키로 하므로, 접수대장 행이 정렬·이동돼도
 * 항상 올바른 값을 따라갑니다. 접수대장에서 회사·제품정보를 고치면
 * 일정관리에 자동 반영됩니다.
 */
function _일정관리행추가(ss, 접수번호, 직접값) {
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 신청연도 = _날짜연도_(직접값.신청일) || new Date().getFullYear();
  _공휴일연도확보_(ss, [신청연도, 신청연도 + 1]);
  // ⚠️ 실제 시트의 라이브 헤더를 읽음 (정적 정의가 아니라).
  // 심사원이 컬럼을 수동 추가·삽입해도 값이 밀리지 않고 이름 기준으로 배치됨.
  const 일정H = 일정시트.getRange(1, 1, 1, 일정시트.getLastColumn())
    .getValues()[0].map(v => String(v).trim());
  const 대장H = 시트헤더정의[SHEET.접수대장];

  // 접수대장을 원본으로 참조할 컬럼 (일정관리 컬럼명 → 접수대장 컬럼명)
  // 대부분 이름이 같지만, 명세서/기타제출서류는 가공이 필요해 별도 처리
  const 참조맵 = {
    '기업명': '기업명',
    '담당자명': '담당자명',
    '연락처': '연락처',
    '이메일': '이메일',
    '소재지': '소재지',
    '제품명': '제품명',
    '제품수': '제품수',
    '개요': '개요',
    '제공형태': '제공형태',
    '제품분류': '제품분류',
    '인공지능적용목적': '인공지능적용목적',
    '인공지능적용범위': '인공지능적용범위',
    '명세서작성방식': '명세서작성방식',
    '보유인증': '보유인증',
    '인공지능기능수': '인공지능기능수',
  };

  const 새행번호 = 일정시트.getLastRow() + 1;
  const 대장범위 = `${SHEET.접수대장}!$A:$${columnLetter(대장H.length)}`;
  const iD접수 = 일정H.indexOf('접수번호') + 1;
  const 접수열문자 = columnLetter(iD접수);

  // 각 컬럼별 값/수식 생성
  const 행값 = 일정H.map(h => {
    // 0) 순번 = 헤더 제외한 현재 행 위치 (새행번호 - 1)
    if (h === '순번') return 새행번호 - 1;

    // 1) 직접 입력값 (상태·담당심사원·신청일·심사접수일)
    if (h === '접수번호') return 접수번호;
    if (h === '상태') return '대기';
    if (Object.prototype.hasOwnProperty.call(직접값, h)) return 직접값[h];

    // 2) 마감예정일 = 신청일로부터 15 WD (토·일·공휴일 제외)
    if (h === '마감예정일') {
      const iD신청 = 일정H.indexOf('신청일') + 1;
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
      return `=IF(VLOOKUP($${접수열문자}${새행번호},${대장범위},${_대장열번호('기타제출서류파일명', 대장H)},0)="","","Y")`;
    }

    // 4) 접수대장 참조 컬럼 → VLOOKUP
    if (참조맵[h]) {
      const 열번호 = _대장열번호(참조맵[h], 대장H);
      return `=IFERROR(VLOOKUP($${접수열문자}${새행번호},${대장범위},${열번호},0),"")`;
    }

    return '';
  });

  일정시트.appendRow(행값);
  try {
    _일정관리서식적용_(일정시트, true);
  } catch (e) {
    Logger.log('일정관리 표 갱신 실패: ' + e.message);
  }
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
  const 신청열 = 헤더.indexOf('신청일') + 1;
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
  // 보완요청일이 신청일의 다음 해 이후인 경우 신청일만 보면 해당 공휴일이 누락될 수 있다.
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
  SpreadsheetApp.getUi().alert(`마감일 갱신 완료: ${갱신건수}건\n기본: 신청일 + 15 WD\n보완: 보완요청일 + 30 WD\n(주말·공휴일 제외)`);
}
