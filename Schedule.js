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

const 일정관리_상태목록 = ['대기', '심사중', '보완', '완료(적합)', '종료(부적합)', '종료(취소)'];
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

/**
 * 일정관리 시트의 상태별 행 색칠 조건부서식을 현재 헤더 폭 전체로 (재)적용한다.
 * 안전컬럼갱신()으로 새 컬럼이 오른쪽 끝에 추가된 뒤에도 호출되므로,
 * 매번 현재 마지막 열을 기준으로 다시 계산해 새 컬럼도 빠짐없이 덮는다.
 *
 * 생성 시 현재 헤더에서 열을 찾아 $I2 같은 열 고정·행 상대 참조를 만든다.
 * 각 셀 평가마다 전체 데이터 범위를 INDEX/MATCH로 조회하지 않는다.
 * 열 구성이 변경된 경우 색상 서식 갱신으로 현재 헤더 기준 규칙을 재생성한다.
 */
function _일정관리조건부서식적용_(일정시트) {
  const 일정H = 일정시트.getRange(1, 1, 1, Math.max(1, 일정시트.getLastColumn()))
    .getValues()[0].map(v => String(v).trim());
  const iD마감 = 일정H.indexOf('마감예정일') + 1;
  const iD상태 = 일정H.indexOf('상태') + 1;
  const iD보완요청 = 일정H.indexOf('보완요청일') + 1;
  const iD연장마감 = 일정H.indexOf('연장마감일') + 1;
  if (!iD마감 || !iD상태 || !iD보완요청 || !iD연장마감) {
    const 누락헤더 = [['마감예정일', iD마감], ['상태', iD상태], ['보완요청일', iD보완요청], ['연장마감일', iD연장마감]]
      .filter(([, v]) => !v).map(([n]) => n);
    try {
      SpreadsheetApp.getUi().alert(
        '조건부서식을 다시 적용하지 못했습니다.\n\n' +
        '다음 헤더를 1행에서 찾을 수 없습니다: ' + 누락헤더.join(', ') + '\n\n' +
        '현재 1행 헤더(왼쪽부터): ' + 일정H.join(' | ')
      );
    } catch (e) { /* UI 없는 환경 */ }
    return false;
  }

  // 규칙 우선순위: 위에서부터 먼저 적용됨.
  //   ① 기한 초과(미완료)  → 연빨강   (상태색보다 우선)
  //   ② 완료(적합)         → 연녹색
  //   ③ 종료(부적합)       → 진한 회색
  //   ④ 종료(취소)         → 연한 회색
  //   ⑤ 보완               → 머스터드(짙은 노랑)
  //   ⑥ 심사중             → 연노랑
  //   ⑦ 대기               → 무색 (규칙 없음)
  // 2026-09-30: getMaxRows()(그리드 전체, 1000행 안팎)를 그대로 쓰면 실제
  // 데이터는 100행이 안 되는데도 커스텀 수식(+휴일 시트 참조하는 IFERROR/
  // WORKDAY/INDIRECT) 조건부서식을 수백~천 행에 걸어야 했고, 실측 결과
  // 이 정도 규모에서 수식이 TRUE로 계산돼도 화면에 전혀 반영되지 않았다.
  // 실제 데이터가 있는 행 + 여유분만 덮도록 줄인다. 신규 건 등록마다
  // _일정관리조건부서식적용_이 다시 호출되므로(위 함수 주석 참고) 데이터가
  // 늘어나도 그때그때 범위가 재계산돼 기능 손실은 없다.
  const 마지막행 = Math.max(2, 일정시트.getLastRow() + 20);
  const 행수 = 마지막행 - 1;

  // 상태 드롭다운 목록도 함께 최신화한다. 초기설정 이후 새 상태값이
  // 추가돼도(예: 종료(취소)) 이 함수가 재실행될 때 기존 행에 반영된다.
  일정시트.getRange(2, iD상태, 행수, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(일정관리_상태목록, true).build());

  const 행셀 = 열 => `$${columnLetter(열)}2`;
  const 마감 = 행셀(iD마감);
  const 상태 = 행셀(iD상태);
  const 보완요청 = 행셀(iD보완요청);
  const 연장마감 = 행셀(iD연장마감);

  // 특이사항은 사용자가 자유롭게 글자색·굵기를 지정하는 메모 컬럼이므로,
  // 상태별 행 색칠 범위에서 제외해 수동 서식이 항상 우선하도록 한다.
  const iD특이사항 = 일정H.indexOf('특이사항') + 1;
  const 전체범위 = [];
  if (iD특이사항 > 0) {
    if (iD특이사항 > 1) 전체범위.push(일정시트.getRange(2, 1, 행수, iD특이사항 - 1));
    if (iD특이사항 < 일정H.length) 전체범위.push(일정시트.getRange(2, iD특이사항 + 1, 행수, 일정H.length - iD특이사항));
  } else {
    전체범위.push(일정시트.getRange(2, 1, 행수, 일정H.length));
  }

  // 2026-09-30: 실측 결과 "여러 범위(A2:M / O2:AQ) + 커스텀 수식" 조합의
  // 규칙 하나는 수식이 TRUE로 계산돼도 화면에 반영되지 않았고, 범위 1개짜리
  // 단순 규칙은 정상 렌더링됐다. 그래서 범위 2개짜리 규칙 하나 대신, 같은
  // 조건·서식으로 "범위 1개짜리 규칙"을 전체범위 개수만큼(1~2개) 만든다.
  // 특이사항 컬럼을 제외하는 효과(범위 분리)는 그대로 유지된다.
  const 단일범위규칙들 = (formula, bg, fontColor) => 전체범위.map(범위 =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(bg).setFontColor(fontColor)
      .setRanges([범위]).build()
  );

  // ① 기한 초과 & 미완료
  const 초과 = 단일범위규칙들(
    `=AND(NOT(OR(${상태}="완료",${상태}="완료(적합)",${상태}="종료(부적합)",${상태}="종료(취소)")),IF(${보완요청}<>"",AND(${연장마감}<>"",${연장마감}<TODAY()),AND(${마감}<>"",${마감}<TODAY())))`,
    '#F4CCCC', '#990000'
  );

  // ② 완료(적합) → 연녹색 (구버전 '완료' 값도 호환)
  const 완료 = 단일범위규칙들(`=OR(${상태}="완료(적합)",${상태}="완료")`, '#D9EAD3', '#38761D');

  // ③ 종료(부적합) → 진한 회색
  const 종료 = 단일범위규칙들(`=${상태}="종료(부적합)"`, '#666666', '#FFFFFF');

  // ④ 종료(취소) → 연한 회색 (종료(부적합)과 구분되는 밝은 톤)
  const 취소 = 단일범위규칙들(`=${상태}="종료(취소)"`, '#D9D9D9', '#666666');

  // ⑤ 보완 → 머스터드(짙은 노랑)
  const 보완 = 단일범위규칙들(`=${상태}="보완"`, '#F9CB9C', '#783F04');

  // ⑥ 심사중 → 연노랑
  const 심사중 = 단일범위규칙들(`=${상태}="심사중"`, '#FCE8B2', '#7F6000');

  // 2026-09-30: 원래 이 규칙이 조건부서식 안에서 INDIRECT+WORKDAY를 직접
  // 계산했는데(다른 시트=공휴일 참조 때문에 INDIRECT가 필요했다), 실측 결과
  // 그게 조건부서식 전체 렌더링을 불안정하게 만들었다("고치면 잠깐 보였다가
  // 다시 하얗게" 반복). 그래서 "3영업일 전" 날짜를 일반 수식(다른 시트를
  // INDIRECT 없이 직접 참조 가능)으로 숨김 보조 컬럼(D3기준일_*)에 미리
  // 계산해두고, 조건부서식은 그 값을 TODAY()와 단순 비교만 하도록 가볍게
  // 바꿨다. 보조 컬럼이 아직 없는 시트(안전컬럼갱신 실행 전)에서는 해당
  // 임박 규칙만 조용히 건너뛴다.
  // ①초과 규칙과 동일한 우선순위 규칙을 따른다: 보완요청일이 있으면(연장
  // 마감일로 넘어간 건) 연장마감일만, 없으면 마감예정일만 D-3 대상으로 본다.
  // (이전 버전은 두 컬럼을 서로 무관하게 각각 체크해, 연장마감일이 생긴
  // 뒤에도 이미 의미 없어진 원래 마감예정일 기준으로 같이 강조되는 문제가 있었다.)
  const iD_D3마감 = 일정H.indexOf('D3기준일_마감') + 1;
  const iD_D3연장 = 일정H.indexOf('D3기준일_연장마감') + 1;
  const 임박대상 = [
    { 마감열: iD마감, D3열: iD_D3마감, 활성조건: `${보완요청}=""` },
    { 마감열: iD연장마감, D3열: iD_D3연장, 활성조건: `${보완요청}<>""` },
  ].filter(x => x.D3열 > 0);
  const 임박규칙 = 임박대상.map(({ 마감열, D3열, 활성조건 }) => {
    const 마감셀 = 행셀(마감열);
    const D3셀 = 행셀(D3열);
    // 신규 행의 보조 날짜가 비어 있거나 0인 동안 임박으로 잘못 판정하지 않는다.
    const 임박조건 = `IFERROR(AND(${활성조건},ISNUMBER(${마감셀}),ISNUMBER(${D3셀}),${D3셀}>0,NOT(OR(${상태}="완료",${상태}="완료(적합)",${상태}="종료(부적합)",${상태}="종료(취소)")),TODAY()>=${D3셀},TODAY()<=${마감셀}),FALSE)`;
    // 마감 셀의 노란 배경을 상태별 행 색상보다 우선 적용한다.
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=${임박조건}`)
      .setBackground('#FFFF00').setFontColor('#800020')
      .setRanges([일정시트.getRange(2, 마감열, 행수, 1)]).build();
  });

  일정시트.setConditionalFormatRules(임박규칙.concat(초과, 완료, 종료, 취소, 보완, 심사중));
  return true;
}

function _일정관리서식적용_(시트, 요약뷰) {
  const lastCol = Math.max(1, 시트.getLastColumn());
  const 헤더 = 시트.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

  // 일정관리 날짜는 실제 날짜값을 유지하고 화면에는 두 자리 연도로 간결하게 표시.
  // 구글 표의 열에 타입이 지정된 상태(예: DATE)에서는 setNumberFormat이 예외를
  // 던지므로, 서식 실패가 이후 처리(조건부서식 등)까지 막지 않도록 감싼다.
  // 열마다 flush()하면 이 함수가 행 추가마다 호출되는 만큼 왕복이 누적되므로,
  // 날짜 컬럼 전체를 큐에 올린 뒤 flush()는 한 번만 부른다.
  const 날짜열목록 = ['접수일자', '심사접수일', '마감예정일', '보완요청일', '연장마감일', '적합통보일']
    .map(날짜헤더 => ({ 헤더: 날짜헤더, 열: 헤더.indexOf(날짜헤더) + 1 }))
    .filter(x => x.열 > 0);
  if (날짜열목록.length) {
    try {
      날짜열목록.forEach(({ 열 }) =>
        시트.getRange(2, 열, Math.max(1, 시트.getMaxRows() - 1), 1).setNumberFormat('yy-mm-dd')
      );
      SpreadsheetApp.flush();
    } catch (e) {
      const 헤더목록 = 날짜열목록.map(x => x.헤더).join(', ');
      Logger.log(`일정관리 "${헤더목록}" 열 날짜 서식 설정 실패(표 열 타입 충돌 가능): ${e.message}`);
    }
  }

  시트.setHiddenGridlines(false);
  시트.setFrozenRows(1);
  시트.setFrozenColumns(Math.min(2, lastCol));

  시트.setRowHeight(1, 34);
  시트.getRange(1, 1, Math.max(1, 시트.getMaxRows()), lastCol)
    .setFontSize(9)
    .setVerticalAlignment('middle');
  시트.getRange(2, 1, Math.max(1, 시트.getMaxRows() - 1), lastCol)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // 2026-09-30: 이 시트를 구글 시트 "표(Table)"로 감싸면 표 자체 스타일이
  // 헤더 배경색(_일정관리헤더색적용_)과 조건부서식(상태별 행 색상·마감 임박
  // 강조) 위에 그려져서 둘 다 화면에 아예 안 보이는 현상이 실제 확인됐다
  // (새로고침 직후 스크린샷에서 헤더도 남색이 아니고 상태색도 전부 안 보임).
  // 표 갱신 API 호출 자체를 멈춰 더 이상 이 범위를 표로 재감싸지 않는다.
  // 기존에 이미 만들어진 "일정관리_표"는 시트에서 표 셀 클릭 → 표 이름
  // 옆 드롭다운 → "범위로 변환"으로 한 번 수동 해제해야 색이 돌아온다.
  // (관련 함수/컬럼 타입 매핑은 나중에 표를 다시 쓰게 되면 참고할 수 있게 남겨둠)
  // try {
  //   _일정관리구글표적용_(시트, 헤더);
  // } catch (e) {
  //   Logger.log('일정관리 Google Sheets 표 적용 실패: ' + e.message);
  // }

  _일정관리헤더색적용_(시트);

  const 너비맵 = {
    '순번': 45, '접수번호': 115,
    '접수일자': 90, '심사접수일': 95, '마감예정일': 95,
    '상태': 75, '보완요청일': 95, '연장마감일': 95, '적합통보일': 95, '담당심사원': 95, '특이사항': 260,
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
  // 조건부서식 계산용 숨김 보조 컬럼은 요약뷰 여부와 무관하게 항상 숨긴다.
  ['D3기준일_마감', 'D3기준일_연장마감'].forEach(h => {
    const idx = 헤더.indexOf(h);
    if (idx >= 0) 시트.hideColumns(idx + 1);
  });
  if (요약뷰) {
    const 표시컬럼 = new Set([
      '순번', '접수번호',
      '접수일자', '심사접수일', '마감예정일',
      '상태', '보완요청일', '연장마감일', '적합통보일', '담당심사원', '특이사항',
      '기업명', '담당자명', '담당자직급', '담당자전화', '담당자휴대전화', '소재지',
      '제품명', '제품수', '제공형태', '제품분류',
      '인공지능기능수',
    ]);
    헤더.forEach((h, idx) => {
      if (h && !표시컬럼.has(h)) 시트.hideColumns(idx + 1);
    });
  }
}

/**
 * 표가 없거나 헤더 구성(이름·타입)이 바뀐 경우에만 삭제 후 재생성한다.
 * 건 등록마다 호출되는데, 매번 delete+add로 표를 통째로 다시 만들면
 * 그 범위에 이미 적용돼 있던 조건부서식 배경색 렌더링이 초기화돼 버려서
 * "서식이 적용됐다가 반영이 안 되는" 현상의 원인이 됐다(구글 표 열에 타입을
 * 다시 지정하는 과정이 이미 알려진 대로 setNumberFormat도 실패시킨다).
 * 행이 늘어 범위만 커진 경우(가장 흔한 경우)는 updateTable로 range만
 * 확장해 표를 지웠다 새로 만들지 않는다.
 */
function _일정관리구글표적용_(시트, 헤더) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const sheetId = 시트.getSheetId();
  const tableName = '일정관리_표';
  const lastCol = 헤더.length;
  const lastRow = Math.max(2, 시트.getLastRow());
  const 기존표목록 = _시트표목록조회_(spreadsheetId, sheetId);
  const 기존표 = 기존표목록.find(t => t.name === tableName || (t.range && t.range.sheetId === sheetId));

  const 원하는열속성 = 헤더.map((h, idx) => ({
    columnIndex: idx,
    columnName: h || `Column ${idx + 1}`,
    columnType: _일정관리표컬럼타입_(h),
  }));
  const 원하는범위 = {
    sheetId: sheetId,
    startRowIndex: 0,
    endRowIndex: lastRow,
    startColumnIndex: 0,
    endColumnIndex: lastCol,
  };

  if (기존표) {
    const 기존범위 = 기존표.range || {};
    const 범위동일 = 기존범위.startRowIndex === 0 && 기존범위.startColumnIndex === 0
      && 기존범위.endRowIndex === lastRow && 기존범위.endColumnIndex === lastCol;
    const 열동일 = _표컬럼속성동일_(기존표.columnProperties, 원하는열속성);

    if (범위동일 && 열동일) return; // 이미 최신 상태 — API 호출 없이 기존 서식을 그대로 둔다.

    if (열동일) {
      _sheetsBatchUpdate_(spreadsheetId, [{
        updateTable: {
          table: { tableId: 기존표.tableId, range: 원하는범위 },
          fields: 'range',
        },
      }]);
      return;
    }
  }

  const requests = 기존표목록
    .filter(t => t.name === tableName || (t.range && t.range.sheetId === sheetId))
    .map(t => ({ deleteTable: { tableId: t.tableId } }));

  requests.push({
    addTable: {
      table: {
        name: tableName,
        range: 원하는범위,
        rowsProperties: {
          headerColorStyle: { rgbColor: _hexToRgb_(일정관리_헤더색) },
          firstBandColorStyle: { rgbColor: _hexToRgb_('#ffffff') },
          secondBandColorStyle: { rgbColor: _hexToRgb_('#f8fbfb') },
        },
        columnProperties: 원하는열속성,
      },
    },
  });

  _sheetsBatchUpdate_(spreadsheetId, requests);
}

/** 기존 표의 columnProperties(이름·타입)가 원하는 구성과 완전히 같은지 비교 */
function _표컬럼속성동일_(기존, 원함) {
  const 목록a = (기존 || []).slice().sort((x, y) => x.columnIndex - y.columnIndex);
  const 목록b = 원함 || [];
  if (목록a.length !== 목록b.length) return false;
  return 목록b.every((열, idx) => {
    const 대응 = 목록a[idx] || {};
    return 대응.columnName === 열.columnName && 대응.columnType === 열.columnType;
  });
}

function _일정관리표컬럼타입_(헤더명) {
  if (['순번', '제품수', '인공지능기능수'].indexOf(헤더명) >= 0) return 'DOUBLE';
  if (['접수일자', '심사접수일', '마감예정일', '보완요청일', '연장마감일', '적합통보일'].indexOf(헤더명) >= 0) return 'DATE';
  return 'TEXT';
}

function _시트표목록조회_(spreadsheetId, sheetId) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties(sheetId),tables(tableId,name,range,columnProperties(columnIndex,columnName,columnType)))`;
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
/** 한 번의 다건 등록 동안 일정관리 중복 검사에 재사용할 인덱스를 만든다. */
function _일정관리등록컨텍스트생성(ss) {
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 데이터 = 일정시트.getDataRange().getDisplayValues();
  const 헤더 = (데이터[0] || []).map(v => String(v).trim());
  const 접수번호열 = 헤더.indexOf('접수번호');
  const 접수번호행맵 = new Map();
  if (접수번호열 >= 0) {
    for (let r = 1; r < 데이터.length; r++) {
      const 번호 = String(데이터[r][접수번호열] || '').trim();
      if (번호 && !접수번호행맵.has(번호)) 접수번호행맵.set(번호, r + 1);
    }
  }
  return { 접수번호행맵 };
}

function _일정관리행추가(ss, 접수번호, 직접값, 등록컨텍스트) {
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 컨텍스트 = 등록컨텍스트 || _일정관리등록컨텍스트생성(ss);
  const 정규화접수번호 = String(접수번호).trim();
  if (컨텍스트.접수번호행맵.has(정규화접수번호)) {
    const 기존행번호 = 컨텍스트.접수번호행맵.get(정규화접수번호);
    Logger.log(`일정관리 기존 행 유지 - append 생략: ${접수번호} (${기존행번호}행)`);
    return { 신규: false, 행번호: 기존행번호 };
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

    // 2-2) D3기준일_* = 해당 마감일의 3영업일 전 (조건부서식용 숨김 보조 컬럼)
    if (h === 'D3기준일_마감') {
      const iD마감 = 일정H.indexOf('마감예정일') + 1;
      const 마감셀 = `${columnLetter(iD마감)}${새행번호}`;
      return `=IF(${마감셀}="","",WORKDAY(${마감셀},-3,'${공휴일시트명}'!$A$2:$A))`;
    }
    if (h === 'D3기준일_연장마감') {
      const iD연장 = 일정H.indexOf('연장마감일') + 1;
      const 연장셀 = `${columnLetter(iD연장)}${새행번호}`;
      return `=IF(${연장셀}="","",WORKDAY(${연장셀},-3,'${공휴일시트명}'!$A$2:$A))`;
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
  컨텍스트.접수번호행맵.set(정규화접수번호, 새행번호);
  try {
    _일정관리서식적용_(일정시트, true);
    // 다른 호출부(헤더마이그레이션·안전컬럼갱신·초기설정실행)와 동일하게
    // 표 갱신 직후 조건부서식(상태별 배경색·마감 임박 강조)을 다시 씌운다.
    // 이 호출이 빠져 있으면 새 건이 등록될 때마다(가장 빈번한 경로) 색상
    // 서식이 반영되지 않는 현상이 반복된다.
    _일정관리조건부서식적용_(일정시트);
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
  // 구글 표의 열에 타입(DATE)이 지정된 상태에서는 setNumberFormat이 예외를
  // 던지므로, 수식 입력(핵심 동작)과 분리해 서식 실패가 수식 갱신을 막지 않게 한다.
  시트.getRange(2, 마감열, 행수, 1).setFormulas(수식);
  try {
    시트.getRange(2, 마감열, 행수, 1).setNumberFormat('yy-mm-dd');
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log(`일정관리 "마감예정일" 열 날짜 서식 설정 실패(표 열 타입 충돌 가능): ${e.message}`);
  }
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
      .setFormulas(연장수식);
    try {
      시트.getRange(2, 연장마감열, 행수, 1).setNumberFormat('yy-mm-dd');
      SpreadsheetApp.flush();
    } catch (e) {
      Logger.log(`일정관리 "연장마감일" 열 날짜 서식 설정 실패(표 열 타입 충돌 가능): ${e.message}`);
    }
  }

  // D3기준일_* (조건부서식용 숨김 보조 컬럼) 백필: 마감예정일/연장마감일 수식을
  // 방금 다시 썼으니, 그 값의 3영업일 전 날짜도 함께 갱신한다. 헤더가 아직
  // 없는(안전컬럼갱신 전) 시트에서는 조용히 건너뛴다.
  const D3마감열 = 헤더.indexOf('D3기준일_마감') + 1;
  const D3연장열 = 헤더.indexOf('D3기준일_연장마감') + 1;
  if (D3마감열 > 0) {
    const 마감열문자 = columnLetter(마감열);
    const D3마감수식 = 신청값.map((_, i) => {
      const 행 = i + 2;
      const 마감셀 = `${마감열문자}${행}`;
      return [`=IF(${마감셀}="","",WORKDAY(${마감셀},-3,'${공휴일시트명}'!$A$2:$A))`];
    });
    시트.getRange(2, D3마감열, 행수, 1).setFormulas(D3마감수식);
  }
  if (D3연장열 > 0 && 연장마감열 > 0) {
    const 연장마감열문자 = columnLetter(연장마감열);
    const D3연장수식 = 신청값.map((_, i) => {
      const 행 = i + 2;
      const 연장셀 = `${연장마감열문자}${행}`;
      return [`=IF(${연장셀}="","",WORKDAY(${연장셀},-3,'${공휴일시트명}'!$A$2:$A))`];
    });
    시트.getRange(2, D3연장열, 행수, 1).setFormulas(D3연장수식);
  }
  if (D3마감열 > 0 || D3연장열 > 0) SpreadsheetApp.flush();

  return 행수;
}

/**
 * 기존 일정 행의 값·수식은 건드리지 않고, 계산에 필요한 공휴일 연도만 확보한다.
 * 엑셀 등록 뒤 전체 마감 수식을 다시 쓰지 않으면서 장기 진행 건의 연도 범위를 보완한다.
 */
function _일정관리공휴일연도만확보_(ss) {
  const 시트 = ss.getSheetByName(SHEET.일정관리);
  const 현재연도 = new Date().getFullYear();
  const 연도집합 = new Set([현재연도, 현재연도 + 1]);
  if (시트 && 시트.getLastRow() >= 2) {
    const 데이터 = 시트.getDataRange().getValues();
    const 헤더 = 데이터[0].map(v => String(v).trim());
    const 날짜열들 = ['접수일자', '심사접수일', '보완요청일']
      .map(h => 헤더.indexOf(h)).filter(i => i >= 0);
    for (let r = 1; r < 데이터.length; r++) {
      날짜열들.forEach(c => {
        const 연도 = _날짜연도_(데이터[r][c]);
        if (연도) {
          연도집합.add(연도);
          연도집합.add(연도 + 1);
        }
      });
    }
  }
  _공휴일연도확보_(ss, Array.from(연도집합));
  return 연도집합.size;
}

/**
 * Sheets REST API(403 이슈) 없이 순수 SpreadsheetApp만으로 실제 서버가
 * 조건부서식 수식을 참(TRUE)으로 평가하는지 직접 확인한다.
 * 임시 숨김 시트에 동일한 수식을 그대로 써서 계산시킨 뒤 즉시 지운다.
 * 커서가 있는 행(없으면 2행)을 대상으로 한다.
 */
function 일정관리조건부서식수식테스트() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getSheetByName(SHEET.일정관리);
  const ui = SpreadsheetApp.getUi();
  if (!시트 || 시트.getLastRow() < 2) { ui.alert('테스트할 일정관리 데이터가 없습니다.'); return; }

  const 일정H = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const iD마감 = 일정H.indexOf('마감예정일') + 1;
  const iD상태 = 일정H.indexOf('상태') + 1;
  const iD보완요청 = 일정H.indexOf('보완요청일') + 1;
  const iD연장마감 = 일정H.indexOf('연장마감일') + 1;
  if (!iD마감 || !iD상태 || !iD보완요청 || !iD연장마감) { ui.alert('필수 헤더(마감예정일/상태/보완요청일/연장마감일)를 찾을 수 없습니다.'); return; }

  const 선택행 = ss.getActiveSheet().getSheetId() === 시트.getSheetId()
    ? 시트.getActiveRange().getRow() : 2;
  const 행 = 선택행 >= 2 ? 선택행 : 2;

  const 셀 = 열 => `$${columnLetter(열)}${행}`;
  const 마감 = 셀(iD마감), 상태 = 셀(iD상태), 보완요청 = 셀(iD보완요청), 연장마감 = 셀(iD연장마감);

  const 수식목록 = [
    ['①초과', `AND(NOT(OR(${상태}="완료",${상태}="완료(적합)",${상태}="종료(부적합)",${상태}="종료(취소)")),IF(${보완요청}<>"",AND(${연장마감}<>"",${연장마감}<TODAY()),AND(${마감}<>"",${마감}<TODAY())))`],
    ['②완료(적합)', `OR(${상태}="완료(적합)",${상태}="완료")`],
    ['③종료(부적합)', `${상태}="종료(부적합)"`],
    ['④종료(취소)', `${상태}="종료(취소)"`],
    ['⑤보완', `${상태}="보완"`],
    ['⑥심사중', `${상태}="심사중"`],
  ];

  // ⚠️ 이전 버전은 별도 임시 시트에 수식을 써서 계산했는데, $I2 같은 참조가
  // "그 임시 시트의" I2(빈 칸)를 가리켜 버려 결과가 전부 잘못 나왔다. 그래서
  // 같은 일정관리 시트의 빈 열 1행에 썼더니, 이번엔 이 시트가 아직 구글
  // "표(Table)"로 감싸여 있어 "표 헤더 행에는 수식이 지원되지 않습니다"로
  // 막혔다. 표의 행·열 범위 양쪽 다 확실히 벗어난, 맨 아래쪽 빈 행에 쓴다.
  const 스크래치행 = Math.max(시트.getLastRow(), 시트.getMaxRows()) + 20;
  const 스크래치열 = 일정H.length + 3;
  if (시트.getMaxRows() < 스크래치행 + 수식목록.length) {
    시트.insertRowsAfter(시트.getMaxRows(), 스크래치행 + 수식목록.length - 시트.getMaxRows());
  }
  if (시트.getMaxColumns() < 스크래치열) 시트.insertColumnsAfter(시트.getMaxColumns(), 스크래치열 - 시트.getMaxColumns());
  시트.getRange(스크래치행, 스크래치열, 수식목록.length, 1).setFormulas(수식목록.map(([, f]) => [`=${f}`]));
  SpreadsheetApp.flush();
  const 결과 = 시트.getRange(스크래치행, 스크래치열, 수식목록.length, 1).getDisplayValues();
  시트.getRange(스크래치행, 스크래치열, 수식목록.length, 1).clearContent();

  const 규칙목록 = 시트.getConditionalFormatRules();
  const 규칙범위요약 = 규칙목록.map((r, i) => {
    const 범위들 = r.getRanges().map(rg => rg.getA1Notation()).join(' / ');
    let 실제수식 = '';
    try { 실제수식 = '=' + r.getBooleanCondition().getCriteriaValues()[0]; } catch (e) { 실제수식 = '(수식 아님)'; }
    return `${i + 1}. [${범위들}] ${실제수식}`;
  }).join('\n');

  const 상태원본 = 시트.getRange(행, iD상태).getValue();
  const 마감값 = 시트.getRange(행, iD마감).getDisplayValue();
  const 요약 = 수식목록.map(([k], i) => `${k}: ${결과[i][0]}`).join('\n');

  // 눈으로 똑같아 보여도 코드값이 다른 문자(전각 괄호·공백류)가 섞였을 수 있으므로
  // 상태 컬럼 실제 값과 코드 목록의 가장 비슷한 후보를 문자 코드 단위로 비교한다.
  const 코드 = s => Array.from(String(s)).map(c => c.codePointAt(0)).join(',');
  const 실제문자열 = String(상태원본);
  const 후보 = 일정관리_상태목록.find(v => v === 실제문자열)
    || 일정관리_상태목록.find(v => v.replace(/\s/g, '') === 실제문자열.replace(/\s/g, ''))
    || 일정관리_상태목록[0];
  const 코드비교 = `실제값 길이:${실제문자열.length} 코드:[${코드(실제문자열)}]\n` +
    `"${후보}" 길이:${후보.length} 코드:[${코드(후보)}]\n` +
    `완전일치(===): ${실제문자열 === 후보}`;

  ui.alert(
    `${행}행 조건부서식 수식 실측 (상태 원본 값="${상태원본}", 마감예정일="${마감값}")\n\n` +
    `[수식 계산 결과]\n${요약}\n\n` +
    `[현재 시트에 실제로 붙어있는 규칙 수: ${규칙목록.length}개]\n${규칙범위요약}\n\n` +
    `[상태값 문자 코드 비교]\n${코드비교}`
  );
}

/** 규칙과 서버가 계산한 색상을 읽기만 한다. 갱신·flush·셀 쓰기를 하지 않는다. */
function 일정관리색상진단() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getSheetByName(SHEET.일정관리);
  const ui = SpreadsheetApp.getUi();
  if (!시트 || 시트.getLastRow() < 2) {
    ui.alert('진단할 일정관리 데이터가 없습니다.');
    return;
  }
  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getDisplayValues()[0].map(v => v.trim());
  const 상태열 = 헤더.indexOf('상태') + 1;
  if (!상태열) { ui.alert('상태 헤더를 찾을 수 없습니다.'); return; }
  const 상태값 = 시트.getRange(2, 상태열, 시트.getLastRow() - 1, 1).getDisplayValues();
  const 완료위치 = 상태값.findIndex(r => ['완료(적합)', '완료'].includes(r[0]));
  const 선택행 = ss.getActiveSheet().getSheetId() === 시트.getSheetId()
    ? 시트.getActiveRange().getRow() : 1;
  const 행 = 선택행 > 1 && 선택행 <= 시트.getLastRow() ? 선택행 : (완료위치 >= 0 ? 완료위치 + 2 : 2);
  const 열목록 = ['상태', '마감예정일', '연장마감일'].map(h => 헤더.indexOf(h) + 1).filter(c => c > 0);
  const 범위 = 열목록.map(c => `'${시트.getName().replace(/'/g, "''")}'!${columnLetter(c)}${행}`);
  const fields = 'sheets(properties(sheetId,title),conditionalFormats,tables(tableId,name,range),data(startRow,startColumn,rowData(values(formattedValue,effectiveValue,userEnteredFormat,effectiveFormat))))';
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${ss.getId()}?fields=${encodeURIComponent(fields)}&` +
    범위.map(r => 'ranges=' + encodeURIComponent(r)).join('&');
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    ui.alert(`진단 조회 실패 (${res.getResponseCode()}): ${res.getContentText()}`);
    return;
  }
  const 결과 = {
    진단버전: '2026-09-30', 행, 상태: 상태값[행 - 2][0],
    상태열: columnLetter(상태열), 조회범위: 범위,
    서버응답: JSON.parse(res.getContentText()),
  };
  const 내용 = JSON.stringify(결과, null, 2).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  ui.showModalDialog(HtmlService.createHtmlOutput(
    '<p>서식은 변경하지 않았습니다. 아래 결과를 복사해 전달해주세요.</p>' +
    '<textarea readonly style="width:100%;height:440px;box-sizing:border-box" onclick="this.select()">' + 내용 + '</textarea>'
  ).setWidth(720).setHeight(520), '일정관리 색상 진단');
}

/** 색상만 갱신: 공휴일 조회·날짜 수식 재입력·Google 표 재생성을 생략한다. */
function 일정관리색상서식갱신() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getSheetByName(SHEET.일정관리);
  if (!시트) {
    SpreadsheetApp.getUi().alert('일정관리 시트를 찾을 수 없습니다.');
    return;
  }
  const 시작 = Date.now();
  if (!_일정관리조건부서식적용_(시트)) return;
  SpreadsheetApp.flush();
  const 소요초 = ((Date.now() - 시작) / 1000).toFixed(1);
  Logger.log(`일정관리 색상 서식 갱신: ${소요초}초`);
  ss.toast(`색상 서식 적용 완료 (${소요초}초). 화면 반영에는 시간이 더 걸릴 수 있습니다.`, '일정관리', 8);
}

/** 공휴일·마감일 수식만 갱신한다. 기존 조건부서식 규칙은 재설정하지 않는다. */
function 마감예정일갱신() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 갱신건수 = _마감예정일수식갱신_(ss);
  SpreadsheetApp.getUi().alert(`마감일 갱신 완료: ${갱신건수}건\n기본: 접수일자 + 15 WD\n보완: 보완요청일 + 30 WD\n(주말·공휴일 제외)`);
}
