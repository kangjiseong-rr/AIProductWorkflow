/**
 * ============================================================
 *  심사원 배정·Google Chat 알림 — 라운드로빈 배분, 배정 확정/재발송
 * ============================================================
 *
 *  심사원관리 시트를 기준으로 신규 건 배분(_담당자배분, SheetWriter.js가 호출)과,
 *  "선택 행 배정 확정·알림 발송" 등 메뉴로 실행되는 독립 기능(라운드로빈 추천,
 *  Chat 알림 발송/재발송, Chat 사용자 ID 조회)을 함께 둡니다.
 * ============================================================
 */

/** 심사원관리 시트 기본값·표시 형식 설정 */
function _심사원관리설정_(ss) {
  const 시트 = ss.getSheetByName(SHEET.심사원관리);
  if (!시트) return;

  시트.setFrozenRows(1);
  시트.setColumnWidth(1, 55);
  시트.setColumnWidth(2, 110);
  시트.setColumnWidth(3, 210);
  시트.setColumnWidth(4, 85);
  시트.setColumnWidth(5, 85);
  시트.setColumnWidth(6, 170);
  시트.getRange(2, 4, Math.max(1, 시트.getMaxRows() - 1), 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Y', 'N'], true).setAllowInvalid(false).build());

  // 번호는 배정순서와 별개인 표시용 일련번호이며 현재 행 순서대로 자동 갱신한다.
  if (시트.getLastRow() >= 2) {
    const 행수 = 시트.getLastRow() - 1;
    시트.getRange(2, 1, 행수, 1).setValues(Array.from({ length: 행수 }, (_, i) => [i + 1]));
  }

  const 로그시트 = ss.getSheetByName(SHEET.배정알림로그);
  if (로그시트) {
    로그시트.setFrozenRows(1);
    [145, 120, 110, 210, 100, 210].forEach((너비, i) => 로그시트.setColumnWidth(i + 1, 너비));
  }
}

/** 심사원관리 시트의 활성 심사원을 배정순서대로 반환 */
function _활성심사원목록_(ss) {
  const 시트 = ss.getSheetByName(SHEET.심사원관리);
  if (!시트 || 시트.getLastRow() < 2) return [];

  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 이름열 = 헤더.indexOf('심사원명');
  const 이메일열 = 헤더.indexOf('이메일');
  const 활성열 = 헤더.indexOf('활성여부');
  const 순서열 = 헤더.indexOf('배정순서');
  const ChatID열 = 헤더.indexOf('Chat사용자ID');
  if ([이름열, 이메일열, 활성열, 순서열, ChatID열].some(i => i < 0)) return [];
  const 값 = 시트.getRange(2, 1, 시트.getLastRow() - 1, 헤더.length).getValues();
  return 값.map((행, i) => ({
    심사원명: String(행[이름열] || '').trim(),
    이메일: String(행[이메일열] || '').trim(),
    활성여부: String(행[활성열] || '').trim().toUpperCase(),
    배정순서: Number(행[순서열]) || (i + 1),
    Chat사용자ID: String(행[ChatID열] || '').trim(),
  }))
    .filter(심사원 => 심사원.심사원명 && 심사원.활성여부 === 'Y')
    .sort((a, b) => a.배정순서 - b.배정순서 || a.심사원명.localeCompare(b.심사원명));
}

/** 심사원 이메일을 Workspace Directory에서 조회해 Chat 멘션용 사용자 ID를 일괄 입력 */
function 심사원Chat사용자ID일괄갱신() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getSheetByName(SHEET.심사원관리);
  if (!시트 || 시트.getLastRow() < 2) throw new Error('심사원관리 시트에 심사원을 먼저 등록하세요.');

  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 이름열 = 헤더.indexOf('심사원명');
  const 이메일열 = 헤더.indexOf('이메일');
  const ID열 = 헤더.indexOf('Chat사용자ID');
  if (이름열 < 0 || 이메일열 < 0 || ID열 < 0) throw new Error('심사원관리 시트의 표준 헤더를 확인하세요.');

  const 행수 = 시트.getLastRow() - 1;
  const 값 = 시트.getRange(2, 1, 행수, 헤더.length).getValues();
  const 기존ID = 값.map(행 => [행[ID열]]);
  const 실패 = [];
  let 성공 = 0;
  let 건너뜀 = 0;

  값.forEach((행, i) => {
    const 이름 = String(행[이름열] || '').trim();
    const 이메일 = String(행[이메일열] || '').trim();
    if (!이름 && !이메일) { 건너뜀++; return; }
    if (!이메일) { 실패.push(`${이름 || (i + 2) + '행'}: 이메일 없음`); return; }
    try {
      // 같은 Workspace 조직의 일반 사용자도 조회할 수 있는 공개 프로필만 요청한다.
      // 조직의 연락처 공유가 켜져 있어야 하며 관리자 전용 Directory 권한은 필요하지 않다.
      const 사용자 = AdminDirectory.Users.get(이메일, {
        viewType: 'domain_public',
        projection: 'basic',
      });
      if (!사용자 || !사용자.id) throw new Error('사용자 ID 없음');
      기존ID[i][0] = String(사용자.id);
      성공++;
    } catch (e) {
      실패.push(`${이름 || 이메일}: ${e.message || e}`);
    }
  });

  시트.getRange(2, ID열 + 1, 행수, 1).setValues(기존ID);
  const 결과문 = [`Chat 사용자 ID 갱신 완료: ${성공}건`];
  if (건너뜀) 결과문.push(`빈 행 건너뜀: ${건너뜀}건`);
  if (실패.length) 결과문.push(`실패: ${실패.length}건`, '', 실패.slice(0, 10).join('\n'));
  SpreadsheetApp.getUi().alert(결과문.join('\n'));
}


function _담당자배분(ss) {
  const 심사원목록 = _활성심사원목록_(ss);
  if (!심사원목록.length) throw new Error('심사원관리 시트에 활성 심사원이 없습니다.');
  const 시트 = ss.getSheetByName(SHEET.접수대장);
  const 마지막행 = 시트.getLastRow() - 1; // 헤더 제외
  const idx = 마지막행 % 심사원목록.length;
  return 심사원목록[idx].심사원명;
}

// ─────────────────────────────────────────────
// 심사원 배정 확정 및 Google Chat 알림
// ─────────────────────────────────────────────

/** 일정관리에서 선택한 데이터 행을 헤더 기반 객체로 반환 */
function _선택일정행목록_(ss) {
  const 시트 = ss.getActiveSheet();
  if (시트.getName() !== SHEET.일정관리) throw new Error('일정관리 시트에서 데이터 행을 선택하세요.');
  const 선택 = 시트.getActiveRange();
  const 시작행 = Math.max(2, 선택.getRow());
  const 끝행 = 선택.getLastRow();
  if (끝행 < 2) throw new Error('헤더가 아닌 데이터 행을 선택하세요.');
  if (끝행 - 시작행 + 1 > 50) throw new Error('한 번에 최대 50개 행까지 발송할 수 있습니다.');

  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 값 = 시트.getRange(시작행, 1, 끝행 - 시작행 + 1, 헤더.length).getDisplayValues();
  return 값.map((행값, i) => {
    const 건 = { _행번호: 시작행 + i };
    헤더.forEach((h, j) => { 건[h] = 행값[j]; });
    return 건;
  }).filter(건 => String(건['접수번호'] || '').trim());
}

/** 선택 행 중 담당심사원이 비어 있는 행에만 라운드로빈 초안을 입력 */
function 선택행라운드로빈추천() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getActiveSheet();
  const 건목록 = _선택일정행목록_(ss);
  const 심사원목록 = _활성심사원목록_(ss);
  if (!심사원목록.length) throw new Error('심사원관리 시트에 활성 심사원이 없습니다.');
  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 담당열 = 헤더.indexOf('담당심사원') + 1;
  let 입력수 = 0;
  const 시작순서 = Math.max(0, ss.getSheetByName(SHEET.접수대장).getLastRow() - 1);
  건목록.forEach(건 => {
    if (String(건['담당심사원'] || '').trim()) return;
    const 심사원 = 심사원목록[(시작순서 + 입력수) % 심사원목록.length];
    시트.getRange(건._행번호, 담당열).setValue(심사원.심사원명);
    입력수++;
  });
  SpreadsheetApp.getUi().alert(`라운드로빈 추천 완료: ${입력수}건\n기존 담당심사원 값은 변경하지 않았습니다.`);
}

function _성공발송키목록_(ss) {
  const 시트 = ss.getSheetByName(SHEET.배정알림로그);
  if (!시트 || 시트.getLastRow() < 2) return new Set();
  return new Set(
    시트.getRange(2, 1, 시트.getLastRow() - 1, 6).getValues()
      .filter(행 => String(행[4]).trim() === '발송완료')
      .map(행 => `${String(행[1]).trim()}|${String(행[2]).trim()}`)
  );
}

function _Chat멘션_(사용자ID, 심사원명) {
  const id = String(사용자ID || '').trim();
  if (!id) return `*${심사원명}*`;
  return `<${id.indexOf('users/') === 0 ? id : 'users/' + id}>`;
}

/** 선택 행을 담당자별로 묶어 공용 Chat에 한 번 발송 */
function _선택행배정알림발송_(재발송) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const 건목록 = _선택일정행목록_(ss);
  if (!건목록.length) throw new Error('발송할 접수 건이 없습니다.');
  if (!CONFIG.챗_공통Webhook) throw new Error('Script Properties의 CHAT_COMMON_WEBHOOK을 먼저 설정하세요.');

  const 심사원맵 = {};
  _활성심사원목록_(ss).forEach(심사원 => { 심사원맵[심사원.심사원명] = 심사원; });
  const 누락 = 건목록.filter(건 => !심사원맵[String(건['담당심사원'] || '').trim()]);
  if (누락.length) {
    throw new Error('활성 심사원이 지정되지 않은 행이 있습니다: ' + 누락.map(건 => 건['접수번호']).join(', '));
  }
  const 이메일누락 = Array.from(new Set(건목록.map(건 => String(건['담당심사원']).trim())))
    .filter(이름 => !심사원맵[이름].이메일);
  if (이메일누락.length) throw new Error('심사원관리 시트에 이메일을 입력하세요: ' + 이메일누락.join(', '));

  if (!재발송) {
    const 성공키 = _성공발송키목록_(ss);
    const 중복 = 건목록.filter(건 => 성공키.has(`${건['접수번호']}|${건['담당심사원']}`));
    if (중복.length) {
      throw new Error('이미 발송된 건이 있습니다: ' + 중복.map(건 => 건['접수번호']).join(', ') + '\n재발송 메뉴를 사용하세요.');
    }
  }

  const 그룹 = {};
  건목록.forEach(건 => {
    const 이름 = String(건['담당심사원']).trim();
    if (!그룹[이름]) 그룹[이름] = [];
    그룹[이름].push(건);
  });
  const 요약 = Object.keys(그룹).map(이름 => `${이름} ${그룹[이름].length}건`).join(', ');
  if (ui.alert('배정 알림 발송 확인', `${요약}\n\n선택한 ${건목록.length}건을 공용 Google Chat 방에 발송할까요?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const 본문 = ['📨 *AI 기술심사 배정 확정*'];
  Object.keys(그룹).forEach(이름 => {
    const 심사원 = 심사원맵[이름];
    본문.push('', `${_Chat멘션_(심사원.Chat사용자ID, 이름)} — ${그룹[이름].length}건`);
    그룹[이름].forEach(건 => {
      본문.push(`• *${건['접수번호']}* | ${건['기업명'] || '-'} | ${건['제품명'] || '-'} | 마감 ${건['마감예정일'] || '-'}`);
    });
  });
  본문.push('', `<${ss.getUrl()}|일정관리 바로가기>`);

  const 결과 = _챗POST(CONFIG.챗_공통Webhook, { text: 본문.join('\n') });
  const 실행자 = (() => { try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; } })();
  const 발송일시 = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const 로그값 = 건목록.map(건 => {
    const 이름 = String(건['담당심사원']).trim();
    const 심사원 = 심사원맵[이름];
    return [발송일시, 건['접수번호'], 이름, 심사원.이메일, 결과.success ? '발송완료' : '발송실패', 실행자];
  });
  const 로그시트 = ss.getSheetByName(SHEET.배정알림로그);
  로그시트.getRange(로그시트.getLastRow() + 1, 1, 로그값.length, 6).setValues(로그값);
  if (!결과.success) throw new Error(`Google Chat 발송 실패 (${결과.code}): ${결과.message}`);
  ui.alert(`배정 알림 발송 완료: ${건목록.length}건`);
}

function 선택행배정확정알림발송() { _선택행배정알림발송_(false); }
function 선택행배정알림재발송() { _선택행배정알림발송_(true); }

/** Webhook POST 공통 함수 */
function _챗POST(url, payload) {
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    return { success: code >= 200 && code < 300, code: code, message: res.getContentText() };
  } catch (e) {
    Logger.log('Chat 알림 실패: ' + e.message);
    return { success: false, code: 0, message: e.message };
  }
}

/** Webhook URL 설정 안내 및 테스트 발송 */
function 챗알림테스트() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Google Chat Webhook 테스트',
    'Webhook URL을 입력하세요.\n\n' +
    '[URL 발급 방법]\n' +
    '1. Google Chat 스페이스 열기\n' +
    '2. 스페이스 이름 클릭 > 앱 및 통합\n' +
    '3. Webhook > Webhook 추가 > URL 복사\n\n' +
    '(CONFIG에 저장된 URL을 테스트하려면 빈 칸으로 OK)',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const url = res.getResponseText().trim() || CONFIG.챗_공통Webhook;
  if (!url) { ui.alert('Webhook URL이 없습니다. CONFIG.챗_공통Webhook을 먼저 설정하세요.'); return; }

  const 결과 = _챗POST(url, {
    text: '✅ *AI 기술심사 관리 시스템* 연결 테스트\n알림이 정상적으로 도착했습니다! 🎉',
  });
  ui.alert(결과.success
    ? '테스트 메시지를 발송했습니다. Google Chat을 확인하세요.'
    : `테스트 발송 실패 (${결과.code}): ${결과.message}`);
}
