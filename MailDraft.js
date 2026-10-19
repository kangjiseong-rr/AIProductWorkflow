/**
 * ============================================================
 *  메일 초안 생성 — Gmail 임시보관함
 * ============================================================
 *  기존 "업체 메일 작성 도우미"(Python 데스크톱 도구)는 Gmail API로
 *  서명(sendAs)을 직접 조회하는데, 계정에 서명이 여러 개 등록된 경우
 *  해당 조회가 실패해 메일을 만들지 못하는 문제가 있었다.
 *  GmailApp(Apps Script 내장 서비스)만으로는 임시보관함 초안에 서명이
 *  자동으로 붙지 않는다. 대신 Gmail 고급 서비스(appsscript.json의
 *  enabledAdvancedServices — Gmail v1)로 실행 계정 본인의 주소를 정확히
 *  지정해 Gmail.Users.Settings.SendAs.get('me', 그이메일)을 호출한다.
 *  목록 전체를 가져와 추측하지 않고 주소를 정확히 지정해서 조회하므로,
 *  계정에 서명(=sendAs 별칭)이 여러 개 등록돼 있어도 깨지지 않는다.
 *
 *  CC는 고정값이 아니라 실행 시 심사원관리 시트의 활성 심사원 목록에서
 *  체크박스로 직접 선택한다 (CC 대상 본인이 이 메뉴를 실행할 수도 있어서
 *  항상 특정 인원을 끼워 넣는 것은 맞지 않기 때문).
 */

/**
 * 실행 계정의 이메일로 심사원관리 시트를 조회해 심사원명을 반환.
 * 메일 서명은 그 건의 배정된 담당심사원이 아니라, 실제로 메일을 보내는(실행하는) 사람 기준.
 */
function _실행자심사원명_(ss) {
  let 이메일 = '';
  try { 이메일 = (Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) {}
  if (!이메일) return '';
  const 심사원 = _활성심사원목록_(ss).find(s => s.이메일.toLowerCase() === 이메일);
  return 심사원 ? 심사원.심사원명 : '';
}

/** 실행 계정의 Gmail 서명(HTML)을 조회. 서명이 없거나 조회 실패 시 빈 문자열(메일 생성 자체는 막지 않음). */
function _실행자Gmail서명HTML_() {
  let 이메일 = '';
  try { 이메일 = (Session.getActiveUser().getEmail() || '').trim(); } catch (e) {}
  if (!이메일) return '';
  try {
    const sendAs = Gmail.Users.Settings.SendAs.get('me', 이메일);
    return (sendAs && sendAs.signature) ? sendAs.signature : '';
  } catch (e) {
    return '';
  }
}

function _html이스케이프_(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 서명 HTML을 대략적인 평문으로 변환 (plain-text 본문용 fallback). */
function _html을텍스트로_(html) {
  return String(html || '')
    .replace(/<\s*br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .trim();
}

/** 접수번호로 "기술심사 일정 확인 요청" 메일 자료(수신자·제목·본문)를 구성. 누락 시 Error throw. */
function _일정확인요청메일자료_(ss, 접수번호) {
  const 건 = _건조회(ss, 접수번호);
  if (!건) throw new Error(`접수대장에서 접수번호를 찾지 못했습니다: ${접수번호}`);

  const 이메일 = String(건['이메일'] || '').trim();
  const 기업명 = String(건['기업명'] || '').trim();
  const 제품명 = String(건['제품명'] || '').trim();

  const 누락 = [];
  if (!이메일) 누락.push('이메일');
  if (!기업명) 누락.push('기업명');
  if (!제품명) 누락.push('제품명');
  if (누락.length > 0) {
    throw new Error(`접수대장에 다음 항목이 비어 있어 메일 초안을 만들 수 없습니다: ${누락.join(', ')}`);
  }

  const 담당심사원 = _실행자심사원명_(ss);
  if (!담당심사원) {
    throw new Error('실행 계정의 이메일을 심사원관리 시트에서 찾지 못했습니다. 심사원관리 시트에 본인 계정(이메일)이 활성 상태로 등록돼 있는지 확인하세요.');
  }

  const 제목 = `[TTA](AI 확인제 - ${기업명}) 기술심사 일정 확인 요청 건`;
  const 본문 =
`안녕하세요. 한국정보통신기술협회 ${담당심사원}입니다.

신청하신 AI확인제의 기술심사를 위해, 직접 방문(가능한 경우 온라인 가능) 등을 통해 해당 확인을 진행하고자 합니다.

* 대상제품: ${제품명}

기술심사는 다음 내용으로 확인 예정입니다.
 - 제품 확인 및 시연
 - AI모델 적용 여부 확인
 - AI적용 기능 확인(시연) 등

위 각 확인 내용별 증적자료 제출이 요구되오니 사전 검토 부탁드립니다.
* 증적자료는 현장에서 확인한 AI모델 증빙, AI기능 확인, (필요시)환경 및 연산자원 등 입니다.
** 증적자료 예) 캡쳐 이미지, 영상, 샘플 데이터 등

가능하신 일정의 확인 및 회신 요청드립니다.
* 예) 0/00 오전, 0/00 ~ 0/00 등

⭐ 가이드라인 참고: https://aiverify.kr/bbspost/apply-guide/detail/2Y

위 내용 및 기타 궁금하신 사항은 언제든지 편하게 연락주시기 바랍니다.

감사합니다.
${담당심사원} 드림`;

  const 서명HTML = _실행자Gmail서명HTML_();
  const 본문최종 = 서명HTML ? `${본문}\n\n${_html을텍스트로_(서명HTML)}` : 본문;
  const 본문HTML최종 =
    `${_html이스케이프_(본문).replace(/\n/g, '<br>')}` + (서명HTML ? `<br><br>${서명HTML}` : '');

  return { 접수번호, 이메일, 기업명, 제품명, 담당심사원, 제목, 본문: 본문최종, 본문HTML: 본문HTML최종 };
}

/**
 * 다이얼로그를 여는 시점(기술심사일정확인요청메일초안)과 제출 시점(_기술심사일정확인요청메일초안_생성)에
 * 똑같은 접수대장·심사원관리 조회를 두 번 반복하지 않도록 사용자별 캐시에 잠깐(3분) 담아 재사용한다.
 * (접수번호를 바꿔 치기해도 서버가 그 접수번호 기준으로 다시 조회·검증하므로 안전함 — 캐시는 순전히 속도용)
 */
function _일정확인요청메일자료_캐시조회_(ss, 접수번호) {
  const cache = CacheService.getUserCache();
  const key = `일정확인요청메일:${접수번호}`;
  const 캐시값 = cache.get(key);
  if (캐시값) {
    try { return JSON.parse(캐시값); } catch (e) {}
  }
  const 자료 = _일정확인요청메일자료_(ss, 접수번호);
  cache.put(key, JSON.stringify(자료), 180);
  return 자료;
}

/**
 * 접수대장/일정관리에서 행을 선택한 뒤 실행 →
 * CC 선택용 다이얼로그(MailDraftCC.html)를 띄운다. 시트 사용자 누구나 사용 가능.
 */
function 기술심사일정확인요청메일초안() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 활성시트 = ss.getActiveSheet();

  if (활성시트.getName() !== SHEET.접수대장 && 활성시트.getName() !== SHEET.일정관리) {
    ui.alert('접수대장 또는 일정관리 시트에서 대상 건의 행을 클릭(선택)한 뒤 다시 실행하세요.');
    return;
  }
  const 활성행 = ss.getActiveRange().getRow();
  if (활성행 < 2) {
    ui.alert('헤더가 아니라 데이터 행을 선택한 뒤 다시 실행하세요.');
    return;
  }

  const 헤더 = 활성시트.getRange(1, 1, 1, 활성시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 접수번호열 = 헤더.indexOf('접수번호') + 1;
  if (접수번호열 < 1) {
    ui.alert('이 시트에서 접수번호 컬럼을 찾지 못했습니다.');
    return;
  }
  const 접수번호 = String(활성시트.getRange(활성행, 접수번호열).getValue() || '').trim();
  if (!접수번호) {
    ui.alert('선택한 행에 접수번호가 없습니다.');
    return;
  }

  let 자료;
  try {
    자료 = _일정확인요청메일자료_캐시조회_(ss, 접수번호);
  } catch (e) {
    ui.alert(e.message || String(e));
    return;
  }

  const template = HtmlService.createTemplateFromFile('MailDraftCC');
  template.접수번호 = 접수번호;
  template.받는사람 = 자료.이메일;
  template.제목 = 자료.제목;
  template.심사원목록JSON = JSON.stringify(
    _활성심사원목록_(ss).map(s => ({ 심사원명: s.심사원명, 이메일: s.이메일 }))
  );
  const html = template.evaluate().setWidth(480).setHeight(420);
  ui.showModalDialog(html, `기술심사 일정 확인 요청 메일 — ${접수번호}`);
}

/** MailDraftCC.html에서 호출 — 선택된 CC로 실제 Gmail 임시보관함 초안 생성. */
function _기술심사일정확인요청메일초안_생성(접수번호, 선택CC) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 자료 = _일정확인요청메일자료_캐시조회_(ss, 접수번호);

  const options = { htmlBody: 자료.본문HTML };
  const cc = (Array.isArray(선택CC) ? 선택CC : []).map(v => String(v || '').trim()).filter(Boolean);
  if (cc.length > 0) options.cc = cc.join(',');

  GmailApp.createDraft(자료.이메일, 자료.제목, 자료.본문, options);
  return 'Gmail 임시보관함에 초안이 생성되었습니다. 내용을 확인한 뒤 직접 전송해 주세요.';
}
