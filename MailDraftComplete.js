/**
 * ============================================================
 *  메일 초안 생성 — 기술심사 완료 내부 보고 (Gmail 임시보관함)
 * ============================================================
 *  업체가 아니라 내부 보고 라인(임우섭 팀장)에게 "이 건 심사가 끝났고
 *  보고서는 여기 있다"고 보고하는 메일. MailDraft.js(일정 확인 요청)와
 *  달리 완료 여부(종합판정 등)는 검증하지 않는다 — 메뉴를 실행하는 것
 *  자체가 심사원이 이미 완료로 판단했다는 전제다.
 *
 *  보고서 링크는 심사원이 직접 붙여넣는다. 다만 엉뚱한 건의 링크를
 *  잘못 붙여넣는 실수를 잡기 위해, 링크가 가리키는 Drive 문서 제목
 *  맨 앞의 "(접수번호)"를 읽어 선택한 행의 접수번호와 대조한다.
 *  다이얼로그에서 "확인"을 누르는 즉시 검증하면 Drive 호출 때문에
 *  체감 속도가 느려지므로, 검증은 실제 임시보관함 초안을 만드는 단계
 *  (_기술심사완료보고메일초안_생성)에서 조용히 수행한다. 불일치가
 *  감지되면 초안을 만들지 않고 "취소/계속 진행"을 다이얼로그에
 *  되묻고, 계속 진행을 선택한 경우에만(force=true) 그대로 생성한다.
 */

// 받는사람은 조직 보고 라인이 바뀌기 전까지 고정이므로 하드코딩한다.
const 완료보고_고정수신자 = '임우섭 팀장 소프트웨어품질인증팀 <woosup.lim@tta.or.kr>';

/** 보고서링크(Drive/Docs URL)에서 파일 ID를 추출. 못 찾으면 빈 문자열. */
function _보고서링크파일ID_(url) {
  const m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

/** 보고서링크가 가리키는 문서 제목 맨 앞의 "(접수번호)"를 추출. 조회 실패 시 빈 문자열(검증 생략). */
function _보고서링크접수번호_(url) {
  const fileId = _보고서링크파일ID_(url);
  if (!fileId) return '';
  try {
    const name = DriveApp.getFileById(fileId).getName();
    const m = String(name || '').match(/^\(([^)]+)\)/);
    return m ? m[1].trim() : '';
  } catch (e) {
    return '';
  }
}

/** 접수번호로 "기술심사 완료 보고" 메일 자료(수신자는 고정이라 제외하고 제목용 정보)를 구성. */
function _완료보고메일자료_(ss, 접수번호) {
  const 건 = _건조회(ss, 접수번호);
  if (!건) throw new Error(`접수대장에서 접수번호를 찾지 못했습니다: ${접수번호}`);

  const 기업명 = String(건['기업명'] || '').trim();
  const 제품명 = String(건['제품명'] || '').trim();
  const 누락 = [];
  if (!기업명) 누락.push('기업명');
  if (!제품명) 누락.push('제품명');
  if (누락.length > 0) {
    throw new Error(`접수대장에 다음 항목이 비어 있어 메일 초안을 만들 수 없습니다: ${누락.join(', ')}`);
  }

  const 담당심사원 = _실행자심사원명_(ss);
  if (!담당심사원) {
    throw new Error('실행 계정의 이메일을 심사원관리 시트에서 찾지 못했습니다. 심사원관리 시트에 본인 계정(이메일)이 활성 상태로 등록돼 있는지 확인하세요.');
  }

  const 제목 = `(AI확인제)(${접수번호}, ${기업명}, ${제품명}) 기술심사 완료 건`;

  return { 접수번호, 기업명, 제품명, 담당심사원, 제목 };
}

/** 다이얼로그 열기(기술심사완료보고메일초안)와 제출(_기술심사완료보고메일초안_생성) 시점의 중복 조회를 캐시로 줄인다. */
function _완료보고메일자료_캐시조회_(ss, 접수번호) {
  const cache = CacheService.getUserCache();
  const key = `완료보고메일:${접수번호}`;
  const 캐시값 = cache.get(key);
  if (캐시값) {
    try { return JSON.parse(캐시값); } catch (e) {}
  }
  const 자료 = _완료보고메일자료_(ss, 접수번호);
  cache.put(key, JSON.stringify(자료), 180);
  return 자료;
}

function _완료보고본문_(자료, 보고서링크) {
  const 본문 =
`안녕하세요, ${자료.담당심사원}입니다.

${자료.접수번호} 건 기술심사 완료되어 다음 내용 보고드립니다.

${보고서링크}

검토 부탁드립니다.

감사합니다.
${자료.담당심사원} 드림`;

  const 서명HTML = _실행자Gmail서명HTML_();
  const 본문최종 = 서명HTML ? `${본문}\n\n${_html을텍스트로_(서명HTML)}` : 본문;
  const 본문HTML최종 =
    `${_html이스케이프_(본문).replace(/\n/g, '<br>')}` + (서명HTML ? `<br><br>${서명HTML}` : '');
  return { 본문: 본문최종, 본문HTML: 본문HTML최종 };
}

/**
 * 접수대장/일정관리에서 행을 선택한 뒤 실행 →
 * CC 선택 + 보고서링크 입력 다이얼로그(MailDraftCompleteCC.html)를 띄운다.
 * 내부 보고용 메일이라 시트 사용자 누구나 사용 가능.
 */
function 기술심사완료보고메일초안() {
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
    자료 = _완료보고메일자료_캐시조회_(ss, 접수번호);
  } catch (e) {
    ui.alert(e.message || String(e));
    return;
  }

  const template = HtmlService.createTemplateFromFile('MailDraftCompleteCC');
  template.접수번호 = 접수번호;
  template.받는사람 = 완료보고_고정수신자;
  template.제목 = 자료.제목;
  template.심사원목록JSON = JSON.stringify(
    _활성심사원목록_(ss).map(s => ({ 심사원명: s.심사원명, 이메일: s.이메일 }))
  );
  const html = template.evaluate().setWidth(480).setHeight(480);
  ui.showModalDialog(html, `기술심사 완료 보고 메일 — ${접수번호}`);
}

/**
 * MailDraftCompleteCC.html에서 호출 — 보고서링크 검증 후 Gmail 임시보관함 초안 생성.
 * force가 false인 상태에서 링크 문서 제목의 접수번호가 선택한 건과 다르게 확인되면
 * 초안을 만들지 않고 { 불일치: true, 링크접수번호 }를 돌려줘 다이얼로그가 사용자에게
 * 취소/계속 진행을 되묻게 한다. force가 true면(계속 진행 선택) 불일치여도 그대로 생성한다.
 */
function _기술심사완료보고메일초안_생성(접수번호, 보고서링크, 선택CC, force) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 자료 = _완료보고메일자료_캐시조회_(ss, 접수번호);

  const 링크값 = String(보고서링크 || '').trim();
  if (!링크값) throw new Error('보고서 링크를 입력하세요.');

  if (!force) {
    const 링크접수번호 = _보고서링크접수번호_(링크값);
    if (링크접수번호 && 링크접수번호 !== 접수번호) {
      return { 불일치: true, 링크접수번호 };
    }
  }

  const { 본문, 본문HTML } = _완료보고본문_(자료, 링크값);
  const options = { htmlBody: 본문HTML };
  const cc = (Array.isArray(선택CC) ? 선택CC : []).map(v => String(v || '').trim()).filter(Boolean);
  if (cc.length > 0) options.cc = cc.join(',');

  GmailApp.createDraft(완료보고_고정수신자, 자료.제목, 본문, options);
  return { 불일치: false };
}
