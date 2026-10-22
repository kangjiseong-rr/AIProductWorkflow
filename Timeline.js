/**
 * Timeline.js — 관리자 전용 "심사 일정 타임라인"·"심사 현황 대시보드" (웹 앱, 새 탭)
 *
 * 일정관리 시트를 헤더명 기준으로 읽어 두 화면(TimelineView.html, DashboardView.html)이
 * 같은 자료를 쓰도록 한다. 경과·잔여·처리기간은 마감 계산(WORKDAY)과 같은 기준인
 * 영업일(WD: 주말·공휴일 시트 날짜 제외)로 서버에서 한 번만 계산한다.
 * 시트는 읽기만 하고 서식·수식·공휴일 동기화(네트워크 조회)는 건드리지 않는다.
 */

const 타임라인_컬럼 = {
  seq: '순번',
  id: '접수번호',
  received: '접수일자',
  review: '심사접수일',
  original: '마감예정일',
  status: '상태',
  supplement: '보완요청일',
  extension: '연장마감일',
  completion: '심사완료일',
  notice: '적합통보일',
  reviewer: '담당심사원',
  note: '특이사항',
  company: '기업명',
  product: '제품명',
  count: '제품수',
};
const 타임라인_날짜키 = ['received', 'review', 'original', 'supplement', 'extension', 'completion', 'notice'];
const 타임라인_하루 = 86400000;

// 웹 앱 배포 ID (Apps Script 프로젝트 scriptId별). 같은 배포를 버전만 갱신해 쓰므로 URL은 바뀌지 않는다.
const 심사일정웹앱배포 = {
  '1bx5uSpQYH1EHz39W_OcyZRcP2EVKX4AzRObZyvcAyO60XjpkirllRbE3': 'AKfycbx2k9LCdYyg-udjEgwDsMZ45Zn7TC49wjONd3hUOAm87SW9qeDhEjRtIcleDLMpHFSF', // dev
  '1UXFGYy3N5UkyqLkZdsd6TwVsgS583kMUmi-BtVPLvJ6zmew7ccPtRior': 'AKfycbwJJXf9RcMoGwFWRteGcKA7fId5oKEJ8vvKKWG_P9UTjaWD3IhAHk-prTm1F5-ikI26OA', // prod
};
const 심사일정_시트ID속성 = 'SCHEDULE_SPREADSHEET_ID';

function 심사일정타임라인열기() {
  _심사일정새탭열기_('timeline');
}

function 심사현황대시보드열기() {
  _심사일정새탭열기_('dashboard');
}

/** 시트 메뉴에서 웹 앱을 새 탭으로 연다. 팝업이 막히면 버튼을 대신 보여준다. */
function _심사일정새탭열기_(화면) {
  const ui = SpreadsheetApp.getUi();
  if (!_조회권한여부()) {
    ui.alert('관리자 또는 조회자로 등록된 계정만 볼 수 있는 화면입니다.');
    return;
  }
  // 웹 앱에는 "열린 시트"가 없으므로 이 시트의 ID를 기억해 둔다(바뀔 때만 기록).
  const 속성 = PropertiesService.getScriptProperties();
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  if (속성.getProperty(심사일정_시트ID속성) !== id) 속성.setProperty(심사일정_시트ID속성, id);

  const 기본URL = _심사일정웹앱URL_();
  if (!기본URL) {
    ui.alert('심사 일정 웹 앱이 아직 배포되지 않았습니다. 관리자에게 웹 앱 배포를 요청하세요.');
    return;
  }
  const url = JSON.stringify(`${기본URL}?view=${화면}`).replace(/</g, '\u003c');
  const html = HtmlService.createHtmlOutput(
    '<div style="font:13px Arial,\'Noto Sans KR\',sans-serif;padding:4px 6px">' +
    '<p id="m" style="margin:0 0 12px">새 탭을 여는 중입니다…</p>' +
    '<a id="a" target="_blank" rel="noopener" style="display:inline-block;padding:7px 14px;border-radius:6px;background:#306cc5;color:#fff;text-decoration:none">새 탭에서 열기</a></div>' +
    `<script>const u=${url};const a=document.getElementById('a');a.href=u;` +
    'a.addEventListener("click",()=>setTimeout(()=>google.script.host.close(),300));' +
    'const w=window.open(u,"_blank");' +
    'if(w){google.script.host.close();}else{document.getElementById("m").textContent="브라우저가 새 탭을 막았습니다. 아래 버튼을 눌러 주세요.";}</script>'
  ).setWidth(360).setHeight(120);
  ui.showModalDialog(html, 화면 === 'dashboard' ? '심사 현황 대시보드' : '심사 일정 타임라인');
}

function _심사일정웹앱URL_() {
  const 배포ID = 심사일정웹앱배포[ScriptApp.getScriptId()];
  if (배포ID) {
    // 도메인 전용 웹 앱은 /a/macros/<도메인>/ 주소로 열어야 여러 계정이 로그인된 브라우저에서도
    // 회사 계정으로 열린다(일반 /macros/s/ 주소는 "현재 파일을 열 수 없습니다"가 날 수 있음).
    let 도메인 = '';
    try { 도메인 = (Session.getActiveUser().getEmail() || '').split('@')[1] || ''; } catch (e) {}
    const 경로 = 도메인 && 도메인 !== 'gmail.com' ? `a/macros/${도메인}` : 'macros';
    return `https://script.google.com/${경로}/s/${배포ID}/exec`;
  }
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/** 시트 메뉴에서 실행되면 열린 시트, 웹 앱에서는 메뉴가 기억해 둔 시트를 연다. */
function _대상스프레드시트_() {
  const 활성 = SpreadsheetApp.getActiveSpreadsheet();
  if (활성) return 활성;
  const id = PropertiesService.getScriptProperties().getProperty(심사일정_시트ID속성);
  if (!id) throw new Error('시트 메뉴에서 타임라인이나 대시보드를 한 번 연 뒤 다시 시도하세요.');
  return SpreadsheetApp.openById(id);
}

/**
 * 웹 앱 진입점 — ?view=timeline | dashboard.
 * 배포한 계정 권한으로 실행되며(조회자에게 권한 동의 요구 없음), 접속한 사람의 이메일로 관리자·조회자만 허용한다.
 */
function doGet(e) {
  const 화면 = e && e.parameter && e.parameter.view === 'dashboard' ? 'dashboard' : 'timeline';
  if (!_조회권한여부()) {
    return HtmlService.createHtmlOutput('<p style="font:14px Arial,sans-serif;padding:24px">관리자 또는 조회자로 등록된 계정만 볼 수 있는 화면입니다.</p>')
      .setTitle('접근 권한 없음');
  }
  const 템플릿 = HtmlService.createTemplateFromFile(화면 === 'dashboard' ? 'DashboardView' : 'TimelineView');
  템플릿.appUrl = _심사일정웹앱URL_();
  return 템플릿.evaluate()
    .setTitle(화면 === 'dashboard' ? 'AI확인제 심사 현황' : 'AI확인제 심사 일정')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 타임라인·대시보드에서 google.script.run으로 호출 — 일정관리 시트를 화면용 자료로 변환한다. */
function 심사일정자료조회() {
  if (!_조회권한여부()) throw new Error('관리자 또는 조회자로 등록된 계정만 조회할 수 있습니다.');
  const ss = _대상스프레드시트_();
  const 시트 = ss.getSheetByName(SHEET.일정관리);
  if (!시트) throw new Error('일정관리 시트를 찾을 수 없습니다.');

  const tz = 'Asia/Seoul';
  const 지금 = new Date();
  const 오늘 = Utilities.formatDate(지금, tz, 'yyyy-MM-dd');
  const 휴일 = _타임라인공휴일_(ss, tz);
  const 결과 = {
    today: 오늘,
    readAt: Utilities.formatDate(지금, tz, 'yyyy.MM.dd HH:mm'),
    sourceUrl: `${ss.getUrl()}#gid=${시트.getSheetId()}`,
    gid: 시트.getSheetId(),
    lastColumn: columnLetter(Math.max(1, 시트.getLastColumn())),
    holidayWarning: 휴일.연도.has(Number(오늘.slice(0, 4)))
      ? '' : `공휴일 시트에 ${오늘.slice(0, 4)}년 공휴일이 없어 WD가 주말만 제외하고 계산되었습니다.`,
    // 대시보드 일별 히트맵에서 공휴일 칸을 표시하기 위한 목록(yyyy-MM-dd)
    holidays: [...휴일.날짜].sort(),
    // 현황 요약의 심사원 평균 투입률(스크립트 속성 REVIEW_RATE, 예: 80%). 화면에서 배정 건수 비례로 사람별로 나눈다(1인 최대 100%). 없으면 null(=100%).
    reviewRate: _투입률정규화_(PropertiesService.getScriptProperties().getProperty(현황요약_투입률속성)),
    // Gemini 보고 문장 버튼 표시 여부(키 값은 화면에 보내지 않는다)
    aiSummary: !!PropertiesService.getScriptProperties().getProperty(현황요약_키속성),
    // 심사원관리 시트의 활성여부·제외(제외 시작일). 현황 요약에서 현재 가용 인원과 담당자 제외 건을 계산한다.
    reviewers: _심사원제외읽기_(ss, tz, 오늘),
    // 표준 건당 MD(스크립트 속성 REVIEW_MD_PER_CASE, 예: 2.5). 있으면 인력 판단을 "필요 MD vs 가용 MD"로 한다.
    mdPerCaseStd: (v => (isFinite(v) && v > 0 ? v : null))(parseFloat(PropertiesService.getScriptProperties().getProperty('REVIEW_MD_PER_CASE'))),
    records: [],
  };
  const 마지막행 = 시트.getLastRow();
  const 마지막열 = 시트.getLastColumn();
  if (마지막행 < 2) return 결과;

  const 전체 = 시트.getRange(1, 1, 마지막행, 마지막열).getValues();
  const 헤더 = 전체[0].map(v => String(v).trim());
  const 위치 = {};
  Object.keys(타임라인_컬럼).forEach(키 => { 위치[키] = 헤더.indexOf(타임라인_컬럼[키]); });
  if (위치.id < 0) throw new Error('일정관리 시트에서 접수번호 헤더를 찾을 수 없습니다.');

  const 값 = (행, 키) => {
    if (위치[키] < 0) return '';
    const v = 행[위치[키]];
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    }
    return v == null ? '' : String(v).trim();
  };

  for (let i = 1; i < 전체.length; i++) {
    const 행 = 전체[i];
    const id = 값(행, 'id');
    if (!id) continue;
    const 레코드 = { row: i + 1 };
    Object.keys(타임라인_컬럼).forEach(키 => { 레코드[키] = 값(행, 키); });
    const 날짜오류 = [];
    타임라인_날짜키.forEach(키 => {
      if (!레코드[키]) return;
      const 정규 = _타임라인날짜정규화_(레코드[키]);
      if (!정규) 날짜오류.push(`${타임라인_컬럼[키]} 값 확인 필요(${레코드[키]})`);
      레코드[키] = 정규;
    });
    if (날짜오류.length) 레코드.dateIssues = 날짜오류;
    _타임라인WD계산_(레코드, 오늘, 휴일.날짜);
    결과.records.push(레코드);
  }
  return 결과;
}

/** 공휴일 시트(A열 날짜)를 읽기만 한다. 누락 연도 동기화는 마감예정일 갱신 메뉴가 담당한다. */
function _타임라인공휴일_(ss, tz) {
  const 날짜 = new Set();
  const 연도 = new Set();
  const 시트 = ss.getSheetByName(공휴일시트명);
  if (!시트 || 시트.getLastRow() < 2) return { 날짜, 연도 };
  시트.getRange(2, 1, 시트.getLastRow() - 1, 1).getValues().forEach(r => {
    const v = r[0];
    const 키 = Object.prototype.toString.call(v) === '[object Date]'
      ? (isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd'))
      : _타임라인날짜정규화_(v);
    if (!키) return;
    날짜.add(키);
    연도.add(Number(키.slice(0, 4)));
  });
  return { 날짜, 연도 };
}

/** 'yy-MM-dd'·'yyyy-MM-dd'·'yyyy.MM.dd' 등을 'yyyy-MM-dd'로 맞춘다. 날짜가 아니면 ''. */
function _타임라인날짜정규화_(값) {
  const m = String(값 == null ? '' : 값).trim().match(/^(\d{2}|\d{4})[-/.]\s*(\d{1,2})[-/.]\s*(\d{1,2})\s*\.?$/);
  if (!m) return '';
  const y = m[1].length === 2 ? 2000 + Number(m[1]) : Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const t = new Date(Date.UTC(y, mo - 1, d));
  if (t.getUTCFullYear() !== y || t.getUTCMonth() !== mo - 1 || t.getUTCDate() !== d) return '';
  return t.toISOString().slice(0, 10);
}

/**
 * 시작일 다음 날부터 종료일까지(시작 제외·종료 포함)의 영업일 수. 종료가 앞서면 음수.
 * WORKDAY(시작, n) = 종료 일 때 n과 같은 값이 되도록 맞춘 정의다.
 */
function _영업일수_(시작, 종료, 휴일) {
  if (!시작 || !종료) return null;
  const a = Date.parse(시작 + 'T00:00:00Z');
  const b = Date.parse(종료 + 'T00:00:00Z');
  if (isNaN(a) || isNaN(b)) return null;
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  let n = 0;
  for (let t = lo + 타임라인_하루; t <= hi; t += 타임라인_하루) {
    const d = new Date(t);
    const 요일 = d.getUTCDay();
    if (요일 === 0 || 요일 === 6) continue;
    if (휴일.has(d.toISOString().slice(0, 10))) continue;
    n++;
  }
  return b >= a ? n : -n;
}

/**
 * 심사원 처리와 행정 통보를 나눠 계산한다.
 * elapsedWD      : 접수일자 → 심사완료일(없으면 적합통보일). 아직 진행 중이면 오늘까지.
 * remainingWD    : 오늘 → 현재 마감(연장마감일 우선). 적합통보 전·종료 아닌 건만, 지났으면 음수.
 * reviewWD       : 심사접수일 → 심사완료일(없으면 적합통보일). TTA 인수 이후 실제 심사 기간. 둘 중 하나라도 없으면 null.
 * noticeLagWD    : 심사완료일 → 적합통보일(통보 전이면 오늘까지). 심사완료일이 없으면 null.
 * pendingNotice  : 통보 대기 — 종료가 아니고 적합통보일이 없는데 심사완료일이 있거나 상태가 완료인 건.
 */
function _타임라인WD계산_(레코드, 오늘, 휴일) {
  const 종료 = /^종료/.test(레코드.status);
  const 상태완료 = /^완료/.test(레코드.status);
  const 마감 = 레코드.extension || 레코드.original;
  const 심사끝 = 레코드.completion || 레코드.notice;
  레코드.pendingNotice = !종료 && !레코드.notice && (!!레코드.completion || 상태완료);
  레코드.elapsedWD = 심사끝
    ? _영업일수_(레코드.received, 심사끝, 휴일)
    : 종료 || 상태완료 ? null : _영업일수_(레코드.received, 오늘, 휴일);
  레코드.remainingWD = 종료 || 레코드.notice ? null : _영업일수_(오늘, 마감, 휴일);
  레코드.reviewWD = 레코드.review && 심사끝 ? _영업일수_(레코드.review, 심사끝, 휴일) : null;
  레코드.noticeLagWD = 레코드.completion
    ? _영업일수_(레코드.completion, 레코드.notice || 오늘, 휴일)
    : null;
}

/* ─────────────────────────────────────────────
 * 현황 요약: 투입률 설정 + Gemini 보고 문장
 * ───────────────────────────────────────────── */
const 현황요약_투입률속성 = 'REVIEW_RATE';
const 현황요약_키속성 = 'GEMINI_API_KEY';
const 현황요약_모델속성 = 'GEMINI_MODEL';
const 현황요약_예비모델속성 = 'GEMINI_FALLBACK_MODEL';
const 현황요약_기본모델 = 'gemini-3.8-flash';

/** 30%·0.3·30 → 0.3. 비었거나 숫자가 아니면 null. */
function _투입률정규화_(v) {
  if (v === '' || v == null) return null;
  const 문자 = String(v).trim();
  let n = typeof v === 'number' ? v : Number(문자.replace('%', ''));
  if (!isFinite(n) || n < 0) return null;
  if (문자.includes('%') || n > 1) n /= 100;
  return Math.min(n, 1);
}

/**
 * 이 API 키로 generateContent를 쓸 수 있는 flash 계열 모델을 찾는다(기본 모델 제외, 최신 버전 우선, 최대 2개).
 * 목록은 하루 캐시. 실패하면 빈 목록.
 */
function _Gemini예비모델_(키, 기본) {
  const 캐시 = CacheService.getScriptCache();
  const 저장 = 캐시.get('gemini-models');
  let 이름들 = 저장 ? JSON.parse(저장) : null;
  if (!이름들) {
    try {
      const 응답 = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
        headers: { 'x-goog-api-key': 키 }, muteHttpExceptions: true,
      });
      if (응답.getResponseCode() !== 200) return [];
      이름들 = (JSON.parse(응답.getContentText()).models || [])
        .filter(m => (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0)
        .map(m => String(m.name || '').replace(/^models\//, ''))
        .filter(n => /^gemini-[\d.]+-flash(-lite)?$/.test(n));
      캐시.put('gemini-models', JSON.stringify(이름들), 24 * 3600);
    } catch (e) {
      return [];
    }
  }
  const 버전 = n => parseFloat((n.match(/gemini-([\d.]+)/) || [])[1]) || 0;
  return 이름들.filter(n => n !== 기본)
    .sort((a, b) => 버전(b) - 버전(a) || (a.includes('lite') ? 1 : 0) - (b.includes('lite') ? 1 : 0))
    .slice(0, 2);
}

/**
 * 심사원관리 시트(읽기만): 심사원명·활성여부·제외·투입(투입일). 제외는 '10/6'처럼 월/일만 적어도 되며 오늘이 속한 연도로 본다.
 * 활성여부가 N인데 제외일이 비어 있으면 excludeFrom='' + active=false(시점 모름 → 지금은 제외).
 */
function _심사원제외읽기_(ss, tz, 오늘) {
  const 시트 = ss.getSheetByName(SHEET.심사원관리);
  if (!시트 || 시트.getLastRow() < 2) return [];
  const 값 = 시트.getRange(1, 1, 시트.getLastRow(), 시트.getLastColumn()).getValues();
  const 헤더 = 값[0].map(v => String(v).trim());
  const 이름 = 헤더.indexOf('심사원명');
  const 활성 = 헤더.indexOf('활성여부');
  const 제외 = 헤더.indexOf('제외');
  const 투입 = 헤더.indexOf('투입') >= 0 ? 헤더.indexOf('투입') : 헤더.indexOf('투입일');
  if (이름 < 0) return [];
  return 값.slice(1).map(행 => {
    const name = String(행[이름] == null ? '' : 행[이름]).trim();
    if (!name) return null;
    return {
      name,
      active: 활성 < 0 || String(행[활성] || '').trim().toUpperCase() !== 'N',
      excludeFrom: 제외 < 0 ? '' : _제외일정규화_(행[제외], tz, 오늘),
      startFrom: 투입 < 0 ? '' : _제외일정규화_(행[투입], tz, 오늘),
    };
  }).filter(Boolean);
}

function _제외일정규화_(v, tz, 오늘) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  const 문자 = String(v == null ? '' : v).trim();
  const 전체 = _타임라인날짜정규화_(문자);
  if (전체) return 전체;
  const m = 문자.match(/^(d{1,2})s*[/.-월]s*(d{1,2})s*일?.?$/);
  return m ? _타임라인날짜정규화_(`${오늘.slice(0, 4)}-${m[1]}-${m[2]}`) : '';
}

/**
 * 대시보드 "Gemini 분석" — 화면이 만든 현황 MD(심사원 이름은 심사원1·2…로 바꾼 것)와 집계값을 Gemini에 보내
 * 경영진 보고용 분석(핵심 메시지·위험·권고)을 받는다. 기업명·제품명·접수번호는 MD에 없다. 같은 자료는 6시간 캐시.
 */
function 현황요약AI문장(사실) {
  if (!_조회권한여부()) throw new Error('관리자 또는 조회자로 등록된 계정만 사용할 수 있습니다.');
  const 속성 = PropertiesService.getScriptProperties();
  const 키 = 속성.getProperty(현황요약_키속성);
  if (!키) throw new Error(`스크립트 속성 ${현황요약_키속성}가 없습니다.`);
  const 본문 = JSON.stringify(사실 || {});
  if (본문.length > 80000) throw new Error('요약 자료가 너무 큽니다.');
  const 모델 = 속성.getProperty(현황요약_모델속성) || 현황요약_기본모델;

  const 캐시 = CacheService.getScriptCache();
  const 캐시키 = 'ai-summary-' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 모델 + 본문)).slice(0, 40);
  const 저장 = 캐시.get(캐시키);
  if (저장) return { text: 저장, model: 모델, cached: true };

  const 지시 = [
    '너는 TTA 인공지능 제품·서비스 확인제 심사 운영을 분석해 경영진에게 보고하는 분석가다.',
    '사용자가 주는 md(대시보드 전체 현황, 마크다운)와 facts(같은 내용의 집계값 JSON)만 근거로 쓴다.',
    '대시보드에 이미 있는 문장을 되풀이하지 말고, 여러 지표를 엮어 "그래서 무엇이 문제이고 무엇을 해야 하는가"를 말한다.',
    '규칙:',
    '1) md·facts에 없는 숫자·사실을 만들지 않는다. 숫자는 그대로 인용하고, 새로 계산한 값은 쓰지 않는다.',
    '2) 원인을 단정하지 않는다(예: 야근 때문이라고 쓰지 않는다). 가능성은 "~로 보입니다", "~일 수 있습니다"로 쓴다.',
    '3) 인력판단이 "여유"가 아니면 인력에 여유가 있다고 쓰지 않는다. 풀타임 환산 인원과 일수(MD)는 배정 건수로 추정한 값임을 한 번 밝힌다. MD·명분 같은 용어 대신 "풀타임 N명", "N일"처럼 쉬운 말로 쓴다.',
    '4) 형식(평문, 마크다운 제목·굵게 없이):',
    '   결론: 한두 문장. "수요가 늘고 있는가", "인력이 남는가 모자라는가"에 직접 답한다.',
    '   핵심 메시지: "- "로 시작하는 3줄 이내. 각 줄에 근거 숫자 1~2개.',
    '   위험: "- "로 시작하는 2줄 이내. 지금 추세가 이어질 때 생길 일.',
    '   권고: "- "로 시작하는 3줄 이내. 경영진이 결정할 수 있는 조치(인력·재배정·프로세스 등).',
    '   마지막 줄: "※ "로 시작해 자료의 한계 한 줄.',
    '5) 보고체(~입니다/~습니다), 전체 900자 이내.',
  ].join('\n');

  // 혼잡(429·500·503)은 잠깐 쉬고 다시 시도한다. 그래도 안 되면 GEMINI_FALLBACK_MODEL(쉼표로 여러 개)을 차례로 쓴다.
  // 예비 모델: 스크립트 속성에 있으면 그것, 없으면 이 키로 쓸 수 있는 flash 계열 모델을 API에서 찾아 최대 2개.
  const 지정예비 = (속성.getProperty(현황요약_예비모델속성) || '').split(',').map(v => v.trim()).filter(v => v && v !== 모델);
  const 예비 = 지정예비.length ? 지정예비 : _Gemini예비모델_(키, 모델);
  const 요청본문 = JSON.stringify({
    systemInstruction: { parts: [{ text: 지시 }] },
    contents: [{ role: 'user', parts: [{ text: 본문 }] }],
    generationConfig: { temperature: 0.2 },
  });
  let 결과 = null;
  let 사용모델 = '';
  let 마지막오류 = '';
  for (const m of [모델].concat(예비)) {
    for (let 시도 = 0; 시도 < 3 && !결과; 시도++) {
      if (시도) Utilities.sleep(시도 === 1 ? 2000 : 5000);
      const 응답 = UrlFetchApp.fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(m)}:generateContent`, {
          method: 'post',
          contentType: 'application/json',
          headers: { 'x-goog-api-key': 키 },
          muteHttpExceptions: true,
          payload: 요청본문,
        });
      const 코드 = 응답.getResponseCode();
      let 내용 = null;
      try { 내용 = JSON.parse(응답.getContentText()); } catch (e) {}
      if (코드 === 200) { 결과 = 내용; 사용모델 = m; break; }
      마지막오류 = `${m} ${코드}: ${내용 && 내용.error && 내용.error.message || 응답.getContentText().slice(0, 200)}`;
      if ([429, 500, 503].indexOf(코드) < 0) break;   // 키·모델명 문제 등은 재시도해도 같으므로 다음 모델로
    }
    if (결과) break;
  }
  if (!결과) {
    throw new Error(`Gemini 호출 실패(${마지막오류}). 혼잡 오류면 잠시 뒤 다시 눌러 주세요` +
      (예비.length ? `(예비 모델 ${예비.join(', ')}도 실패).` : '.'));
  }
  const 후보 = 결과 && 결과.candidates && 결과.candidates[0];
  const 문장 = 후보 && 후보.content && (후보.content.parts || []).map(p => p.text || '').join('').trim();
  if (!문장) throw new Error('Gemini 응답에 문장이 없습니다' + (후보 && 후보.finishReason ? `(${후보.finishReason})` : '') + '.');
  캐시.put(캐시키, 문장, 6 * 3600);
  return { text: 문장, model: 사용모델, cached: false };
}
