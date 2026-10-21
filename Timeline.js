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
  if (!_관리자여부()) {
    ui.alert('관리자만 볼 수 있는 화면입니다.');
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

/** 웹 앱 진입점 — ?view=timeline | dashboard. 실행 계정(접속한 사용자) 기준으로 관리자만 허용. */
function doGet(e) {
  const 화면 = e && e.parameter && e.parameter.view === 'dashboard' ? 'dashboard' : 'timeline';
  if (!_관리자여부()) {
    return HtmlService.createHtmlOutput('<p style="font:14px Arial,sans-serif;padding:24px">관리자만 볼 수 있는 화면입니다.</p>')
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
  if (!_관리자여부()) throw new Error('관리자만 조회할 수 있습니다.');
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
