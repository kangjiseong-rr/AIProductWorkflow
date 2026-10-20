/**
 * Timeline.js — 관리자 전용 "심사 일정 타임라인"·"심사 현황 대시보드" 다이얼로그
 *
 * 일정관리 시트를 헤더명 기준으로 읽어 두 화면(TimelineView.html, DashboardView.html)이
 * 같은 자료를 쓰도록 한다. 경과·잔여·처리기간은 마감 계산(WORKDAY)과 같은 기준인
 * 영업일(WD: 주말·공휴일 시트 날짜 제외)로 서버에서 한 번만 계산한다.
 * 시트는 읽기만 하고 서식·수식·공휴일 동기화(네트워크 조회)는 건드리지 않는다.
 */

const 타임라인_컬럼 = {
  id: '접수번호',
  received: '접수일자',
  review: '심사접수일',
  original: '마감예정일',
  status: '상태',
  supplement: '보완요청일',
  extension: '연장마감일',
  notice: '적합통보일',
  reviewer: '담당심사원',
  note: '특이사항',
  company: '기업명',
  product: '제품명',
  count: '제품수',
};
const 타임라인_날짜키 = ['received', 'review', 'original', 'supplement', 'extension', 'notice'];
const 타임라인_하루 = 86400000;

function 심사일정타임라인열기() {
  _심사일정창열기_('TimelineView', 'AI확인제 심사 일정 타임라인');
}

function 심사현황대시보드열기() {
  _심사일정창열기_('DashboardView', 'AI확인제 심사 현황 대시보드');
}

function _심사일정창열기_(파일명, 제목) {
  if (!_관리자여부()) {
    SpreadsheetApp.getUi().alert('관리자만 볼 수 있는 화면입니다.');
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile(파일명).setWidth(1280).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 제목);
}

/** 타임라인·대시보드에서 google.script.run으로 호출 — 일정관리 시트를 화면용 자료로 변환한다. */
function 심사일정자료조회() {
  if (!_관리자여부()) throw new Error('관리자만 조회할 수 있습니다.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
 * elapsedWD  : 접수일자 → (완료·종료는 적합통보일, 진행 중은 오늘)
 * remainingWD: 오늘 → 현재 마감(연장마감일 우선). 진행 중인 건만, 지났으면 음수.
 */
function _타임라인WD계산_(레코드, 오늘, 휴일) {
  const 종결 = /^(완료|종료)/.test(레코드.status);
  const 마감 = 레코드.extension || 레코드.original;
  레코드.elapsedWD = 종결
    ? _영업일수_(레코드.received, 레코드.notice, 휴일)
    : _영업일수_(레코드.received, 오늘, 휴일);
  레코드.remainingWD = 종결 ? null : _영업일수_(오늘, 마감, 휴일);
}
