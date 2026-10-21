const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const context = {
  console,
  SHEET: { 일정관리: '일정관리' },
  공휴일시트명: '공휴일',
  columnLetter: n => String.fromCharCode(64 + n),
  _관리자여부: () => true,
  _조회권한여부: () => true,
  Utilities: {
    formatDate: (d, tz, f) => {
      const k = new Date(d.getTime() + 9 * 3600000).toISOString();
      return f === 'yyyy-MM-dd' ? k.slice(0, 10) : k.slice(0, 16);
    },
  },
};
vm.createContext(context);
vm.runInContext(fs.readFileSync(path.join(root, 'Timeline.js'), 'utf8') +
  '\nglobalThis.__t = { _영업일수_, _타임라인날짜정규화_, 심사일정자료조회 };', context);
const { _영업일수_, _타임라인날짜정규화_, 심사일정자료조회 } = context.__t;

const eq = (actual, expected, label) => {
  if (actual !== expected) throw new Error(`${label}: expected ${expected}, got ${actual}`);
};

const 휴일 = new Set(['2026-10-05', '2026-10-09']);
// 2026-10-02(금) → 2026-10-12(월): 10/5 대체공휴일·10/9 한글날, 주말 제외 → 10/6,7,8,12 = 4WD
eq(_영업일수_('2026-10-02', '2026-10-12', 휴일), 4, '공휴일·주말 제외');
eq(_영업일수_('2026-10-12', '2026-10-02', 휴일), -4, '역방향은 음수');
eq(_영업일수_('2026-10-12', '2026-10-12', 휴일), 0, '같은 날은 0');
eq(_영업일수_('2026-10-16', '2026-10-17', 휴일), 0, '금→토는 0WD');
eq(_영업일수_('', '2026-10-17', 휴일), null, '빈 날짜');
// WORKDAY(2026-09-01, 15) = 2026-09-22 (공휴일 없는 구간)
eq(_영업일수_('2026-09-01', '2026-09-22', new Set()), 15, 'WORKDAY 15WD 정의 일치');

eq(_타임라인날짜정규화_('26-07-23'), '2026-07-23', '2자리 연도');
eq(_타임라인날짜정규화_('2026.10.06'), '2026-10-06', '점 구분');
eq(_타임라인날짜정규화_('#N/A'), '', '수식 오류');
eq(_타임라인날짜정규화_('2026-02-30'), '', '없는 날짜');

const d = s => new Date(s + 'T00:00:00+09:00');
const 일정 = [
  ['접수번호', '접수일자', '심사접수일', '마감예정일', '상태', '연장마감일', '심사완료일', '적합통보일', '기업명', '제품명'],
  ['AI-1', d('2026-09-01'), '', d('2026-09-22'), '완료(적합)', '', '', d('2026-09-15'), 'A', 'P'],
  ['AI-2', d('2026-10-01'), '', d('2026-10-23'), '심사중', '#N/A', '', '', 'B', 'Q'],
  ['', '', '', '', '', '', '', '', '', ''],
  ['AI-3', d('2026-09-01'), d('2026-09-03'), d('2026-09-22'), '완료(적합)', '', d('2026-09-10'), d('2026-09-15'), 'C', 'R'],
  ['AI-4', d('2026-10-01'), '', d('2026-10-23'), '심사중', '', d('2026-10-13'), '', 'D', 'S'],
  ['AI-5', d('2026-10-01'), '', d('2026-10-23'), '완료(적합)', '', '', '', 'E', 'T'],
];
const 공휴일 = [['날짜'], [d('2026-10-05')], [d('2026-10-09')]];
const sheet = rows => ({
  getSheetId: () => 7,
  getLastRow: () => rows.length,
  getLastColumn: () => rows[0].length,
  getRange: (r, c, nr, nc) => ({
    getValues: () => rows.slice(r - 1, r - 1 + nr).map(row => row.slice(c - 1, c - 1 + nc)),
  }),
});
context.SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getUrl: () => 'https://sheet',
    getSheetByName: n => (n === '일정관리' ? sheet(일정) : n === '공휴일' ? sheet(공휴일) : null),
  }),
};
const 원래Date = context.Date;
const 결과 = vm.runInContext(`(() => {
  const RealDate = Date;
  globalThis.Date = class extends RealDate {
    constructor(...a) { if (a.length) super(...a); else super('2026-10-16T01:00:00Z'); }
    static parse(s) { return RealDate.parse(s); }
    static UTC(...a) { return RealDate.UTC(...a); }
  };
  try { return 심사일정자료조회(); } finally { globalThis.Date = RealDate; }
})()`, context);
void 원래Date;

eq(결과.today, '2026-10-16', '오늘(한국)');
eq(결과.records.length, 5, '빈 행 제외');
eq(결과.records[0].elapsedWD, 10, '완료 건 처리기간 = 접수일자→적합통보일 WD');
eq(결과.records[0].remainingWD, null, '완료 건은 잔여 없음');
// 10/1 → 10/16: 10/2,6,7,8,12,13,14,15,16 = 9WD (10/5·10/9 공휴일)
eq(결과.records[1].elapsedWD, 9, '진행 건 경과 WD');
// 10/16 → 10/23: 19,20,21,22,23 = 5WD (연장마감 #N/A → 최초 마감 사용)
eq(결과.records[1].remainingWD, 5, '진행 건 잔여 WD');
eq(결과.records[1].extension, '', '오류 날짜는 비움');
if (!결과.records[1].dateIssues || !결과.records[1].dateIssues.length) throw new Error('날짜 오류 표시 누락');
eq(결과.holidayWarning, '', '공휴일 연도 있음');
eq(결과.records[0].pendingNotice, false, '통보 완료 건은 통보 대기 아님');
// 심사완료일이 있으면 심사원 처리기간은 접수일자 → 심사완료일 (9/1→9/10 = 7WD), 통보 소요 9/10→9/15 = 3WD
eq(결과.records[2].elapsedWD, 7, '심사원 처리기간은 심사완료일까지');
eq(결과.records[2].noticeLagWD, 3, '심사완료 → 통보 WD');
// 심사접수일 기준: 9/3 → 심사완료 9/10 = 5WD
eq(결과.records[2].reviewWD, 5, '심사접수일 → 심사완료일 WD');
eq(결과.records[0].reviewWD, null, '심사접수일 없으면 null');
// 심사완료 후 통보 전: 통보 대기, 처리기간 10/1→10/13 = 6WD, 통보 대기 10/13→10/16 = 3WD, 마감까지 5WD
eq(결과.records[3].pendingNotice, true, '심사완료·통보 전은 통보 대기');
eq(결과.records[3].elapsedWD, 6, '통보 대기 건 처리기간');
eq(결과.records[3].noticeLagWD, 3, '통보 대기 경과 WD');
eq(결과.records[3].remainingWD, 5, '통보 전이면 마감까지 잔여 계산');
// 상태만 완료이고 날짜가 모두 빈 건도 통보 대기로 잡는다(행정 착오 점검)
eq(결과.records[4].pendingNotice, true, '완료 상태인데 통보일 없음');
eq(결과.records[4].elapsedWD, null, '완료일 정보 없으면 처리기간 없음');

const main = fs.readFileSync(path.join(root, 'Main.js'), 'utf8');
const ignore = fs.readFileSync(path.join(root, '.claspignore'), 'utf8');
for (const [fn, file] of [['심사일정타임라인열기', 'TimelineView.html'], ['심사현황대시보드열기', 'DashboardView.html']]) {
  if (!main.includes(`'${fn}'`)) throw new Error(`메뉴 연결 누락: ${fn}`);
  if (!ignore.includes(`!${file}`)) throw new Error(`.claspignore 누락: ${file}`);
  const html = fs.readFileSync(path.join(root, file), 'utf8');
  if (!html.includes('.심사일정자료조회()')) throw new Error(`${file}가 공용 조회 함수를 쓰지 않습니다.`);
  new Function(html.match(/<script>([\s\S]*?)<\/script>/)[1]);
}
if (!ignore.includes('!Timeline.js')) throw new Error('.claspignore 누락: Timeline.js');

// 조회 웹 앱은 배포한 계정 권한으로 실행해 조회자에게 프로젝트 전체 권한 동의를 요구하지 않는다(도메인 한정 + 관리자 확인 유지).
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
if (!manifest.webapp || manifest.webapp.executeAs !== 'USER_DEPLOYING' || manifest.webapp.access !== 'DOMAIN') {
  throw new Error('웹 앱은 executeAs USER_DEPLOYING, access DOMAIN 이어야 합니다: ' + JSON.stringify(manifest.webapp));
}
const timelineSrc = fs.readFileSync(path.join(root, 'Timeline.js'), 'utf8');
if (!/function doGet\(e\) \{[\s\S]*?if \(!_조회권한여부\(\)\)/.test(timelineSrc)) throw new Error('doGet의 조회 권한 확인이 없습니다.');
if (!/function 심사일정자료조회\(\) \{\s*if \(!_조회권한여부\(\)\)/.test(timelineSrc)) throw new Error('심사일정자료조회의 조회 권한 확인이 없습니다.');

console.log('Timeline workday test passed.');
