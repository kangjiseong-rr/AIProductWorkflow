// 2026-10-21 사고 회귀 테스트: 빈칸("")을 돌려주는 수식이 시트 끝(1008행)까지 채워져 있으면
// getLastRow()가 1008을 돌려줘 새 건이 맨 아래에 붙었다. 접수번호 기준 마지막 행 다음에 써야 한다.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const eq = (a, b, label) => { if (a !== b) throw new Error(`${label}: expected ${b}, got ${a}`); };

function fakeSheet({ dataRows, lastRow, id = 1 }) {
  // dataRows: 접수번호가 들어 있는 행 수(헤더 제외). lastRow: getLastRow()가 돌려줄 값(수식 꼬리 포함)
  const writes = [];
  let last = lastRow;
  const keys = Array.from({ length: Math.max(0, last - 1) }, (_, i) => [i < dataRows ? `AI-${i + 1}` : '']);
  return {
    writes,
    getSheetId: () => id,
    getName: () => '시트',
    getLastRow: () => last,
    getLastColumn: () => 3,
    getMaxRows: () => 1008,
    getRange: (r, c, nr, nc) => ({
      getValues: () => (r === 1 ? [['순번', '접수번호', '값']] : keys.slice(r - 2, r - 2 + nr)),
      setValues: v => { writes.push(r); if (keys[r - 2]) keys[r - 2] = [v[0][1]]; last = Math.max(last, r); return { setFontSize() {} }; },
      setNumberFormat: () => {},
    }),
  };
}

const ctx = { console, SHEET: { 접수대장: '접수대장', AI기능상세: '인공지능기능상세' }, SpreadsheetApp: { flush() {} }, Logger: { log() {} } };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'SheetWriter.js'), 'utf8') +
  '\nglobalThis.__a = { _데이터마지막행_, _테이블행추가 };', ctx);
const { _데이터마지막행_, _테이블행추가 } = ctx.__a;
const H = ['순번', '접수번호', '값'];

// 1) 수식 꼬리 1008행, 실제 데이터 116건(117행까지) → 마지막 데이터 행 117
eq(_데이터마지막행_(fakeSheet({ dataRows: 116, lastRow: 1008 }), H), 117, '수식 꼬리 무시');
// 2) 꼬리 없음 → getLastRow와 같음
eq(_데이터마지막행_(fakeSheet({ dataRows: 116, lastRow: 117 }), H), 117, '정상 시트');
// 3) 헤더만 → 1
eq(_데이터마지막행_(fakeSheet({ dataRows: 0, lastRow: 1 }), H), 1, '빈 시트');

// 4) 연속 추가(모델 여러 건): 118, 119, 120 순서로 쓰여야 한다(같은 실행에서 위치 기억)
const s = fakeSheet({ dataRows: 116, lastRow: 1008, id: 7 });
_테이블행추가(s, [1, 'AI-new', 'x'], H, null);
_테이블행추가(s, [2, 'AI-new', 'y'], H, null);
_테이블행추가(s, [3, 'AI-new', 'z'], H, null);
eq(s.writes.join(','), '118,119,120', '연속 추가 위치');

// 5) 꼬리 없는 일반 시트도 기존처럼 마지막 행 다음에 쓴다
const t = fakeSheet({ dataRows: 5, lastRow: 6, id: 8 });
_테이블행추가(t, [1, 'AI-new', 'x'], H, null);
_테이블행추가(t, [2, 'AI-new', 'y'], H, null);
eq(t.writes.join(','), '7,8', '일반 시트 추가 위치');

// 6) 일정관리 행 추가도 같은 기준을 써야 한다(appendRow 금지)
const schedule = fs.readFileSync(path.join(root, 'Schedule.js'), 'utf8');
if (/일정시트\.appendRow\(/.test(schedule)) throw new Error('일정관리 행 추가에 appendRow를 다시 쓰고 있습니다.');
if (!/const 새행번호 = _데이터마지막행_\(일정시트, 일정H\) \+ 1;/.test(schedule)) throw new Error('일정관리 새 행 번호가 접수번호 기준이 아닙니다.');

console.log('Append row test passed.');
