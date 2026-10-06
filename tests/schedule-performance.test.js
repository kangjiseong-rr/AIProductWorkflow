const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'Schedule.js'), 'utf8');
const context = {
  SHEET: { 일정관리: '일정관리' },
  Logger: { log() {} },
  columnLetter(n) {
    let s = '';
    while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); }
    return s;
  },
};
context.SpreadsheetApp = {
  newDataValidation: () => ({ requireValueInList() { return this; }, build() { return {}; } }),
  newConditionalFormatRule() {
    const rule = {};
    return {
      whenFormulaSatisfied(v) { rule.formula = v; return this; },
      setBackground(v) { rule.background = v; return this; },
      setFontColor(v) { rule.fontColor = v; return this; },
      setRanges(v) { rule.ranges = v; return this; },
      build() { return rule; },
    };
  },
};
vm.createContext(context);
vm.runInContext(source, context);

const standardHeaders = ['마감예정일', '상태', '보완요청일', '연장마감일', '특이사항', 'D3기준일_마감', 'D3기준일_연장마감'];
for (const headers of [standardHeaders, [...standardHeaders].reverse()]) {
  let rules;
  const sheet = {
    getLastColumn: () => headers.length,
    getLastRow: () => 100,
    getMaxRows() { throw new Error('조건부서식이 전체 그리드를 조회했습니다.'); },
    getRange(row, col, height, width) {
      assert(row + height - 1 <= 120, '조건부서식 범위가 데이터 + 20행을 넘었습니다.');
      return { row, col, height, width, getValues: () => [headers], setDataValidation() {} };
    },
    setConditionalFormatRules(v) { rules = v; },
  };
  context._일정관리조건부서식적용_(sheet);
  assert.strictEqual(rules.length, 14);
  const noteCol = headers.indexOf('특이사항') + 1;
  for (const rule of rules) {
    assert.strictEqual(rule.ranges.length, 1, '분리된 범위를 하나의 규칙으로 합쳤습니다.');
    assert(!/\b(WORKDAY|INDIRECT|INDEX|MATCH)\s*\(/i.test(rule.formula), '조건부서식에 무거운 조회/계산이 들어갔습니다.');
    const r = rule.ranges[0];
    assert(!(r.col <= noteCol && noteCol < r.col + r.width), '특이사항 수동 서식을 덮습니다.');
  }
  for (const [i, name] of ['마감예정일', '연장마감일'].entries()) {
    assert.strictEqual(rules[i].ranges[0].col, headers.indexOf(name) + 1);
    const helper = context.columnLetter(headers.indexOf(i === 0 ? 'D3기준일_마감' : 'D3기준일_연장마감') + 1);
    assert(rules[i].formula.includes(`ISNUMBER($${helper}2)`), '빈 D3 보조셀 검증이 빠졌습니다.');
  }
}

// 색상 갱신 메뉴가 고비용 작업을 끌어오면 즉시 실패한다.
const forbidden = () => { throw new Error('색상 갱신에서 날짜/공휴일/표 재생성을 호출했습니다.'); };
context._마감예정일수식갱신_ = forbidden;
context._공휴일연도확보_ = forbidden;
context._일정관리서식적용_ = forbidden;
context._일정관리구글표적용_ = forbidden;
context.UrlFetchApp = { fetch: forbidden };
let ruleUpdates = 0;
let flushes = 0;
context._일정관리조건부서식적용_ = () => { ruleUpdates++; return true; };
context.SpreadsheetApp.getActiveSpreadsheet = () => ({ getSheetByName: () => ({}), toast() {} });
context.SpreadsheetApp.flush = () => { flushes++; };
context.일정관리색상서식갱신();
assert.strictEqual(ruleUpdates, 1);
assert.strictEqual(flushes, 1);

// 마감일 메뉴가 색상 규칙을 다시 쓰지 않는 것도 검증한다.
context._마감예정일수식갱신_ = () => 1;
context.SpreadsheetApp.getUi = () => ({ alert() {} });
context.마감예정일갱신();
assert.strictEqual(ruleUpdates, 1);
console.log('Schedule performance guardrails passed.');
