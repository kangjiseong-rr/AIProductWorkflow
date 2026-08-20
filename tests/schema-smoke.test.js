const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const parserSource = fs.readFileSync(path.join(root, 'ExcelParser.js'), 'utf8');
if (!/function _엑셀파싱처리\([^)]*\)\s*\{\s*const ss = SpreadsheetApp\.getActiveSpreadsheet\(\);/.test(parserSource)) {
  throw new Error('_엑셀파싱처리의 활성 스프레드시트 선언이 없습니다.');
}
const uploadSource = fs.readFileSync(path.join(root, 'ExcelUpload.html'), 'utf8');
if (!uploadSource.includes("submitButton.type = 'button'") ||
    !uploadSource.includes('submitButton.disabled = false') ||
    !uploadSource.includes('google.script.host.close()')) {
  throw new Error('엑셀 업로드 완료 버튼의 활성화·닫기 처리가 없습니다.');
}
const mainSource = fs.readFileSync(path.join(root, 'Main.js'), 'utf8');
if (!/if \(이름 === SHEET\.일정관리\) \{\s*_일정관리서식적용_\(시트, true\);/.test(mainSource)) {
  throw new Error('일정관리 신규 컬럼 추가 후 표 서식 확장 처리가 없습니다.');
}
if (!/if \(!예정\.length\) \{[\s\S]*?_일정관리서식적용_\(일정시트, true\);/.test(mainSource)) {
  throw new Error('기존 일정관리 컬럼의 표 서식 새로고침 처리가 없습니다.');
}
const context = { console };
const reportSource = fs.readFileSync(path.join(root, 'ReportGenerator.js'), 'utf8');
for (const section of [
  '1. 심사 개요',
  '3. 심사 대상 제품',
  '5. 종합 심사 의견',
  '붙임 1. 기능별 인공지능 구현 세부 사항',
  '붙임 2. 데이터 구조도',
  '붙임 3. 첨부자료 목록',
]) {
  if (!reportSource.includes(`_새페이지섹션(body, '${section}')`)) {
    throw new Error(`새 페이지 시작 설정 누락: ${section}`);
  }
}
if (!reportSource.includes('p.setLineSpacing(1.0)') || !reportSource.includes('setFontSize(11)')) {
  throw new Error('심사결과 요약 표의 11pt·한 줄 간격 설정이 없습니다.');
}
if (!reportSource.includes('function _모든표공통스타일_') || !reportSource.includes('setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER)')) {
  throw new Error('보고서 전체 표의 한 줄 간격·세로 가운데 정렬 설정이 없습니다.');
}
vm.createContext(context);
vm.runInContext(
  fs.readFileSync(path.join(root, 'FieldDefinition.js'), 'utf8') +
  '\n' + parserSource +
  '\n' + reportSource +
  '\nglobalThis.__schemaTest = { 포매터, 필드정의, _기능탭데이터병합, _기타포함표시 };',
  context
);

const { 포매터, 필드정의, _기능탭데이터병합, _기타포함표시 } = context.__schemaTest;

const scheduleContext = {
  console,
  SHEET: { 접수대장: '접수대장' },
  columnLetter(n) {
    let s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  },
};
vm.createContext(scheduleContext);
vm.runInContext(
  fs.readFileSync(path.join(root, 'Schedule.js'), 'utf8') +
  '\nglobalThis.__formulaTest = _접수대장조회수식_;',
  scheduleContext
);
const lookupFormula = scheduleContext.__formulaTest(['순번', '접수번호', '기업명'], '기업명', '$B2', false);
if (!lookupFormula.includes('$B:$B') || !lookupFormula.includes('$C:$C') || lookupFormula.includes('VLOOKUP')) {
  throw new Error(`컬럼 이동 안전 조회 수식 생성 실패: ${lookupFormula}`);
}
const 날짜필드 = 필드정의.find(f => f.sheetColumn === '접수일자');
if (!날짜필드 || !날짜필드.excelAliases.includes('접수일자')) {
  throw new Error('접수일자 필드 매핑이 없습니다.');
}
if (필드정의.some(f => f.sheetColumn === '신청일')) {
  throw new Error('구형 신청일 시트 컬럼명이 남아 있습니다.');
}
const normalized = [
  포매터.사업자번호정규화('123-45-67890'),
  포매터.전화번호정규화('02-2188-6980'),
  포매터.전화번호정규화('010-1234-5678'),
  포매터.사업자번호정규화(123456789),
  포매터.전화번호정규화(1012345678),
  포매터.전화번호정규화(221886980),
  포매터.전화번호정규화(311234567),
  포매터.전화번호정규화(111234567),
  포매터.전화번호정규화(161234567),
  포매터.전화번호정규화(7012345678),
  포매터.전화번호정규화(15881234),
];
const expected = [
  '1234567890', '0221886980', '01012345678',
  '0123456789', '01012345678', '0221886980', '0311234567',
  '0111234567', '0161234567', '07012345678', '15881234',
];
if (JSON.stringify(normalized) !== JSON.stringify(expected)) {
  throw new Error(`번호 정규화 실패: ${JSON.stringify(normalized)}`);
}

const jsonPathExpectations = {
  '제공형태기타': 'serviceInfo.serviceTypeEtc',
  '제품분류기타': 'serviceInfo.aiTechCategoryEtc',
  '서비스도메인': 'serviceInfo.serviceDomain',
  '최종신청동의여부': 'agreements.finalSubmissionConsent',
};
Object.entries(jsonPathExpectations).forEach(([sheetColumn, expectedPath]) => {
  const definition = 필드정의.find(f => f.sheetColumn === sheetColumn);
  if (!definition || definition.path !== expectedPath || Object.prototype.hasOwnProperty.call(definition, 'fixedValue')) {
    throw new Error(`JSON 필드 경로 설정 실패: ${sheetColumn}`);
  }
});

const 핵심 = [
  ['신청번호', '기능번호', '기능명', '인공지능의 역할', '입력 데이터', '출력 데이터', '매뉴얼 참조 위치'],
  ['AI-1', 1, '분류', '분류 역할', '문서', '범주', 'p.1'],
];
const 구현 = [
  ['신청번호', '기능번호', '기능명', '구현방식', '연산자원', '흐름도파일명'],
  ['AI-1', 1, '분류', 'B', 'GPU', 'flow.pdf'],
];
const merged = _기능탭데이터병합(핵심, 구현);
const header = merged[0];
const row = merged[1];
if (
  merged.length !== 2 ||
  row[header.indexOf('구현방식')] !== 'B' ||
  row[header.indexOf('인공지능의 역할')] !== '분류 역할'
) {
  throw new Error(`기능 탭 병합 실패: ${JSON.stringify(merged)}`);
}

if (_기타포함표시('기타', 'API형') !== '기타(API형)') {
  throw new Error('제공형태 기타 표시 실패');
}

console.log(`Schema smoke test passed (${header.length} merged columns).`);
