const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const context = { console };
vm.createContext(context);
vm.runInContext(
  fs.readFileSync(path.join(root, 'FieldDefinition.js'), 'utf8') +
  '\n' + fs.readFileSync(path.join(root, 'ExcelParser.js'), 'utf8') +
  '\n' + fs.readFileSync(path.join(root, 'ReportGenerator.js'), 'utf8') +
  '\nglobalThis.__schemaTest = { 포매터, _기능탭데이터병합, _기타포함표시 };',
  context
);

const { 포매터, _기능탭데이터병합, _기타포함표시 } = context.__schemaTest;
const normalized = [
  포매터.사업자번호정규화('123-45-67890'),
  포매터.전화번호정규화('02-2188-6980'),
  포매터.전화번호정규화('010-1234-5678'),
];
const expected = ['1234567890', '0221886980', '01012345678'];
if (JSON.stringify(normalized) !== JSON.stringify(expected)) {
  throw new Error(`번호 정규화 실패: ${JSON.stringify(normalized)}`);
}

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
