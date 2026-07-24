/**
 * ============================================================
 *  기술심사보고서 생성 — 접수 건 데이터 → Google Docs 증적 명세서
 * ============================================================
 *
 *  "기술심사보고서 생성 (선택 행)" 메뉴로 실행되는 독립 기능입니다. 접수대장·
 *  제품모델·기능상세 시트에 이미 기록된 데이터를 모아 신청서 전 항목 + 구조도 +
 *  심사결과 체크리스트를 갖춘 Docs 문서를 생성합니다.
 * ============================================================
 */

const 보고서_표헤더색 = '#f1f3f3';
const 보고서_표헤더글자색 = '#202124';

// 심사 체크리스트 항목 정의 (엑셀 데이터로 자동 채워질 항목들)
// id: 고유키 / 항목: 질문 / 참조: 접수대장·AI기능상세에서 끌어올 값의 출처
// 심사 체크리스트 항목 정의 — 공식 확인 기준 (AI 제품·서비스 확인제 가이드라인)
//   구분1: 인공지능처리 연산체계 확인 [1.1~1.3]
//   구분2: 인공지능기능 확인 [2.1~2.3]
//   구분3: 외부 인공지능 서비스 연동 확인 [3.1]
//   방법: 서류 / 현장 / 서류·현장
const 체크항목정의 = [
  { id: 'C1', no: '1.1', 구분: '인공지능처리 연산체계 확인', 방법: '서류',
    항목: '제품·서비스에 인공지능처리 연산체계가 적용되어 있는가?', 참조: '인공지능기능수' },
  { id: 'C2', no: '1.2', 구분: '인공지능처리 연산체계 확인', 방법: '서류·현장',
    항목: '적용된 인공지능처리 연산체계가 학습·추론 등 인공지능으로서의 특성을 가지고 있는가?', 참조: '구현방식' },
  { id: 'C3', no: '1.3', 구분: '인공지능처리 연산체계 확인', 방법: '서류·현장',
    항목: '입력된 데이터가 인공지능처리 연산체계를 거쳐 유의미한 결과(예측·생성 등)를 출력하는가?', 참조: '입출력' },
  { id: 'C4', no: '2.1', 구분: '인공지능기능 확인', 방법: '서류',
    항목: '구조도상 인공지능(기술)의 위치가 명확하며, 제품·서비스 구동 체계와 결합되어 동작하는가?', 참조: '구조도' },
  { id: 'C5', no: '2.2', 구분: '인공지능기능 확인', 방법: '서류·현장',
    항목: '인공지능처리 연산체계가 제품·서비스의 기능·편의성·접근성·효율성 등에 활용되고 있음이 확인 가능한가?', 참조: '적용목적범위' },
  { id: 'C6', no: '2.3', 구분: '인공지능기능 확인', 방법: '현장',
    항목: '인공지능처리 연산체계가 적용된 주요 기능이 정상적으로 동작하는가?', 참조: '인공지능기능수' },
  { id: 'C7', no: '3.1', 구분: '외부 인공지능 서비스 연동 확인', 방법: '서류·현장',
    항목: '외부 인공지능처리 연산체계와 연동된 경우 호출 로그나 API Key를 확인 가능한가?', 참조: '외부API정보' },
  { id: 'C12', no: '0.0', 구분: '종합', 방법: '-',
    항목: '종합 — 위 확인 기준을 종합할 때 인공지능 제품·서비스로 확인 가능한가?', 참조: '종합' },
];


// ─────────────────────────────────────────────
// 5. 보고서 자동 생성 (Google Docs 병합)
// ─────────────────────────────────────────────

/**
 * 접수대장에서 행을 선택한 상태로 이 함수를 실행하면
 * 해당 건의 기술심사 보고서 초안을 Google Docs로 자동 생성합니다.
 *
 * 사용 방법:
 *  1. 접수대장 시트에서 보고서를 생성할 행 클릭
 *  2. 메뉴 > 인공지능심사관리 > 보고서 생성 (또는 아래 함수 직접 실행)
 */
/**
 * 활성 행에서 접수번호를 추출한다 (접수대장·일정관리 어느 탭이든 동작).
 * 트리거는 접수번호 하나이므로, 시트 종류와 무관하게 '접수번호' 컬럼 값만 읽는다.
 * @return {string|null} 접수번호 (없으면 null)
 */
function _활성행접수번호(ss) {
  const 시트 = ss.getActiveSheet();
  const 행 = 시트.getActiveRange().getRow();
  if (행 <= 1) return null;
  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0]
    .map(v => String(v).trim());
  const iNo = 헤더.indexOf('접수번호');
  if (iNo < 0) return null;
  const 접수번호 = 시트.getRange(행, iNo + 1).getValue();
  return 접수번호 ? String(접수번호).trim() : null;
}


function _구현방식레이블(코드) {
  const 맵 = {
    A: '자체 학습 모델',
    B: '오픈소스 모델',
    C: '외부 AI API 연동',
    D: '온디바이스 AI',
    E: '혼합형 AI 구성',
  };
  const raw = String(코드 || '').trim().charAt(0).toUpperCase();
  return 맵[raw] ?? (String(코드 || '').trim() ? String(코드).trim() : '미입력');
}

/**
 * 신청 폼의 "AI 기능 구현 방식" 섹션과 동일한 구조로 붙임1 세부 항목 행을 만든다.
 * 공통 항목 + 구현방식(A~E)별 전용 항목만 선택적으로 포함해, 폼에 없는 항목이
 * 보고서에 끼어들지 않게 한다.
 */
function _구현방식별세부행(f) {
  const 행 = [
    ['기능명', _v(f['기능명']), ''],
    ['구현 방식', _구현방식레이블(f['구현방식']), ''],
    ['인공지능역할', _v(f['인공지능역할']), ''],
  ];

  const 코드 = String(f['구현방식'] || '').trim().charAt(0).toUpperCase();
  const 전용필드 = {
    A: [
      ['학습 데이터 사양', f['학습데이터사양'], '데이터명·라이선싱/보유형태 포함'],
      ['개발환경·라이브러리·알고리즘', f['개발환경라이브러리알고리즘'], ''],
    ],
    B: [
      ['Base Model 명칭', f['BaseModel명칭'], ''],
      ['튜닝 방법', f['튜닝방법'], ''],
      ['튜닝 데이터셋', f['튜닝데이터셋'], ''],
    ],
    C: [['외부 API·모델', f['외부API정보'], '제공사 + 모델명/버전']],
    D: [
      ['타겟 하드웨어·OS', f['타겟HW_OS'], ''],
      ['추론 런타임', f['추론런타임'], ''],
    ],
    E: [
      ['혼합 구성 설명', f['혼합구성설명'], ''],
      ['모델별 역할 및 입출력 흐름', f['모델별역할및입출력흐름'], ''],
    ],
  };

  if (전용필드[코드]) {
    전용필드[코드].forEach(([label, v, note]) => 행.push([label, _v(v), note]));
  } else {
    // 구현방식 미입력 — 어느 유형인지 판단 불가하므로 값이 채워진 전용 필드는 누락 없이 표시
    Object.values(전용필드).flat().forEach(([label, v, note]) => {
      if (String(v || '').trim()) 행.push([label, _v(v), note]);
    });
  }

  행.push(['AI 연산 자원 요약', _v(f['연산자원요약']), '']);
  if (String(f['실행환경요약'] || '').trim()) {
    행.push(['AI 실행 환경 요약', _v(f['실행환경요약']), '선택 입력']);
  }
  행.push(['입력 데이터 설명', _v(f['입력데이터설명']), '']);
  행.push(['출력 데이터 설명', _v(f['출력데이터설명']), '']);
  if (String(f['기타참고자료파일명'] || '').trim()) {
    행.push(['기타 참고자료', _v(f['기타참고자료파일명']), '선택 입력']);
  }
  return 행;
}

/** 보고서·증적 연결용 기능 ID (원본 기능번호는 변경하지 않음) */
function _보고서기능ID(index) {
  return `F${String(index + 1).padStart(2, '0')}`;
}

// ═════════════════════════════════════════════════════════
// 10. 증적 명세서 자동 생성 (증적자료용 — 신청서 전 항목 + 구조도 + 심사결과)
// ═════════════════════════════════════════════════════════
//
// 증적 명세서는 신청서의 모든 데이터 항목을 누락 없이 수록합니다.
// 데이터 구조도 이미지는 접수대장의 '구조도파일명' 칸에 Drive 파일 ID가
// 있으면 끌어와 삽입합니다. 심사 판정·의견은 심사원이 생성된 문서에서
// 직접 기입합니다 (별도 체크리스트 웹앱 없음).

function 증적명세서생성() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 접수번호 = _활성행접수번호(ss);
  if (!접수번호) {
    SpreadsheetApp.getUi().alert('접수번호가 있는 데이터 행을 선택한 뒤 실행해주세요.\n(일정관리 또는 접수대장 탭)');
    return;
  }
  const 건 = _건조회(ss, 접수번호);
  if (!건) { SpreadsheetApp.getUi().alert('접수대장에서 "' + 접수번호 + '" 건을 찾을 수 없습니다.'); return; }

  const doc = _증적명세서Docs생성(ss, 건);
  SpreadsheetApp.getUi().alert('기술심사보고서가 생성되었습니다.\n\n' + doc.getUrl());
}

function 증적명세서생성_byId(접수번호) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 대장 = ss.getSheetByName(SHEET.접수대장);
  const D = 대장.getDataRange().getValues();
  const H = D[0];
  const 행 = D.slice(1).find(r => r[H.indexOf('접수번호')] === 접수번호);
  if (!행) throw new Error('접수번호를 찾을 수 없습니다: ' + 접수번호);
  const 건 = {};
  H.forEach((h, i) => { 건[h] = 행[i]; });
  const doc = _증적명세서Docs생성(ss, 건);
  return { url: doc.getUrl(), name: doc.getName() };
}

function _증적명세서Docs생성(ss, 건) {
  const 접수번호 = 건['접수번호'];
  const 오늘 = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 MM월 dd일');
  // 데이터 수집
  const 기능목록 = _AI기능상세조회(ss, 접수번호);
  const 제품모델목록 = _제품모델목록조회(ss, 접수번호);
  const R = CONFIG.보고서 || {};
  const 문서번호 = [R.문서번호접두, 접수번호, '01'].filter(Boolean).join('-');

  // 심사 항목 체크리스트 — 판정·사유는 심사원이 문서 생성 후 Docs에서 직접 기입
  const 결과행 = 체크항목정의.filter(d => d.id !== 'C12').map(d => [d.no, d.구분 || '', d.항목, '', '']);

  const 제목문 = `(${접수번호}) 기술심사보고서`;
  const doc = DocumentApp.create(제목문);
  const body = doc.getBody();
  _문서여백설정(body, 1, 2, 2, 2);
  _문서번호머리글생성(doc, 문서번호);

  // ═══════════════ 표지 ═══════════════
  body.appendParagraph('인공지능 제품·서비스 기술심사보고서')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const 표지정보 = body.appendParagraph(
    `${R.작성기관 || ''}\n보안등급: ${R.보안등급 || '-'}`
  );
  표지정보.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setForegroundColor('#5f6368');
  body.appendHorizontalRule();

  // ═══════════════ 1. 심사 개요 ═══════════════
  _명세섹션(body, '1. 심사 개요');
  _명세표(body, [
    ['접수번호', _v(접수번호)],
    ['심사 근거', _v(R.심사근거)],
    ['심사 방법', _v(R.심사방법)],
    ['심사일', 오늘],
    ['담당심사원', _v(건['담당심사원'])],
    ['기술 책임자', _v(건['기술책임자'])],
  ]);

  // ═══════════════ 2. 신청기업 현황 ═══════════════
  _명세섹션(body, '2. 신청기업 현황');
  _명세표(body, [
    ['기업명', _v(건['기업명'])],
    ['사업자등록번호', _v(건['사업자번호'])],
    ['대표자명', _v(건['대표자'])],
    ['담당자', `${_v(건['담당자명'])} / ${_v(건['연락처'])} / ${_v(건['이메일'])}`],
    ['주소', _v(건['소재지'])],
  ]);

  // ═══════════════ 3. 심사 대상 제품 ═══════════════
  _명세섹션(body, '3. 심사 대상 제품');
  _명세표(body, [
    ['제품명 / 세부품명번호 / 물품식별번호', _모델번호요약(제품모델목록, 건)],
    ['제품 또는 서비스 수', _제품수표시(제품모델목록, 건)],
    ['제공 형태', _v(건['제공형태'])],
    ['인공지능 제품 분류', _v(건['제품분류'])],
    ['제품 개요', _v(건['개요'])],
    ['인공지능 적용 목적', _v(건['인공지능적용목적'])],
    ['인공지능 적용 범위', _v(건['인공지능적용범위'])],
    ['제품 구조도', _구조도표시(건)],
    ['인공지능 기능 수', _v(건['인공지능기능수'])],
  ]);

  // ═══════════════ 4. 핵심 인공지능 기능 명세 ═══════════════
  _명세섹션(body, '4. 핵심 인공지능 기능 명세');
  if (기능목록.length) {
    기능목록.forEach((f, idx) => {
      const 원본기능번호 = _v(f['기능번호'] || idx + 1);
      const 기능ID = _보고서기능ID(idx);
      body.appendParagraph(`<기능 ${원본기능번호}: ${기능ID}>`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      const 표 = [
        ['구분', '내용'],
        ['기능 번호', 원본기능번호],
        ['기능 ID', 기능ID],
        ['기능명', _v(f['기능명'])],
        ['인공지능 역할', _v(f['인공지능역할'])],
        ['입력', _v(f['입력'] || f['입력데이터설명'])],
        ['출력', _v(f['출력'] || f['출력데이터설명'])],
        ['기타', _v(f['레퍼런스참조위치'] || f['기타참고자료파일명'])],
      ];
      const 기능명세표 = body.appendTable(표);
      _명세표헤더(기능명세표);
      _두열표스타일(기능명세표, _cm(3), _cm(13));
    });
  } else {
    body.appendParagraph('(인공지능 기능 데이터 없음)').editAsText().setForegroundColor('#9aa0a6');
  }

  // ═══════════════ 5. 종합 심사 의견 (접수대장 심사 결과 컬럼 기준) ═══════════════
  body.appendPageBreak();
  _명세섹션(body, '5. 종합 심사 의견');
  _명세표(body, [
    ['종합 판정', `${_v(건['종합판정'])}\n※ 선택값: 적합 / 보완요청 / 부적합(미충족)`],
    ['심사 완료일', _v(건['심사완료일'])],
    ['특이사항', `${_v(건['심사의견'])}\n※ 해당하는 경우 기재`],
  ]);

  _명세섹션(body, '5.1. 심사 항목별 검토 결과');
  _심사항목별검토결과표(body, 결과행);

  // ═══════════════ 붙임 ═══════════════
  body.appendPageBreak();
  _명세섹션(body, '붙임 1. 기능별 인공지능 구현 세부 사항');
  if (기능목록.length) {
    기능목록.forEach((f, idx) => {
      const 원본기능번호 = _v(f['기능번호'] || idx + 1);
      const 기능ID = _보고서기능ID(idx);
      body.appendParagraph(`<기능 ${원본기능번호}: ${기능ID}>`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      const 기능세부표 = [['ID', '구분', '내용', '비고']];
      _구현방식별세부행(f).forEach((row, rowIdx) => {
        const 세부ID = `${기능ID}-${String(rowIdx + 1).padStart(2, '0')}`;
        기능세부표.push([세부ID, row[0], row[1], row[2]]);
      });
      const 기능표 = body.appendTable(기능세부표);
      _명세표헤더(기능표);
      _기능세부표스타일(기능표);
    });
  } else {
    body.appendParagraph('(데이터 없음)').editAsText().setForegroundColor('#9aa0a6');
  }

  // ── 붙임 2. 데이터 구조도 (새 페이지) ──
  body.appendPageBreak();
  _명세섹션(body, '붙임 2. 데이터 구조도');
  const 구조도원본 = String(건['구조도파일명'] || '').trim();
  if (구조도원본) {
    const ID목록 = 구조도원본.split(/[,\n]+/).map(s => s.trim()).filter(Boolean);
    let 삽입수 = 0;
    ID목록.forEach((원본, idx) => {
      try {
        const 파일ID = _드라이브ID추출(원본);
        const 표시본 = _표시본Blob생성(파일ID);
        const img = body.appendImage(표시본);
        const 비율 = img.getHeight() / img.getWidth();
        img.setWidth(480).setHeight(Math.round(480 * 비율));
        const 파일 = DriveApp.getFileById(파일ID);
        body.appendParagraph(`[구조도 ${idx + 1}] ${파일.getName()} · 원본: ${파일.getUrl()}`)
          .editAsText().setForegroundColor('#5f6368').setFontSize(8);
        삽입수++;
      } catch (e) {
        body.appendParagraph(`[구조도 ${idx + 1} 불러오기 실패: ${e.message}]`)
          .editAsText().setForegroundColor('#c5221f');
      }
    });
    body.appendParagraph('');
    body.appendParagraph(`※ 위 이미지는 표시본(축소)입니다. 원본 ${삽입수}건은 Drive(→NAS) 원본 폴더에 보관됩니다.`)
      .editAsText().setForegroundColor('#9aa0a6').setFontSize(8);
  } else {
    body.appendParagraph('(구조도 파일 미등록 — 접수대장 "구조도파일ID" 칸에 입력. 여러 장은 쉼표로 구분)')
      .editAsText().setForegroundColor('#9aa0a6');
  }

  // ── 붙임 3. 첨부자료 목록 ──
  body.appendPageBreak();
  _명세섹션(body, '붙임 3. 첨부자료 목록');
  _명세표(body, [
    ['기존 인증·시험 결과', _v(건['비고'])],
  ]);

  doc.saveAndClose();
  _보고서를접수번호폴더로저장_(DriveApp.getFileById(doc.getId()), 접수번호);
  return doc;
}

function _문서여백설정(body, topCm, bottomCm, leftCm, rightCm) {
  body.setMarginTop(_cm(topCm))
    .setMarginBottom(_cm(bottomCm ?? topCm))
    .setMarginLeft(_cm(leftCm ?? topCm))
    .setMarginRight(_cm(rightCm ?? leftCm ?? topCm));
}

function _문서번호머리글생성(doc, 문서번호) {
  const header = doc.getHeader() || doc.addHeader();
  let paragraph = null;
  for (let i = 0; i < header.getNumChildren(); i++) {
    const child = header.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      paragraph = child.asParagraph();
      break;
    }
  }
  if (!paragraph) paragraph = header.appendParagraph('');
  paragraph.setText(`문서번호: ${문서번호}`);
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  paragraph.editAsText().setBold(true).setForegroundColor('#3c4043');
}

function _구조도표시(건) {
  const 구조도 = String(건['구조도파일명'] || '').trim();
  if (!구조도) return '(미제출)';
  const count = 구조도.split(/[,\n]+/).map(s => s.trim()).filter(Boolean).length;
  return `${count}건`;
}

function _심사항목별검토결과표(body, 결과행) {
  const 그룹순서 = [];
  const 그룹맵 = {};
  결과행.forEach(r => {
    const 그룹 = r[1] || '기타';
    if (!그룹맵[그룹]) {
      그룹맵[그룹] = [];
      그룹순서.push(그룹);
    }
    그룹맵[그룹].push(r);
  });

  그룹순서.forEach((그룹, idx) => {
    body.appendParagraph(`<${idx + 1}. ${그룹}>`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    const rows = [['번호', '심사항목', '판정', '비고']]
      .concat(그룹맵[그룹].map(r => [r[0], r[2], r[3], r[4]]));
    const table = body.appendTable(rows);
    _명세표헤더(table);
    _심사결과표스타일(table);
    _심사결과판정색칠(table);
  });
}

function _기능세부표스타일(table) {
  const widths = [_cm(2), _cm(3), _cm(9), _cm(3)];
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      if (widths[c]) cell.setWidth(widths[c]);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      if (c === 0 || c === 1) {
        _셀문단정렬(cell, DocumentApp.HorizontalAlignment.CENTER);
      } else {
        _셀문단정렬(cell, DocumentApp.HorizontalAlignment.LEFT);
      }
    }
  }
}

function _두열표스타일(table, firstWidth, secondWidth) {
  const widths = [firstWidth, secondWidth];
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      if (widths[c]) cell.setWidth(widths[c]);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      if (r === 0 || c === 0) {
        cell.editAsText().setBold(true);
        _셀문단정렬(cell, DocumentApp.HorizontalAlignment.CENTER);
      } else {
        _셀문단정렬(cell, DocumentApp.HorizontalAlignment.LEFT);
      }
    }
  }
}

function _심사결과표스타일(table) {
  const widths = [_cm(1.3), _cm(10), _cm(2), _cm(2.7)];
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      if (widths[c]) cell.setWidth(widths[c]);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      _셀문단정렬(cell, c === 1 || c === 3 ? DocumentApp.HorizontalAlignment.LEFT : DocumentApp.HorizontalAlignment.CENTER);
    }
  }
}

function _심사결과판정색칠(table) {
  const 색 = { '적합': '#137333', '보완': '#e37400', '부적합': '#c5221f', '미검토': '#9aa0a6', '해당없음': '#9aa0a6' };
  for (let r = 1; r < table.getNumRows(); r++) {
    const cell = table.getRow(r).getCell(2);
    const 판정 = cell.getText().trim();
    if (색[판정]) cell.editAsText().setForegroundColor(색[판정]).setBold(true);
  }
}

function _셀문단정렬(cell, alignment) {
  for (let i = 0; i < cell.getNumChildren(); i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      child.asParagraph().setAlignment(alignment);
    }
  }
}

function _cm(value) {
  return value * 28.3464567;
}

// 텍스트 요약 (길면 잘라서 …)
function _요약(val, 길이) {
  const s = String(val == null ? '' : val).trim();
  if (s === '') return '(미기재)';
  return s.length > 길이 ? s.slice(0, 길이) + '…' : s;
}

function _모델번호요약(제품모델목록, 건) {
  const rows = 제품모델목록.length ? 제품모델목록 : [{
    모델명: 건['제품명'],
    세부품명번호: 건['세부품명번호'],
    물품식별번호: 건['물품식별번호'],
  }];
  return rows.map((m, idx) => {
    const 모델명 = String(m['모델명'] || m.모델명 || `모델 ${idx + 1}`).trim();
    const 세부 = _v(m['세부품명번호'] || m.세부품명번호);
    const 물품 = _v(m['물품식별번호'] || m.물품식별번호);
    return `${idx + 1}. ${모델명} / ${세부} / ${물품}`;
  }).join('\n');
}

function _제품수표시(제품모델목록, 건) {
  if (제품모델목록.length) return `${제품모델목록.length}개`;
  const 제품수 = String(건['제품수'] || '').trim();
  if (!제품수) return '1개';
  return 제품수.endsWith('개') ? 제품수 : `${제품수}개`;
}

function _모델번호연결확인(제품모델목록, 건) {
  const rows = 제품모델목록.length ? 제품모델목록 : [{
    세부품명번호: 건['세부품명번호'],
    물품식별번호: 건['물품식별번호'],
  }];
  const 누락 = rows.filter(m => {
    const 세부 = String(m['세부품명번호'] || m.세부품명번호 || '').trim();
    const 물품 = String(m['물품식별번호'] || m.물품식별번호 || '').trim();
    return !세부 || !물품;
  }).length;
  return 누락 ? `확인 필요 (${누락}건 번호 누락)` : `연결 확인 (${rows.length}건)`;
}

// 빈 값은 "(미기재)"로 — 증적 문서이므로 누락을 명시
function _v(val) {
  const s = String(val == null ? '' : val).trim();
  return s === '' ? '(미기재)' : s;
}

function _명세섹션(body, 제목) {
  body.appendParagraph('');
  body.appendParagraph(제목).setHeading(DocumentApp.ParagraphHeading.HEADING2);
}

function _명세표(body, 행들) {
  const t = body.appendTable(행들);
  for (let r = 0; r < t.getNumRows(); r++) {
    const row = t.getRow(r);
    row.getCell(0).setBackgroundColor('#f1f3f4').setWidth(160);
    for (let c = 0; c < row.getNumCells(); c++) {
      row.getCell(c).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
      if (c === 0) {
        row.getCell(c).editAsText().setBold(true);
        _셀문단정렬(row.getCell(c), DocumentApp.HorizontalAlignment.CENTER);
      }
    }
  }
  return t;
}

function _명세표헤더(t) {
  const h = t.getRow(0);
  for (let c = 0; c < h.getNumCells(); c++) {
    h.getCell(c).setBackgroundColor(보고서_표헤더색);
    h.getCell(c).editAsText().setForegroundColor(보고서_표헤더글자색).setBold(true);
    _셀문단정렬(h.getCell(c), DocumentApp.HorizontalAlignment.CENTER);
  }
}

// ═════════════════════════════════════════════════════════
// 11. 이미지 표시본(축소) 생성 — 원본은 건드리지 않음
// ═════════════════════════════════════════════════════════
//
// 원본 20MB 이미지를 Docs에 그대로 넣으면 문서가 무거워지므로,
// 폭 1000px 표시본을 만들어 삽입합니다. 원본 파일은 Drive에 그대로 보존됩니다.
//
// 방법: Drive 썸네일 API (getThumbnail은 작아서, URL 파라미터로 폭 지정)
//       UrlFetchApp으로 원하는 폭의 썸네일을 직접 받아옵니다. (무료, 추가 변환 없음)

const 표시본_폭 = 1000;  // px (필요시 조정)

function _표시본Blob생성(파일ID) {
  // 1순위: Drive 썸네일 URL에서 원하는 폭으로 직접 가져오기
  try {
    const token = ScriptApp.getOAuthToken();
    // Drive API v3 thumbnailLink 조회
    const metaRes = UrlFetchApp.fetch(
      `https://www.googleapis.com/drive/v3/files/${파일ID}?fields=thumbnailLink,name,mimeType`,
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    const meta = JSON.parse(metaRes.getContentText());

    if (meta.thumbnailLink) {
      // thumbnailLink 끝의 =s220 같은 크기 파라미터를 원하는 폭으로 교체
      const 큰썸네일 = meta.thumbnailLink.replace(/=s\d+$/, '=s' + 표시본_폭)
                                        .replace(/=w\d+-h\d+/, '=w' + 표시본_폭);
      const imgRes = UrlFetchApp.fetch(큰썸네일, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      });
      if (imgRes.getResponseCode() === 200) {
        return imgRes.getBlob().setName((meta.name || 'thumb') + '_표시본.png');
      }
    }
  } catch (e) {
    Logger.log('썸네일 생성 실패, 원본으로 폴백: ' + e.message);
  }

  // 2순위(폴백): 원본 blob 그대로 (썸네일 불가한 형식 등)
  // appendImage가 자체적으로 표시 폭을 480pt로 줄이므로 문서 표시엔 문제 없으나,
  // 파일 자체는 큼. 원본이 매우 크면 경고 로그만 남김.
  const 원본 = DriveApp.getFileById(파일ID);
  const blob = 원본.getBlob();
  const 크기MB = blob.getBytes().length / (1024 * 1024);
  if (크기MB > 10) {
    Logger.log(`경고: ${원본.getName()} 표시본 생성 실패, 원본 ${크기MB.toFixed(1)}MB 삽입`);
  }
  return blob;
}

/**
 * (선택) 구조도 원본들을 한 폴더에 모아 "원본 보관 폴더" 만들기
 * NAS Cloud Sync가 이 폴더 하나만 동기화하면 되도록 정리하는 헬퍼.
 * 접수번호별 하위 폴더에 원본을 복사합니다. (원본 위치는 유지, 복사본 생성)
 */
function 구조도원본정리(접수번호) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 대장 = ss.getSheetByName(SHEET.접수대장);
  const D = 대장.getDataRange().getValues();
  const H = D[0];
  const 행 = D.slice(1).find(r => r[H.indexOf('접수번호')] === 접수번호);
  if (!행) throw new Error('접수번호 없음: ' + 접수번호);

  const 원본ID들 = String(행[H.indexOf('구조도파일명')] || '')
    .split(/[,\n]+/).map(s => _드라이브ID추출(s.trim())).filter(Boolean);
  if (!원본ID들.length) return { ok: false, msg: '구조도 없음' };

  // 보관 루트 폴더 (CONFIG.드라이브폴더ID 또는 새로 생성)
  let 루트;
  if (CONFIG.드라이브폴더ID) {
    루트 = DriveApp.getFolderById(CONFIG.드라이브폴더ID);
  } else {
    const 이름 = '심사증적_원본보관';
    const 기존 = DriveApp.getFoldersByName(이름);
    루트 = 기존.hasNext() ? 기존.next() : DriveApp.createFolder(이름);
  }

  // 접수번호 하위 폴더
  let 하위;
  const 하위검색 = 루트.getFoldersByName(접수번호);
  하위 = 하위검색.hasNext() ? 하위검색.next() : 루트.createFolder(접수번호);

  let 복사수 = 0;
  원본ID들.forEach(id => {
    try {
      const f = DriveApp.getFileById(id);
      // 이미 있으면 건너뛰기
      const 중복 = 하위.getFilesByName(f.getName());
      if (!중복.hasNext()) {
        f.makeCopy(f.getName(), 하위);
        복사수++;
      }
    } catch (e) {
      Logger.log('복사 실패: ' + id + ' / ' + e.message);
    }
  });

  return { ok: true, folder: 하위.getUrl(), copied: 복사수 };
}

/** 보관루트/접수번호 폴더 준비 */
function _보관폴더준비(접수번호) {
  let 루트;
  if (CONFIG.드라이브폴더ID) {
    루트 = DriveApp.getFolderById(CONFIG.드라이브폴더ID);
  } else {
    const 이름 = '심사증적_원본보관';
    const 기존 = DriveApp.getFoldersByName(이름);
    루트 = 기존.hasNext() ? 기존.next() : DriveApp.createFolder(이름);
  }
  const 하위 = 루트.getFoldersByName(접수번호);
  return 하위.hasNext() ? 하위.next() : 루트.createFolder(접수번호);
}

/** 생성된 보고서 파일을 보관루트/접수번호 폴더로 옮김 (동일 이름 이전 파일은 휴지통 처리) */
function _보고서를접수번호폴더로저장_(file, 접수번호) {
  try {
    const 폴더 = _보관폴더준비(접수번호);
    const 기존 = 폴더.getFilesByName(file.getName());
    while (기존.hasNext()) {
      const f = 기존.next();
      if (f.getId() !== file.getId()) f.setTrashed(true);
    }
    const 부모목록 = file.getParents();
    폴더.addFile(file);
    while (부모목록.hasNext()) {
      const 부모 = 부모목록.next();
      if (부모.getId() !== 폴더.getId()) 부모.removeFile(file);
    }
  } catch (e) {
    Logger.log('보고서 파일 폴더 이동 실패: ' + e.message);
  }
}

