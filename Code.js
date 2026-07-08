/**
 * ============================================================
 *  인공지능 제품 기술심사 관리 시스템 — Google Apps Script
 *  대상: Google Sheets (스프레드시트에 이 스크립트를 붙여넣기)
 * ============================================================
 *
 *  [시트 구성]
 *  1. 접수대장          — 심사 건별 기본정보 + 진행상태 (전체 데이터베이스)
 *  2. 인공지능제품모델  — 세부품명번호·물품식별번호 반복 목록 (모델 여러 개인 경우)
 *  3. 인공지능기능상세  — 기능별 업무설명 + 구현방식 세부 정보
 *  4. 일정관리          — 스케줄/배정 확인용 메인 뷰 (최소 필드만)
 *  5. 파싱로그          — 자동 파싱 이력
 *  6. 인공지능심사체크결과 — 심사폼 판정 결과 (원본과 완전 분리)
 *
 *  [사용 방법]
 *  1. Google Sheets 새 파일 생성
 *  2. 확장 프로그램 > Apps Script > 이 코드 전체 붙여넣기
 *  3. 저장 후 '초기설정실행' 함수 한 번 실행 (시트 구조 자동 생성)
 *  4. 이후 엑셀 수신 시 '엑셀파싱등록' 함수 실행
 * ============================================================
 */

// ─────────────────────────────────────────────
// 0. 설정값 (환경에 맞게 수정)
// ─────────────────────────────────────────────
const CONFIG = {
  // 담당자 목록 (라운드로빈 자동 배분)
  담당자목록: ['홍길동', '김심사', '이검토', '박분석'],

  // 기본 심사 기간 (일)
  기본심사기간: 14,

  // 알림 이메일 (담당자 배분 시 발송, 빈 문자열이면 발송 안 함)
  알림이메일: '',   // 예: 'manager@yourorg.kr'

  // Google Drive 폴더 ID (엑셀 파일을 업로드할 폴더, 빈 문자열이면 루트)
  드라이브폴더ID: '',

  // ── Google Chat Webhook ──────────────────────
  // Space별 Webhook URL을 설정합니다.
  // Google Chat > 스페이스 > 앱 및 통합 > Webhook > URL 복사
  //
  // 공통 알림 채널 (전체 접수 알림)
  챗_공통Webhook: '',
  // 예: 'https://chat.googleapis.com/v1/spaces/XXXXX/messages?key=...&token=...'

  // 담당자별 개인 Webhook (담당자 이름을 키로 설정, 없으면 공통 채널로 발송)
  // DM Webhook: Google Chat > 앱 > Incoming Webhook 추가 > 개인 스페이스에 설치
  챗_담당자Webhook: {
    '홍길동': '',   // 개인 DM Webhook URL
    '김심사': '',
    '이검토': '',
    '박분석': '',
  },

  // ── 기술심사보고서 — 기관 정보 (한 번 설정하면 모든 보고서에 적용) ──
  보고서: {
    작성기관: '한국정보통신기술협회(TTA)',
    문서번호접두: 'TTA-AI심사',          // 예: TTA-AI심사-2026-0003
    심사근거: '「인공지능 발전과 신뢰 기반 조성 등에 관한 기본법」제16조제3항, ｢인공지능제품·서비스 확인 절차 운영에 관한 고시｣'                         // 예: 'OOO법 제O조, OOO지침 제O조' (없으면 빈칸 유지)
    심사방법: '제출 기술자료 검토 및 항목별 적합성 평가',
    보안등급: '대외제한',                 // 표지 표기
    책임자직위: '심사책임자',
  },

  // ── 심사 폼 웹앱 URL (배포 후 발급된 URL을 여기에 붙여넣기) ──
  심사폼URL: 'https://script.google.com/a/macros/tta.or.kr/s/AKfycby8QNKlBDDUSNEjMWsDXyLGpzeDPiN056q4hJFP0FUtj1G2fdT4u0SBJjFQbrIVEZ1yxQ/exec'
};

// 시트 이름 상수
const SHEET = {
  접수대장: '접수대장',
  제품모델: '인공지능제품모델',        // 신규 — 세부품명번호/물품식별번호 반복행
  AI기능상세: '인공지능기능상세',      // 표시명만 변경 (코드상 키는 호환 유지)
  일정관리: '일정관리',
  로그: '파싱로그',
  체크결과: '인공지능심사체크결과',    // 체크 폼 결과 저장 (원본과 완전 분리)
  기능확인: '인공지능기능확인',        // 심사폼 기능명세 표의 확인·비고 저장
};

// 심사 체크리스트 항목 정의 (엑셀 데이터로 자동 채워질 항목들)
// id: 고유키 / 항목: 질문 / 참조: 접수대장·AI기능상세에서 끌어올 값의 출처
const 체크항목정의 = [
  { id: 'C1', no: '1.3', 항목: '제품 제공 형태가 신청 분류와 일치하는가', 참조: '제공형태+제품분류' },
  { id: 'C2', no: '1.5', 항목: '제품 환경 정보(서버/OS/런타임)가 구체적으로 기재되었는가', 참조: '제품명' },
  { id: 'C3', no: '4.1', 항목: 'AI 기능명과 역할이 모든 기능에 대해 기재되었는가', 참조: '인공지능기능수' },
  { id: 'C4', no: '4.2', 항목: 'AI 기능별 입력·출력 데이터 유형이 명확히 기재되었는가', 참조: '인공지능기능수' },
  { id: 'C5', no: '5.1', 항목: '기능별 AI 구현 방식(A~E)이 적절히 분류되었는가', 참조: '인공지능기능수' },
  { id: 'C6', no: '5.2', 항목: '연산 자원 및 실행 환경이 구현 방식에 부합하는가', 참조: '인공지능기능수' },
  { id: 'C7', no: '5.3', 항목: '외부 API 활용 기능의 데이터 처리 경로가 적절히 설명되었는가', 참조: '외부API모델' },
  { id: 'C8', no: '5.4', 항목: '자체 학습/파인튜닝 모델의 학습 데이터 사양이 기재되었는가', 참조: '인공지능기능수' },
  { id: 'C9', no: '2.1', 항목: '신청기업 정보(사업자번호, 대표자, 연락처)가 완전한가', 참조: '기업명' },
  { id: 'C10', no: '3.1', 항목: '세부품명번호·물품식별번호가 정상 기재되었는가', 참조: '세부품명번호+물품식별번호' },
  { id: 'C11', no: '6.1', 항목: '기존 인증·시험 결과 관련 정보가 확인되었는가', 참조: '비고' },
  { id: 'C12', no: '0.0', 항목: '종합 — 기술 명세가 심사 진행에 충분한가', 참조: '종합' },
];


// ─────────────────────────────────────────────
// 1. 초기 설정 — 시트 구조 자동 생성
// ─────────────────────────────────────────────

// 시트별 표준 헤더 정의 (초기 생성 + 컬럼 마이그레이션에 공용)
const 시트헤더정의 = {
  [SHEET.접수대장]: [
    // ── TTA 관리 (자동생성) ──────────────────────
    '접수번호',
    '신청일',        // 기업이 KOSA에 신청한 날 (15일 기산점, 신청서 원본값)
    '심사접수일',    // TTA가 KOSA로부터 건을 인수한 날 (파싱 시 자동 기록)
    '상태',
    // ── 1. 기업 정보 ─────────────────────────────
    '기업명', '사업자번호', '대표자', '소재지',
    '담당자명', '연락처', '이메일',
    // ── 2. 제품·서비스 정보 ───────────────────────
    '제품명', '제공형태', '제품분류',
    '개요', '인공지능적용목적', '인공지능적용범위',
    '구조도파일명',
    // ── 4. 정부문서 제출 ──────────────────────────
    '명세서작성방식', '명세서파일명', '기타제출서류파일명',
    // ── 5. 기존 인증 ──────────────────────────────
    '보유인증', '인증서파일명',
    // ── 6~7. 동의서·특기사항 ─────────────────────
    '열람이용동의여부', '개인정보수집이용동의여부', '개인정보3자제공동의여부',
    '특기사항',
    // ── TTA 심사 관리 (내부 운영) ────────────────
    '인공지능기능수', '인공지능기능명(요약)', '구현방식(요약)',
    '담당심사원', '심사착수일', '심사마감일', '비고',
  ],
  [SHEET.제품모델]: [
    '접수번호', '연번', '모델명', '세부품명번호', '물품식별번호',
  ],
  [SHEET.AI기능상세]: [
    '접수번호', '기능번호', '기능명',
    '인공지능역할', '입력', '출력', '설명서참조위치',
    '구현방식', '연산자원', '실행환경',
    '학습데이터사양', '재현환경라이선싱',
    'BaseModel', '튜닝방법',
    '외부API모델',
    '타겟HW_OS', '배포포맷정밀도',
    '혼합구성설명',
    '입력데이터설명', '출력데이터설명',
    '기타참고자료파일명', '데이터흐름도파일명',
  ],
  [SHEET.일정관리]: [
    // ── 일정·기한 ─────────────────────────────────
    '순번',          // 표시용 일련번호 (자동)
    '접수번호',
    '심사링크',      // 심사폼 바로가기 (이 시트가 워크플로우 허브)
    '신청일',        // 기산일 (신청서 원본)
    '심사접수일',    // TTA 인수일
    '마감예정일',    // 신청일 + 15일 (수식 자동 생성)
    '상태',          // 대기 / 심사중 / 보완 / 완료
    '담당심사원',
    // ── 신청기업 ──────────────────────────────────
    '기업명', '담당자명', '연락처', '이메일',
    // ── 제품·서비스 핵심 정보 ─────────────────────
    '제품명', '개요', '제공형태', '제품분류',
    '인공지능적용목적', '인공지능적용범위',
    // ── 제출 현황 ─────────────────────────────────
    '명세서작성방식', '기타제출서류여부', '보유인증',
    // ── 심사 참고 ─────────────────────────────────
    '인공지능기능수',
  ],
  [SHEET.로그]: [
    '일시', '파일명', '접수번호', '결과', '메시지',
  ],
  [SHEET.체크결과]: [
    '접수번호', '항목ID', '항목번호', '체크항목',
    '판정', '사유', '심사원', '검토일시',
  ],
  [SHEET.기능확인]: [
    '접수번호', 'ID', '기능번호', '기능명', '확인', '비고', '심사원', '검토일시',
  ],
};

// 구버전 시트명 → 신버전 시트명 (탭 이름만 바꿔서 기존 데이터 그대로 보존)
const 시트이름마이그레이션맵 = {
  'AI기능상세': SHEET.AI기능상세,
  '심사체크결과': SHEET.체크결과,
};

/** 구버전 이름의 시트가 있고 신버전 이름이 아직 없으면 탭 이름만 변경 (데이터 보존) */
function _시트이름마이그레이션(ss) {
  Object.keys(시트이름마이그레이션맵).forEach(옛이름 => {
    const 새이름 = 시트이름마이그레이션맵[옛이름];
    if (옛이름 === 새이름) return;
    const 옛시트 = ss.getSheetByName(옛이름);
    const 새시트 = ss.getSheetByName(새이름);
    if (옛시트 && !새시트) 옛시트.setName(새이름);
  });
}

/**
 * 기존 시트에 새 표준 컬럼이 없으면 맨 뒤에 추가 (데이터·순서 보존)
 * — 행 쓰기가 헤더 이름 기반이라 컬럼이 뒤에 붙어도 안전합니다.
 * 이름이 바뀐 컬럼(역할→인공지능역할 등)은 기존 값도 새 컬럼으로 복사합니다.
 */
const 컬럼값이관맵 = {
  [SHEET.일정관리]: {
    '마감일': '마감예정일',
    '착수일': '심사접수일',
  },
  [SHEET.AI기능상세]: {
    '역할': '인공지능역할',
    '입력데이터': '입력데이터설명',
    '출력데이터': '출력데이터설명',
    '학습데이터': '학습데이터사양',
    '추론런타임': '배포포맷정밀도',
  },
  [SHEET.접수대장]: {
    'AI제품분류': '제품분류',
    '주소': '소재지',
    'AI기능수': '인공지능기능수',
    'AI기능명(요약)': '인공지능기능명(요약)',
    '접수일': '심사접수일',
    '신청일자': '신청일',
  },
};

function 헤더마이그레이션() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let 추가내역 = [];
  Object.keys(시트헤더정의).forEach(이름 => {
    const 시트 = ss.getSheetByName(이름);
    if (!시트 || 시트.getLastRow() < 1) return;
    let 기존 = 시트.getRange(1, 1, 1, Math.max(1, 시트.getLastColumn()))
      .getValues()[0].map(v => String(v).trim());
    const 누락 = 시트헤더정의[이름].filter(h => !기존.includes(h));
    if (누락.length) {
      const 시작열 = 기존.filter(v => v !== '').length + 1;
      시트.getRange(1, 시작열, 1, 누락.length).setValues([누락]);
      시트.getRange(1, 시작열, 1, 누락.length)
        .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
      추가내역.push(`${이름}: ${누락.join(', ')}`);
      기존 = 시트.getRange(1, 1, 1, 시트.getLastColumn())
        .getValues()[0].map(v => String(v).trim());
    }
    // 이름 바뀐 컬럼의 기존 값 이관 (새 컬럼이 비어있는 행만)
    const 이관 = 컬럼값이관맵[이름];
    if (이관 && 시트.getLastRow() >= 2) {
      Object.keys(이관).forEach(옛 => {
        const iOld = 기존.indexOf(옛);
        const iNew = 기존.indexOf(이관[옛]);
        if (iOld < 0 || iNew < 0) return;
        const n = 시트.getLastRow() - 1;
        const 옛값 = 시트.getRange(2, iOld + 1, n, 1).getValues();
        const 새값 = 시트.getRange(2, iNew + 1, n, 1).getValues();
        let 변경 = false;
        for (let r = 0; r < n; r++) {
          if ((새값[r][0] === '' || 새값[r][0] == null) && 옛값[r][0] !== '') {
            새값[r][0] = 옛값[r][0];
            변경 = true;
          }
        }
        if (변경) 시트.getRange(2, iNew + 1, n, 1).setValues(새값);
      });
    }
  });
  if (추가내역.length) {
    try {
      SpreadsheetApp.getUi().alert('컬럼 마이그레이션 완료:\n\n' + 추가내역.join('\n'));
    } catch (e) { /* UI 없는 환경 */ }
  }
  return 추가내역;
}

function 초기설정실행() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 0) 구버전 시트명이면 탭 이름 변경 (기존 데이터 보존)
  _시트이름마이그레이션(ss);

  // 1) 시트 생성 (없을 때만) — 이미 있으면 건드리지 않음
  Object.keys(시트헤더정의).forEach(이름 => {
    _시트초기화(ss, 이름, 시트헤더정의[이름]);
  });

  // 2) 기존 시트에 새 표준 컬럼 자동 추가
  try { 헤더마이그레이션(); } catch (e) { /* UI 없는 환경에서 실행될 수 있음 */ }

  // 3) 컬럼매핑 메타 시트 (엑셀 헤더 별칭을 코드 수정 없이 관리)
  _컬럼매핑시트확보(ss);

  // ── 일정관리 시트 후처리 ─────────────────────────────
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 일정H = 시트헤더정의[SHEET.일정관리];
  const iD마감 = 일정H.indexOf('마감예정일') + 1;
  const iD상태 = 일정H.indexOf('상태') + 1;
  const 끝열문자 = columnLetter(일정H.length);

  // 마감예정일 수식은 _일정관리행추가()에서 행별로 설정
  // (컬럼 전체를 미리 채우면 appendRow가 빈 행을 못 찾아 밀리므로 사전 채움 안 함)

  // 상태 드롭다운
  일정시트.getRange(2, iD상태, 999, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['대기', '심사중', '보완', '완료'], true).build());

  // 조건부서식
  const 마감열문자 = columnLetter(iD마감);
  const 상태열문자 = columnLetter(iD상태);
  const 전체범위 = 일정시트.getRange(`A2:${끝열문자}1000`);

  // 빨강: 마감 초과 & 미완료
  const 빨강 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      `=AND($${마감열문자}2<>"",$${마감열문자}2<TODAY(),$${상태열문자}2<>"완료")`
    )
    .setBackground('#FDECEA').setFontColor('#C5221F')
    .setRanges([전체범위]).build();

  // 주황: 3일 이내 마감 & 미완료
  const 주황 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      `=AND($${마감열문자}2<>"",$${마감열문자}2>=TODAY(),$${마감열문자}2-TODAY()<=3,$${상태열문자}2<>"완료")`
    )
    .setBackground('#FEF7E0').setFontColor('#E37400')
    .setRanges([전체범위]).build();

  일정시트.setConditionalFormatRules([빨강, 주황]);

  // 열 너비 (픽셀) — 22개 컬럼 (순번·심사링크 포함)
  const 일정너비 = [45, 100, 60, 90, 90, 90, 70, 90, 150, 80, 110, 170, 160, 220, 110, 80, 250, 250, 90, 70, 130, 55];
  일정너비.forEach((w, i) => { if (i < 일정H.length) 일정시트.setColumnWidth(i + 1, w); });
  일정시트.setFrozenRows(1);
  일정시트.setFrozenColumns(2);  // 순번 + 접수번호 고정

  // ── 접수대장 상태 드롭다운 ────────────────────────
  // 하드코딩 금지 — 실제 헤더에서 '상태' 위치를 찾아서 설정
  const 대장시트 = ss.getSheetByName(SHEET.접수대장);
  const 대장H = 대장시트.getRange(1, 1, 1, 대장시트.getLastColumn()).getValues()[0];
  const 대장상태열 = 대장H.indexOf('상태') + 1;
  if (대장상태열 > 0) {
    const 대장상태범위 = 대장시트.getRange(2, 대장상태열, 999, 1);
    const 대장상태규칙 = SpreadsheetApp.newDataValidation()
      .requireValueInList(['접수', '심사중', '보완요청', '완료', '반려'], true)
      .build();
    대장상태범위.setDataValidation(대장상태규칙);
  }

  SpreadsheetApp.getUi().alert('초기설정 완료! 시트 7개가 생성되었습니다.');
}

/** 열 번호(1-based) → 알파벳 열 문자 (1→A, 27→AA) */
function columnLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function _시트초기화(ss, 이름, 헤더배열) {
  let sheet = ss.getSheetByName(이름);
  if (!sheet) {
    sheet = ss.insertSheet(이름);
  }

  // 이미 헤더(1행 1열)에 값이 있으면 = 테이블이 만들어진 시트일 수 있으므로
  // 내용을 지우지 않고 그대로 둡니다. (방식 A: 수동 테이블 보존)
  const 기존헤더 = sheet.getRange(1, 1).getValue();
  if (기존헤더 !== '' && 기존헤더 !== null) {
    return sheet;  // 이미 세팅됨 — 건드리지 않음
  }

  // 빈 시트일 때만 헤더 작성
  sheet.getRange(1, 1, 1, 헤더배열.length).setValues([헤더배열]);
  sheet.getRange(1, 1, 1, 헤더배열.length)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

// 강제 재생성용 (정말 처음부터 다시 만들 때만 메뉴에서 호출)
/**
 * 헤더만 재설정 — 데이터는 유지하고 1행 헤더만 표준으로 덮어씀
 * 기존 시트 컬럼 순서가 틀어진 경우 사용 (데이터 날아가지 않음)
 * ※ 단, 컬럼 순서 자체를 바꾸지는 않음 — 헤더 이름만 교정
 * 컬럼 순서까지 맞추려면 '시트 재생성'을 사용 (데이터 소실 주의)
 */
function 헤더만재설정() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '헤더 재설정',
    '각 시트의 1행(헤더)을 표준 컬럼명으로 덮어씁니다.\n데이터 행은 건드리지 않습니다.\n\n계속하시겠습니까?',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  Object.keys(시트헤더정의).forEach(이름 => {
    const 시트 = ss.getSheetByName(이름);
    if (!시트) return;
    const 헤더 = 시트헤더정의[이름];
    시트.getRange(1, 1, 1, 헤더.length).setValues([헤더]);
    시트.getRange(1, 1, 1, 헤더.length)
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  });

  // 드롭다운·수식도 헤더 기준으로 재설정
  try { 초기설정실행(); } catch (e) { /* 이미 시트 있으면 무시 */ }
  ui.alert('헤더 재설정 완료. 드롭다운·수식도 갱신했습니다.');
}

/**
 * 접수대장·일정관리 시트를 완전히 비우고 표준 헤더로 재생성
 * ⚠️ 기존 데이터가 모두 삭제됩니다. 테스트/리셋 용도로만 사용하세요.
 */
function 시트재생성_데이터초기화() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '⚠️ 데이터 초기화',
    '모든 시트의 데이터가 삭제되고 표준 헤더로 재생성됩니다.\n되돌릴 수 없습니다.\n\n정말 진행하시겠습니까?',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  Object.keys(시트헤더정의).forEach(이름 => _시트강제초기화(ss, 이름, 시트헤더정의[이름]));
  초기설정실행();
  ui.alert('재생성 완료.');
}

function _시트강제초기화(ss, 이름, 헤더배열) {
  let sheet = ss.getSheetByName(이름);
  if (!sheet) { sheet = ss.insertSheet(이름); }
  else { sheet.clear(); }
  sheet.getRange(1, 1, 1, 헤더배열.length).setValues([헤더배열]);
  sheet.getRange(1, 1, 1, 헤더배열.length)
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

// 테이블 친화 행 추가
// appendRow는 시트 맨 끝에 붙어 테이블 범위 밖으로 나갈 수 있으므로,
// 마지막 데이터 행 바로 아래(= 테이블 확장 영역)에 값을 써서 테이블이 자동 흡수하게 합니다.
function _테이블행추가(sheet, 값배열) {
  const 마지막 = sheet.getLastRow();
  const 대상행 = 마지막 + 1;
  sheet.getRange(대상행, 1, 1, 값배열.length).setValues([값배열]);
  return 대상행;
}


// ─────────────────────────────────────────────
// 1-1. 컬럼매핑 메타 시트 — 엑셀 헤더 별칭을 코드 수정 없이 관리
// ─────────────────────────────────────────────
// KOSA 엑셀의 컬럼명이 바뀌면 코드가 아니라 이 시트의 '별칭' 칸만 고치면 됩니다.
// A열: 내부키 (코드가 쓰는 표준 이름)  B열: 별칭 (쉼표로 구분, 부분일치 허용)

const 컬럼매핑시트명 = '컬럼매핑';

function _컬럼매핑시트확보(ss) {
  let 시트 = ss.getSheetByName(컬럼매핑시트명);
  if (시트) return 시트;
  시트 = ss.insertSheet(컬럼매핑시트명);
  const 기본값 = [
    ['내부키', '별칭 (쉼표 구분 — 엑셀 헤더가 바뀌면 여기만 수정)'],
    ['접수번호', '접수번호, 접수 번호, 관리번호, 심사번호, 신청번호'],
    ['제품명', '제품 또는 서비스명, 제품·서비스 모델명, 제품명, 서비스명'],
    ['제공형태', '제공 형태, 제공형태'],
    ['제품분류', '인공지능 제품·서비스 분류, 제품분류'],
    ['기업명', '상호(사업자명), 기업명'],
    ['사업자번호', '법인등록번호, 사업자등록번호, 사업자번호'],
    ['대표자', '대표자 성명, 대표자명, 대표자'],
    ['소재지', '소재지'],
    ['담당자명', '업무담당자, 담당자명, 담당자'],
    ['연락처', '담당자 전화번호, 연락처, 전화번호'],
    ['이메일', '담당자 이메일주소, 이메일, email'],
    ['개요', '개요'],
    ['인공지능적용목적', '인공지능 적용 목적'],
    ['인공지능적용범위', '인공지능 적용 범위'],
    ['구조도파일명', '제품 구조도, 인공지능 제품·서비스 구조도'],
    ['보유인증', '보유 인증'],
    ['특기사항', '특기사항'],
    ['신청일', '신청일자, 신청일'],
    ['기능명', '기능명'],
    ['인공지능역할', '인공지능 역할, 인공지능역할'],
    ['구현방식', '인공지능 구현 방식, 구현방식, 구현 방식'],
    ['연산자원', '인공지능 연산 자원, 연산자원, 연산 자원'],
    ['실행환경', '인공지능 실행 환경 세부정보, 실행환경, 실행 환경'],
    ['학습데이터사양', '학습 데이터 사양, 학습데이터사양'],
    ['재현환경라이선싱', '재현환경·라이선싱 알고리즘, 재현환경라이선싱'],
    ['배포포맷정밀도', '온디바이스 배포 포맷 및 정밀도, 배포포맷정밀도'],
    ['모델명', '제품·서비스 모델명, 모델명'],
  ];
  시트.getRange(1, 1, 기본값.length, 2).setValues(기본값);
  시트.getRange(1, 1, 1, 2)
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  시트.setFrozenRows(1);
  시트.setColumnWidth(1, 160);
  시트.setColumnWidth(2, 520);
  return 시트;
}

/** 컬럼매핑 시트 → { 내부키: [별칭, ...] } (실행 1회 캐시) */
let _별칭캐시 = null;
function _커스텀별칭() {
  if (_별칭캐시) return _별칭캐시;
  _별칭캐시 = {};
  try {
    const 시트 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(컬럼매핑시트명);
    if (시트 && 시트.getLastRow() >= 2) {
      const D = 시트.getRange(2, 1, 시트.getLastRow() - 1, 2).getValues();
      D.forEach(([키, 별칭들]) => {
        const k = String(키 || '').trim();
        if (!k) return;
        _별칭캐시[k] = String(별칭들 || '').split(',').map(s => s.trim()).filter(Boolean);
      });
    }
  } catch (e) { /* 매핑 시트 없으면 코드 내장 별칭만 사용 */ }
  return _별칭캐시;
}

/** 내부키의 커스텀 별칭 + 코드 기본 별칭을 합쳐 반환 */
function _별칭합치기(내부키, 기본후보들) {
  const 커스텀 = _커스텀별칭()[내부키] || [];
  return [...커스텀, ...기본후보들];
}


// ─────────────────────────────────────────────
// 2. 메인 — 엑셀 파일 파싱 후 Sheets 등록
// ─────────────────────────────────────────────

/**
 * 사용법:
 *  - Google Drive에 업로드된 엑셀 파일의 URL 또는 파일 ID를 붙여넣으면 자동 파싱
 *  - 또는 직접 구조화된 데이터를 아래 수동입력 함수로 등록
 */
function 엑셀파싱등록() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '엑셀 파일 등록',
    '엑셀 파일 ID 또는 공유 URL을 입력하세요.\n' +
    '(Drive 링크, Sheets 링크 모두 가능합니다.)\n\n' +
    '※ 주의: 스크립트를 실행하는 구글 계정이 해당 파일에 접근할 권한이 있어야 합니다.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const 입력값 = response.getResponseText().trim();
  const 파일ID = _드라이브ID추출(입력값);

  try {
    const file = DriveApp.getFileById(파일ID);
    const mimeType = file.getMimeType();

    // 파일이 원본 엑셀이 아니라 구글 시트로 완전히 변환된 상태인지 사전 체크
    if (mimeType === MimeType.GOOGLE_SHEETS) {
      ui.alert(
        '⛔ 오류: 해당 링크는 이미 구글 시트 포맷으로 변환된 파일입니다.\n\n' +
        '현재 파싱 코드는 원본 "엑셀 파일(.xlsx)" 형태를 요구합니다.\n' +
        '해당 파일을 엑셀로 다운로드한 뒤, 구글 드라이브에 다시 올려서 그 링크를 넣어주세요.'
      );
      return;
    }

    const blob = file.getBlob();
    _엑셀파싱처리(blob, file.getName());
  } catch (e) {
    ui.alert('파일을 찾을 수 없습니다: ' + e.message);
  }
}

function _드라이브ID추출(입력) {
  const match = 입력.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : 입력;
}

/**
 * 엑셀 Blob → 파싱 → Sheets 등록
 *
 * [엑셀 파일 구조 가정]
 * 외부 시스템이 제공하는 심사 기초자료 엑셀은 아래 두 가지 형태 중 하나:
 *
 * [형태 A] 항목-값 2열 구조 (세로형)
 *   A열: 항목명  B열: 값
 *   제품 또는 서비스명 | Mediview AI ...
 *   기업명            | 주식회사 ...
 *
 * [형태 B] 1행 헤더 + 데이터행 (가로형, 수백 건 일괄)
 *   A1: 제품명  B1: 기업명  C1: ...
 *   A2: ...     B2: ...
 *
 * 아래 코드는 형태 A (건별 세로형) 기준.
 * 형태 B를 받는다면 _파싱_가로형() 함수를 대신 호출하세요.
 */
function _엑셀파싱처리(blob, 파일명) {
  const 제목 = '__임시파싱__' + new Date().getTime();
  let 임시파일ID;

  try {
    임시파일ID = _엑셀을Sheets로변환(blob, 제목);
    const 임시SS = SpreadsheetApp.openById(임시파일ID);
    const 시트들 = 임시SS.getSheets();

    // 탭 분류: 기능상세 탭(이름에 '기능' 포함), 제품모델 탭(이름에 '모델' 포함), 나머지 기본정보 탭
    let 기본탭 = null;
    let 기능탭 = null;
    let 모델탭 = null;
    시트들.forEach(sh => {
      const 이름 = sh.getName();
      if (/기능|function|feature/i.test(이름)) {
        if (!기능탭) 기능탭 = sh;
      } else if (/모델|model/i.test(이름)) {
        if (!모델탭) 모델탭 = sh;
      } else if (/안내|guide|readme/i.test(이름)) {
        // 작성안내 탭은 무시
      } else {
        if (!기본탭) 기본탭 = sh;
      }
    });
    if (!기본탭) 기본탭 = 시트들[0];  // 못 찾으면 첫 탭

    // 1) 기본정보 탭 파싱 → 접수대장 등록 (제품명·접수번호 → 접수번호 매핑 보관)
    const 기본데이터 = 기본탭.getDataRange().getValues();
    const 제품명별접수번호 = {};

    const 세로형여부 = _세로형감지(기본데이터);
    const 건목록 = 세로형여부 ? [_파싱_세로형(기본데이터)] : _파싱_가로형(기본데이터);
    건목록.forEach(건 => {
      const 접수번호 = _Sheets에등록(건, 파일명);
      if (건.제품명) 제품명별접수번호[String(건.제품명).trim()] = 접수번호;
      // 접수번호 자기참조도 등록 (기능탭에 접수번호 컬럼이 있으면 직접 매칭)
      제품명별접수번호['__접수__' + 접수번호] = 접수번호;
    });

    // 2) 기능상세 탭 파싱 → 인공지능기능상세 등록 (접수번호 우선, 없으면 제품명으로 연결)
    if (기능탭) {
      const 기능데이터 = 기능탭.getDataRange().getValues();
      _기능탭파싱등록(기능데이터, 제품명별접수번호);
    }

    // 3) 제품모델 탭 파싱 → 인공지능제품모델 등록 (세부품명번호가 여러 건인 경우)
    if (모델탭) {
      const 모델데이터 = 모델탭.getDataRange().getValues();
      _제품모델탭파싱등록(모델데이터, 제품명별접수번호);
    }

  } finally {
    if (임시파일ID) {
      try { DriveApp.getFileById(임시파일ID).setTrashed(true); } catch(e) {}
    }
  }
}

/**
 * 기능상세 탭(가로형: 1행 헤더 + 기능 N행) 파싱 → AI기능상세 시트 등록
 * 각 기능 행의 '제품명'으로 접수번호를 찾아 연결합니다.
 */
function _기능탭파싱등록(데이터, 제품명별접수번호) {
  if (데이터.length < 2) return;
  const 헤더 = 데이터[0].map(h => String(h).trim());

  // 헤더 인덱스 찾기 (정확일치 우선 → 부분일치 폴백 + 컬럼매핑 시트 커스텀 별칭)
  // 정확일치를 먼저 시도해 '입력'이 '입력데이터설명'을 잘못 잡는 충돌을 방지
  const idx = (내부키, 후보들) => {
    const 후보 = _별칭합치기(내부키, 후보들);
    // 1차: 완전 일치
    for (const c of 후보) {
      const i = 헤더.findIndex(h => h === c);
      if (i >= 0) return i;
    }
    // 2차: 부분 일치 (완전 일치가 없을 때만)
    for (const c of 후보) {
      const i = 헤더.findIndex(h => h.includes(c));
      if (i >= 0) return i;
    }
    return -1;
  };
  const I = {
    접수번호: idx('접수번호', ['접수번호', '접수 번호', '관리번호', '심사번호']),
    제품명: idx('제품명', ['제품명', '제품 또는 서비스명', '서비스명']),
    기능번호: idx('기능번호', ['기능번호', '번호']),
    기능명: idx('기능명', ['기능명']),
    역할: idx('인공지능역할', ['인공지능역할', '인공지능 역할', '역할', 'AI 역할']),
    입력: idx('입력', ['입력']),
    출력: idx('출력', ['출력']),
    설명서참조위치: idx('설명서참조위치', ['설명서참조위치', '설명서 참조 위치']),
    구현방식: idx('구현방식', ['구현방식', '구현 방식']),
    연산자원: idx('연산자원', ['연산자원', '연산 자원']),
    실행환경: idx('실행환경', ['실행환경', '실행 환경']),
    학습데이터사양: idx('학습데이터사양', ['학습데이터사양', '학습 데이터 사양', '학습데이터']),
    재현환경라이선싱: idx('재현환경라이선싱', ['재현환경라이선싱', '재현환경', '라이선싱']),
    BaseModel: idx('BaseModel', ['Base Model', 'BaseModel', '베이스']),
    튜닝방법: idx('튜닝방법', ['튜닝']),
    외부API: idx('외부API모델', ['외부 API', '외부API']),
    타겟HW: idx('타겟HW_OS', ['타겟HW_OS', '타겟', 'HW', 'OS']),
    배포포맷정밀도: idx('배포포맷정밀도', ['배포포맷정밀도', '배포 포맷', '정밀도']),
    혼합구성: idx('혼합구성설명', ['혼합']),
    입력데이터설명: idx('입력데이터설명', ['입력데이터설명']),
    출력데이터설명: idx('출력데이터설명', ['출력데이터설명']),
    기타참고자료: idx('기타참고자료파일명', ['기타참고자료', '기타 참고자료']),
    데이터흐름도: idx('데이터흐름도파일명', ['데이터흐름도', '아카이빙']),
    비고: idx('비고', ['비고']),
  };

  // 접수번호별로 기능 목록 묶기
  const 묶음 = {};  // 접수번호 → [기능객체...]
  for (let r = 1; r < 데이터.length; r++) {
    const 행 = 데이터[r];
    if (행.every(v => v === '' || v == null)) continue;

    // 접수번호 컬럼이 있으면 그걸로 직접 연결 (가장 정확), 없으면 제품명으로
    let 접수번호 = '';
    if (I.접수번호 >= 0) {
      const 직접 = String(행[I.접수번호] ?? '').trim();
      if (직접 && 제품명별접수번호['__접수__' + 직접]) 접수번호 = 직접;
    }
    if (!접수번호 && I.제품명 >= 0) {
      const 제품명 = String(행[I.제품명]).trim();
      접수번호 = 제품명별접수번호[제품명];
    }
    if (!접수번호) continue;  // 연결 안 되는 행은 스킵

    const get = (i) => (i >= 0 ? String(행[i] ?? '').trim() : '');
    const 기능 = {
      기능번호: get(I.기능번호),
      기능명: get(I.기능명),
      인공지능역할: get(I.역할),
      입력: get(I.입력),
      출력: get(I.출력),
      설명서참조위치: get(I.설명서참조위치),
      구현방식: get(I.구현방식),
      연산자원: get(I.연산자원),
      실행환경: get(I.실행환경),
      학습데이터사양: get(I.학습데이터사양),
      재현환경라이선싱: get(I.재현환경라이선싱),
      BaseModel: get(I.BaseModel),
      튜닝방법: get(I.튜닝방법),
      외부API모델: get(I.외부API),
      타겟HW_OS: get(I.타겟HW),
      배포포맷정밀도: get(I.배포포맷정밀도),
      혼합구성설명: get(I.혼합구성),
      입력데이터설명: get(I.입력데이터설명),
      출력데이터설명: get(I.출력데이터설명),
      기타참고자료파일명: get(I.기타참고자료),
      데이터흐름도파일명: get(I.데이터흐름도),
    };
    (묶음[접수번호] = 묶음[접수번호] || []).push(기능);
  }

  // 접수번호별로 기능상세 등록 + 접수대장 기능수 갱신
  Object.keys(묶음).forEach(접수번호 => {
    AI기능상세등록(접수번호, 묶음[접수번호]);
    _접수대장기능수갱신(접수번호, 묶음[접수번호]);
  });
}

/**
 * 제품모델 탭(가로형: 1행 헤더 + 모델 N행) 파싱 → 인공지능제품모델 시트 등록
 * 세부품명번호·물품식별번호가 접수 건당 여러 개인 경우를 위한 반복 목록.
 */
function _제품모델탭파싱등록(데이터, 제품명별접수번호) {
  if (데이터.length < 2) return;
  const 헤더 = 데이터[0].map(h => String(h).trim());
  const idx = (내부키, 후보들) => {
    const 후보 = _별칭합치기(내부키, 후보들);
    for (const c of 후보) {
      const i = 헤더.findIndex(h => h === c);
      if (i >= 0) return i;
    }
    for (const c of 후보) {
      const i = 헤더.findIndex(h => h.includes(c));
      if (i >= 0) return i;
    }
    return -1;
  };
  const I = {
    접수번호: idx('접수번호', ['접수번호', '접수 번호', '관리번호', '심사번호']),
    제품명: idx('제품명', ['제품명', '제품 또는 서비스명', '서비스명']),
    연번: idx('연번', ['연번']),
    모델명: idx('모델명', ['모델명', '제품·서비스 모델명']),
    세부품명번호: idx('세부품명번호', ['세부품명번호']),
    물품식별번호: idx('물품식별번호', ['물품식별번호']),
  };

  const 묶음 = {};
  for (let r = 1; r < 데이터.length; r++) {
    const 행 = 데이터[r];
    if (행.every(v => v === '' || v == null)) continue;

    let 접수번호 = '';
    if (I.접수번호 >= 0) {
      const 직접 = String(행[I.접수번호] ?? '').trim();
      if (직접 && 제품명별접수번호['__접수__' + 직접]) 접수번호 = 직접;
    }
    if (!접수번호 && I.제품명 >= 0) {
      const 제품명 = String(행[I.제품명]).trim();
      접수번호 = 제품명별접수번호[제품명];
    }
    if (!접수번호) continue;

    const get = (i) => (i >= 0 ? String(행[i] ?? '').trim() : '');
    (묶음[접수번호] = 묶음[접수번호] || []).push({
      연번: get(I.연번),
      모델명: get(I.모델명),
      세부품명번호: get(I.세부품명번호),
      물품식별번호: get(I.물품식별번호),
    });
  }

  Object.keys(묶음).forEach(접수번호 => 제품모델등록(접수번호, 묶음[접수번호]));
}

/** 인공지능제품모델 시트에 등록 (재파싱 시 같은 접수번호 기존 행 삭제 후 재기록) */
/**
 * 일정관리 시트에 신규 행 추가 (하이브리드 싱크)
 *  · 접수대장이 원본인 컬럼 → 접수대장 VLOOKUP 수식으로 자동 참조
 *  · 일정관리가 원본인 컬럼(상태·담당심사원) → 직접 입력값 (편집 가능)
 *  · 마감예정일 → 신청일 + 15일 수식
 *
 * VLOOKUP은 접수번호(A열)를 키로 하므로, 접수대장 행이 정렬·이동돼도
 * 항상 올바른 값을 따라갑니다. 접수대장에서 회사·제품정보를 고치면
 * 일정관리에 자동 반영됩니다.
 */
function _일정관리행추가(ss, 접수번호, 직접값) {
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 일정H = 시트헤더정의[SHEET.일정관리];
  const 대장H = 시트헤더정의[SHEET.접수대장];

  // 접수대장을 원본으로 참조할 컬럼 (일정관리 컬럼명 → 접수대장 컬럼명)
  // 대부분 이름이 같지만, 명세서/기타제출서류는 가공이 필요해 별도 처리
  const 참조맵 = {
    '기업명': '기업명',
    '담당자명': '담당자명',
    '연락처': '연락처',
    '이메일': '이메일',
    '제품명': '제품명',
    '개요': '개요',
    '제공형태': '제공형태',
    '제품분류': '제품분류',
    '인공지능적용목적': '인공지능적용목적',
    '인공지능적용범위': '인공지능적용범위',
    '명세서작성방식': '명세서작성방식',
    '보유인증': '보유인증',
    '인공지능기능수': '인공지능기능수',
  };

  const 새행번호 = 일정시트.getLastRow() + 1;
  const 대장범위 = `${SHEET.접수대장}!$A:$${columnLetter(대장H.length)}`;
  const iD접수 = 일정H.indexOf('접수번호') + 1;
  const 접수열문자 = columnLetter(iD접수);

  // 각 컬럼별 값/수식 생성
  const 행값 = 일정H.map(h => {
    // 0) 순번 = 헤더 제외한 현재 행 위치 (새행번호 - 1)
    if (h === '순번') return 새행번호 - 1;

    // 0-1) 심사링크 = 심사폼 URL + 접수번호 (CONFIG.심사폼URL 있을 때만)
    if (h === '심사링크') {
      const url = (CONFIG.심사폼URL || '').trim();
      if (!url) return '';
      const 건URL = url + (url.indexOf('?') >= 0 ? '&' : '?') + 'id=' + encodeURIComponent(접수번호);
      return `=HYPERLINK("${건URL}","▶ 심사")`;
    }

    // 1) 직접 입력값 (상태·담당심사원·신청일·심사접수일)
    if (h === '접수번호') return 접수번호;
    if (h === '상태') return '대기';
    if (Object.prototype.hasOwnProperty.call(직접값, h)) return 직접값[h];

    // 2) 마감예정일 = 신청일 + 15
    if (h === '마감예정일') {
      const iD신청 = 일정H.indexOf('신청일') + 1;
      const 신청셀 = `${columnLetter(iD신청)}${새행번호}`;
      return `=IF(${신청셀}="","",${신청셀}+15)`;
    }

    // 3) 기타제출서류여부 = 접수대장의 파일명 있으면 Y
    if (h === '기타제출서류여부') {
      return `=IF(VLOOKUP($${접수열문자}${새행번호},${대장범위},${_대장열번호('기타제출서류파일명', 대장H)},0)="","","Y")`;
    }

    // 4) 접수대장 참조 컬럼 → VLOOKUP
    if (참조맵[h]) {
      const 열번호 = _대장열번호(참조맵[h], 대장H);
      return `=IFERROR(VLOOKUP($${접수열문자}${새행번호},${대장범위},${열번호},0),"")`;
    }

    return '';
  });

  일정시트.appendRow(행값);
}

/** 접수대장 헤더에서 컬럼의 1-based 번호 반환 (VLOOKUP index용) */
function _대장열번호(컬럼명, 대장H) {
  return 대장H.indexOf(컬럼명) + 1;
}
/** 접수대장 헤더에서 컬럼의 알파벳 열문자 반환 */
function _대장열문자(컬럼명, 대장H) {
  return columnLetter(대장H.indexOf(컬럼명) + 1);
}

function 제품모델등록(접수번호, 모델목록) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getSheetByName(SHEET.제품모델);

  const D = 시트.getDataRange().getValues();
  if (D.length > 1) {
    const iNo = D[0].indexOf('접수번호');
    for (let r = D.length - 1; r >= 1; r--) {
      if (String(D[r][iNo]).trim() === String(접수번호).trim()) {
        시트.deleteRow(r + 1);
      }
    }
  }

  const 헤더행 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0]
    .map(v => String(v).trim());
  모델목록.forEach((m, i) => {
    const 값맵 = {
      '접수번호': 접수번호,
      '연번': m.연번 || (i + 1),
      '모델명': m.모델명 ?? '',
      '세부품명번호': m.세부품명번호 ?? '',
      '물품식별번호': m.물품식별번호 ?? '',
    };
    시트.appendRow(헤더행.map(h => 값맵[h] ?? ''));
  });
}

/** 접수대장 + 일정관리의 인공지능기능수 / 기능명요약 / 구현방식요약 갱신 */
function _접수대장기능수갱신(접수번호, 기능목록) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 접수대장 갱신
  const 대장시트 = ss.getSheetByName(SHEET.접수대장);
  const D = 대장시트.getDataRange().getValues();
  const H = D[0];
  const iNo = H.indexOf('접수번호');
  const i수 = H.indexOf('인공지능기능수');
  const i명 = H.indexOf('인공지능기능명(요약)');
  const i방식 = H.indexOf('구현방식(요약)');
  for (let r = 1; r < D.length; r++) {
    if (D[r][iNo] === 접수번호) {
      if (i수 >= 0) 대장시트.getRange(r + 1, i수 + 1).setValue(기능목록.length);
      if (i명 >= 0) 대장시트.getRange(r + 1, i명 + 1)
        .setValue(기능목록.map(f => f.기능명).filter(Boolean).join(' / '));
      if (i방식 >= 0) 대장시트.getRange(r + 1, i방식 + 1)
        .setValue([...new Set(기능목록.map(f => f.구현방식).filter(Boolean))].join(', '));
      break;
    }
  }

  // 일정관리의 인공지능기능수는 접수대장을 VLOOKUP 참조하므로 자동 반영됨 (별도 쓰기 불필요)
}

/**
 * 엑셀 blob → Google Sheets 변환
 * Drive 고급 서비스(v2/v3) 또는 UrlFetchApp 폴백 순으로 시도
 */
function _엑셀을Sheets로변환(blob, 제목) {
  // 1순위: Drive 고급 서비스 v2 (insert)
  try {
    const f = Drive.Files.insert(
      { title: 제목, mimeType: MimeType.GOOGLE_SHEETS },
      blob,
      { convert: true }
    );
    return f.id;
  } catch(e) {}

  // 2순위: Drive 고급 서비스 v3 (create)
  try {
    const meta = { name: 제목, mimeType: MimeType.GOOGLE_SHEETS };
    const f = Drive.Files.create(meta, blob, { convert: true });
    return f.id;
  } catch(e) {}

  // 3순위: UrlFetchApp 직접 업로드 (고급 서비스 없이도 동작)
  return _엑셀업로드_UrlFetch(blob, 제목);
}

/** Drive 고급 서비스 없이 multipart 업로드로 변환 */
function _엑셀업로드_UrlFetch(blob, 제목) {
  const token = ScriptApp.getOAuthToken();
  const meta = JSON.stringify({ name: 제목, mimeType: MimeType.GOOGLE_SHEETS });
  const boundary = '----AppsScriptBoundary';
  const body =
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    meta + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: ' + blob.getContentType() + '\r\n\r\n';

  const bodyBytes = Utilities.newBlob(body).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob('\r\n--' + boundary + '--').getBytes());

  const res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&convert=true',
    {
      method: 'POST',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'multipart/related; boundary=' + boundary,
      },
      payload: Utilities.newBlob(bodyBytes).getBytes(),
      muteHttpExceptions: true,
    }
  );

  const result = JSON.parse(res.getContentText());
  if (!result.id) throw new Error('업로드 실패: ' + res.getContentText());
  return result.id;
}

function _세로형감지(데이터) {
  if (데이터.length < 2) return false;

  // 가로형 특징: 1행(헤더)에 다수 컬럼이 채워져 있고, 2행 이후도 같은 폭으로 값이 채워짐
  // 세로형 특징: 각 행이 2열짜리 [항목명, 값] 쌍으로만 구성됨 (3번째 열부터는 항상 비어있음)
  const 열수 = 데이터[0].length;

  // 3열 이상 존재하고, 3번째 열에 값이 있는 행이 하나라도 있으면 → 가로형
  const 셋째열이상값있음 = 데이터.some(행 => {
    for (let c = 2; c < 행.length; c++) {
      if (String(행[c] ?? '').trim() !== '') return true;
    }
    return false;
  });
  if (셋째열이상값있음) return false;

  // 2열 이하인 경우: 1행이 헤더+값 한 쌍인지(가로형 1건), 항목명 나열인지(세로형) 확인
  // 세로형은 데이터 행 수가 통상 10개 이상(항목 수만큼) 되는 경우가 많고,
  // A열에 동일 키워드가 반복되지 않음(각 행이 서로 다른 항목명)
  const 키워드 = ['제품', '기업명', '서비스명', '사업자', '기능명', '연락처', '이메일'];
  const A열값들 = 데이터.map(r => String(r[0] ?? ''));
  const 항목명매치수 = A열값들.filter(v => 키워드.some(k => v.includes(k))).length;

  // 헤더 행(1행) 자체가 항목명 키워드를 포함 + 전체 행 수가 적으면(2건 미만) 세로형으로 판단
  // 반대로 1행만 키워드를 포함하고 2행부터는 실제 데이터(값)라면 가로형(헤더+데이터행)
  if (항목명매치수 >= 2 && 데이터.length <= 20) {
    // A열에 항목명이 여러 개 나열 → 세로형 (1건짜리 세로 카드)
    return true;
  }

  return false; // 기본값: 가로형 (헤더 1행 + 데이터 N행)
}

/**
 * [형태 A] 세로형 파싱
 * A열: 항목명, B열: 값  형태의 엑셀
 */
function _파싱_세로형(데이터) {
  const 맵 = {};
  데이터.forEach(행 => {
    const 키 = String(행[0]).trim();
    const 값 = String(행[1] ?? '').trim();
    if (키) 맵[키] = 값;
  });

  const V = (내부키, 기본후보들) => _맵값(맵, _별칭합치기(내부키, 기본후보들));

  return {
    // 기업 정보
    기업명:         V('기업명', ['상호(사업자명)', '기업명']),
    사업자번호:     V('사업자번호', ['법인등록번호', '사업자등록번호', '사업자번호']),
    대표자:         V('대표자', ['대표자 성명', '대표자명', '대표자']),
    소재지:         V('소재지', ['소재지']),
    담당자명:       V('담당자명', ['업무담당자', '담당자명', '담당자']),
    연락처:         V('연락처', ['담당자 전화번호', '연락처', '전화번호']),
    이메일:         V('이메일', ['담당자 이메일주소', '이메일', 'email', 'e-mail']),
    // 제품·서비스 정보
    접수번호:       V('접수번호', ['접수번호', '접수 번호', '관리번호', '관리 번호', '심사번호']),
    제품명:         V('제품명', ['제품 또는 서비스명', '제품·서비스명', '제품명', '서비스명']),
    제공형태:       V('제공형태', ['제품 또는 서비스 제공 형태', '제공 형태', '제공형태']),
    제품분류:       V('제품분류', ['인공지능 제품·서비스 분류', 'AI 제품 분류', 'AI제품분류', '제품분류']),
    개요:           V('개요', ['개요']),
    인공지능적용목적: V('인공지능적용목적', ['인공지능 적용 목적', 'AI 적용 목적', '인공지능적용목적']),
    인공지능적용범위: V('인공지능적용범위', ['인공지능 적용 범위', 'AI 적용 범위', '인공지능적용범위']),
    구조도파일명:   V('구조도파일명', ['제품 구조도', '인공지능 제품·서비스 구조도', '구조도파일명']),
    // 정부문서 제출
    명세서작성방식: V('명세서작성방식', ['명세서 작성 방식', '명세서작성방식']),
    명세서파일명:   V('명세서파일명', ['명세서 파일', '명세서파일명']),
    기타제출서류파일명: V('기타제출서류파일명', ['기타 제출서류', '기타제출서류파일명']),
    // 기존 인증
    보유인증:       V('보유인증', ['보유 인증', '보유인증']),
    인증서파일명:   V('인증서파일명', ['인증서·시험성적서', '인증서파일명']),
    // 동의서·특기사항
    열람이용동의여부: V('열람이용동의여부', ['열람·이용동의', '열람이용동의여부']),
    개인정보수집이용동의여부: V('개인정보수집이용동의여부', ['개인정보 수집·이용', '개인정보수집이용동의여부']),
    개인정보3자제공동의여부: V('개인정보3자제공동의여부', ['개인정보 제3자 제공', '개인정보3자제공동의여부']),
    특기사항:       V('특기사항', ['특기사항']),
    신청일:         V('신청일', ['신청일자', '신청일']),
    // 파서 내부 보조
    AI기능명_원본:  V('AI기능명_원본', ['AI 기능명', '기능명']),
    비고:           V('비고', ['비고', '기타']),
    원본맵: 맵,
  };
}

/**
 * [형태 B] 가로형 파싱 (1행 헤더 + 다수 데이터행)
 */
function _파싱_가로형(데이터) {
  if (데이터.length < 2) return [];
  const 헤더 = 데이터[0].map(v => String(v).trim());
  const 결과 = [];

  for (let i = 1; i < 데이터.length; i++) {
    const 행 = 데이터[i];
    if (행.every(v => !v)) continue; // 빈 행 스킵

    const 맵 = {};
    헤더.forEach((h, idx) => { 맵[h] = String(행[idx] ?? '').trim(); });
    결과.push(_파싱_세로형(Object.entries(맵).map(([k, v]) => [k, v])));
  }
  return 결과;
}

function _맵값(맵, 후보키들) {
  // 1차: 모든 후보에 대해 완전 일치 우선
  for (const 키 of 후보키들) {
    if (맵[키] !== undefined && 맵[키] !== '') return 맵[키];
  }
  // 2차: 완전 일치가 없을 때만 부분 일치 폴백
  for (const 키 of 후보키들) {
    const 일치 = Object.keys(맵).find(k => k.includes(키));
    if (일치 && 맵[일치] !== '') return 맵[일치];
  }
  return '';
}


// ─────────────────────────────────────────────
// 3. Sheets 등록 — 접수대장 + 일정관리 기록
// ─────────────────────────────────────────────
function _Sheets에등록(건, 파일명) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 접수번호: 엑셀에 있으면 그대로 키로 사용. 없으면 자체 생성(폴백)
  const 엑셀접수번호 = String(건.접수번호 || '').trim();
  const 접수번호 = 엑셀접수번호 || _접수번호생성(ss);

  const 오늘 = new Date();
  const 마감일 = new Date(오늘);
  마감일.setDate(마감일.getDate() + CONFIG.기본심사기간);

  const 대장시트 = ss.getSheetByName(SHEET.접수대장);
  const D = 대장시트.getDataRange().getValues();
  const H = D[0];
  const iNo = H.indexOf('접수번호');

  // 기존 행 찾기 (같은 접수번호면 갱신)
  let 기존행번호 = -1;
  for (let r = 1; r < D.length; r++) {
    if (String(D[r][iNo]).trim() === 접수번호) { 기존행번호 = r + 1; break; }
  }
  const 신규 = (기존행번호 === -1);

  // 담당자: 신규면 새로 배분, 갱신이면 기존 담당자 유지
  let 담당자, 착수일, 상태;
  if (신규) {
    담당자 = _담당자배분(ss);
    착수일 = Utilities.formatDate(오늘, 'Asia/Seoul', 'yyyy-MM-dd');
    상태 = '접수';
  } else {
    const i담당 = H.indexOf('담당심사원');
    const i착수 = H.indexOf('심사착수일');
    const i상태 = H.indexOf('상태');
    담당자 = D[기존행번호 - 1][i담당] || _담당자배분(ss);
    착수일 = D[기존행번호 - 1][i착수] || Utilities.formatDate(오늘, 'Asia/Seoul', 'yyyy-MM-dd');
    상태 = D[기존행번호 - 1][i상태] || '접수';  // 기존 상태 보존
  }

  // ── 헤더 이름 기반 쓰기 ──
  // 컬럼 순서가 바뀌거나 시트에 컬럼이 추가/삽입돼도 값이 밀리지 않습니다.
  // 값맵에 없는 컬럼은 갱신 시 기존 값을 보존합니다.
  const 값맵 = {
    // TTA 관리
    '접수번호': 접수번호,
    '심사접수일': 신규 ? Utilities.formatDate(오늘, 'Asia/Seoul', 'yyyy-MM-dd')
                   : (H.indexOf('심사접수일') >= 0 ? D[기존행번호 - 1][H.indexOf('심사접수일')]
                      : Utilities.formatDate(오늘, 'Asia/Seoul', 'yyyy-MM-dd')),
    '상태': 상태,
    // 기업 정보
    '기업명': 건.기업명, '사업자번호': 건.사업자번호, '대표자': 건.대표자, '소재지': 건.소재지,
    '담당자명': 건.담당자명, '연락처': 건.연락처, '이메일': 건.이메일,
    // 제품·서비스 정보
    '제품명': 건.제품명, '제공형태': 건.제공형태, '제품분류': 건.제품분류,
    '개요': 건.개요,
    '인공지능적용목적': 건.인공지능적용목적,
    '인공지능적용범위': 건.인공지능적용범위,
    '구조도파일명': 건.구조도파일명,
    // 정부문서 제출
    '명세서작성방식': 건.명세서작성방식,
    '명세서파일명': 건.명세서파일명,
    '기타제출서류파일명': 건.기타제출서류파일명,
    // 기존 인증
    '보유인증': 건.보유인증, '인증서파일명': 건.인증서파일명,
    // 동의서·특기사항
    '열람이용동의여부': 건.열람이용동의여부,
    '개인정보수집이용동의여부': 건.개인정보수집이용동의여부,
    '개인정보3자제공동의여부': 건.개인정보3자제공동의여부,
    '특기사항': 건.특기사항, '신청일': 건.신청일,
    // TTA 심사 관리 (기능수·구현방식요약은 기능탭 등록 후 별도 갱신)
    '인공지능기능명(요약)': 건.AI기능명_원본,
    '담당심사원': 담당자, '심사착수일': 착수일,
    '심사마감일': Utilities.formatDate(마감일, 'Asia/Seoul', 'yyyy-MM-dd'),
    '비고': 건.비고,
  };

  const 행값 = H.map((h, i) => {
    if (Object.prototype.hasOwnProperty.call(값맵, h)) {
      return 값맵[h] ?? '';
    }
    // 값맵에 정의되지 않은 컬럼: 신규는 빈칸, 갱신은 기존 값 보존
    return 신규 ? '' : D[기존행번호 - 1][i];
  });

  if (신규) {
    대장시트.appendRow(행값);
  } else {
    대장시트.getRange(기존행번호, 1, 1, 행값.length).setValues([행값]);
  }

  // ── 일정관리: 신규만 추가 ──
  // 하이브리드 싱크:
  //   · 신청정보(회사·연락처·제품 등) = 접수대장 VLOOKUP 수식 참조 (접수대장이 원본)
  //   · 상태·담당심사원 = 일정관리에서 직접 편집 (일정관리가 원본)
  //   · 마감예정일 = 신청일 + 15 수식
  if (신규) {
    _일정관리행추가(ss, 접수번호, {
      신청일: 건.신청일 || '',
      심사접수일: Utilities.formatDate(오늘, 'Asia/Seoul', 'yyyy-MM-dd'),
      담당심사원: 담당자,   // 최초 배분값 (이후 일정관리에서 직접 수정 가능)
    });
  }

  // ── 로그 ──
  ss.getSheetByName(SHEET.로그).appendRow([
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    파일명, 접수번호,
    신규 ? '신규등록' : '갱신',
    `${건.기업명} / ${건.제품명} ${신규 ? '등록' : '갱신'} → 담당: ${담당자}`,
  ]);

  // ── 알림: 신규일 때만 ──
  if (신규) {
    if (CONFIG.알림이메일) _이메일발송(CONFIG.알림이메일, 접수번호, 건, 담당자, 마감일);
    _챗알림발송(접수번호, 건, 담당자, 마감일);
  }

  Logger.log(`${신규 ? '등록' : '갱신'}: ${접수번호} / ${건.제품명}`);
  return 접수번호;
}

function _접수번호생성(ss) {
  const 시트 = ss.getSheetByName(SHEET.접수대장);
  const 마지막행 = 시트.getLastRow();
  const 연도 = new Date().getFullYear();
  const 순번 = String(마지막행).padStart(4, '0'); // 헤더행 포함 → 실제 건수
  return `AI-${연도}-${순번}`;
}

function _담당자배분(ss) {
  const 시트 = ss.getSheetByName(SHEET.접수대장);
  const 마지막행 = 시트.getLastRow() - 1; // 헤더 제외
  const idx = 마지막행 % CONFIG.담당자목록.length;
  return CONFIG.담당자목록[idx];
}

// ─────────────────────────────────────────────
// Google Chat 알림
// ─────────────────────────────────────────────

/**
 * 신규 접수 시 Google Chat으로 알림 발송
 *  - 담당자 개인 Webhook이 설정되어 있으면 → DM 발송
 *  - 없으면 → 공통 채널에 발송
 *  - 둘 다 없으면 → 조용히 스킵
 */
function _챗알림발송(접수번호, 건, 담당자, 마감일) {
  const 마감문자열 = Utilities.formatDate(마감일, 'Asia/Seoul', 'yyyy-MM-dd');
  const 시트URL = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  // Card v2 형식 메시지 (Google Chat 기본 카드)
  const payload = {
    cardsV2: [{
      cardId: 접수번호,
      card: {
        header: {
          title: '🔔 AI 기술심사 신규 접수',
          subtitle: `${접수번호}  ·  담당: ${담당자}`,
          imageUrl: 'https://fonts.gstatic.com/s/i/googlematerialicons/assignment/v11/black-48dp/1x/gm_assignment_black_48dp.png',
          imageType: 'CIRCLE',
        },
        sections: [
          {
            header: '접수 정보',
            widgets: [
              { decoratedText: { topLabel: '제품명',   text: `<b>${건.제품명}</b> (${건.제품버전 || '-'})` } },
              { decoratedText: { topLabel: '기업명',   text: 건.기업명 } },
              { decoratedText: { topLabel: '분류',     text: 건.제품분류 || '-' } },
              { decoratedText: { topLabel: '기능 수',  text: `${건.AI기능명_원본 ? 건.AI기능명_원본.split('/').length : '-'}개` } },
            ],
          },
          {
            header: '심사 일정',
            widgets: [
              { decoratedText: { topLabel: '담당심사원', text: `<b>${담당자}</b>` } },
              { decoratedText: {
                topLabel: '심사 마감일',
                text: `<b>${마감문자열}</b>`,
                startIcon: { knownIcon: 'CLOCK' },
              }},
            ],
          },
          {
            widgets: [{
              buttonList: { buttons: [{
                text: '접수대장 바로가기',
                onClick: { openLink: { url: 시트URL } },
              }]},
            }],
          },
        ],
      },
    }],
  };

  // 발송 대상 결정
  const 개인Webhook = CONFIG.챗_담당자Webhook?.[담당자];
  const 대상URL = 개인Webhook || CONFIG.챗_공통Webhook;

  if (!대상URL) return; // Webhook 미설정 시 스킵

  _챗POST(대상URL, payload);

  // 공통 채널이 따로 있고 개인 DM도 보냈다면 → 공통 채널에도 요약 발송
  if (개인Webhook && CONFIG.챗_공통Webhook && 개인Webhook !== CONFIG.챗_공통Webhook) {
    const 요약payload = {
      text: `✅ *${접수번호}* 접수 완료 — ${건.기업명} / ${건.제품명}\n담당: *${담당자}* · 마감: ${마감문자열}`,
    };
    _챗POST(CONFIG.챗_공통Webhook, 요약payload);
  }
}

/**
 * 상태 변경 시 담당자에게 알림
 * 사용 예: 접수대장에서 상태 셀을 바꾼 뒤 이 함수 호출
 */
function 상태변경알림() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getActiveSheet();
  if (시트.getName() !== SHEET.접수대장) return;

  const 행 = 시트.getActiveRange().getRow();
  if (행 <= 1) return;

  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0];
  const 데이터 = 시트.getRange(행, 1, 1, 시트.getLastColumn()).getValues()[0];
  const 건 = {};
  헤더.forEach((h, i) => { 건[h] = 데이터[i]; });

  const 상태이모지 = {
    '접수': '📥', '심사중': '🔍', '보완요청': '🔄', '완료': '✅', '반려': '❌',
  };
  const 이모지 = 상태이모지[건['상태']] || '📋';

  const payload = {
    text: `${이모지} *[${건['접수번호']}] 상태 변경 → ${건['상태']}*\n${건['제품명']} (${건['기업명']})\n담당: ${건['담당심사원']}`,
  };

  const 대상URL = CONFIG.챗_담당자Webhook?.[건['담당심사원']] || CONFIG.챗_공통Webhook;
  if (대상URL) _챗POST(대상URL, payload);
}

/** Webhook POST 공통 함수 */
function _챗POST(url, payload) {
  try {
    UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('Chat 알림 실패: ' + e.message);
  }
}

/** Webhook URL 설정 안내 및 테스트 발송 */
function 챗알림테스트() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Google Chat Webhook 테스트',
    'Webhook URL을 입력하세요.\n\n' +
    '[URL 발급 방법]\n' +
    '1. Google Chat 스페이스 열기\n' +
    '2. 스페이스 이름 클릭 > 앱 및 통합\n' +
    '3. Webhook > Webhook 추가 > URL 복사\n\n' +
    '(CONFIG에 저장된 URL을 테스트하려면 빈 칸으로 OK)',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const url = res.getResponseText().trim() || CONFIG.챗_공통Webhook;
  if (!url) { ui.alert('Webhook URL이 없습니다. CONFIG.챗_공통Webhook을 먼저 설정하세요.'); return; }

  _챗POST(url, {
    text: '✅ *AI 기술심사 관리 시스템* 연결 테스트\n알림이 정상적으로 도착했습니다! 🎉',
  });
  ui.alert('테스트 메시지를 발송했습니다. Google Chat을 확인하세요.');
}

function _이메일발송(수신자, 접수번호, 건, 담당자, 마감일) {
  const 제목 = `[AI기술심사] 신규 접수 ${접수번호} — ${건.제품명}`;
  const 본문 = `
새로운 인공지능 제품 기술심사 건이 접수되었습니다.

■ 접수번호: ${접수번호}
■ 제품명: ${건.제품명} (${건.제품버전})
■ 기업명: ${건.기업명}
■ 담당심사원: ${담당자}
■ 심사마감일: ${Utilities.formatDate(마감일, 'Asia/Seoul', 'yyyy-MM-dd')}

Google Sheets에서 상세 내용을 확인하세요.
  `.trim();
  MailApp.sendEmail(수신자, 제목, 본문);
}


// ─────────────────────────────────────────────
// 4. AI기능상세 등록 (별도 실행 또는 파싱 시 자동 호출)
// ─────────────────────────────────────────────
function AI기능상세등록(접수번호, 기능목록) {
  /**
   * 기능목록 예시:
   * [
   *   { 기능번호:1, 기능명:'흉부 영상 이상 탐지', 역할:'...', 구현방식:'A', ... },
   *   ...
   * ]
   */
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getSheetByName(SHEET.AI기능상세);

  // 같은 접수번호의 기존 기능 행 삭제 (재파싱 시 중복 방지)
  const D = 시트.getDataRange().getValues();
  if (D.length > 1) {
    const iNo = D[0].indexOf('접수번호');
    for (let r = D.length - 1; r >= 1; r--) {
      if (String(D[r][iNo]).trim() === String(접수번호).trim()) {
        시트.deleteRow(r + 1);
      }
    }
  }

  // 헤더 이름 기반 쓰기 — 시트 컬럼 순서와 무관하게 올바른 칸에 기록
  const 헤더행 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0]
    .map(v => String(v).trim());
  기능목록.forEach(f => {
    const 값맵 = {
      '접수번호': 접수번호,
      '기능번호': f.기능번호 ?? '',
      '기능명': f.기능명 ?? '',
      '인공지능역할': f.인공지능역할 ?? '',
      '입력': f.입력 ?? '',
      '출력': f.출력 ?? '',
      '설명서참조위치': f.설명서참조위치 ?? '',
      '구현방식': f.구현방식 ?? '',
      '연산자원': f.연산자원 ?? '',
      '실행환경': f.실행환경 ?? '',
      '학습데이터사양': f.학습데이터사양 ?? '',
      '재현환경라이선싱': f.재현환경라이선싱 ?? '',
      'BaseModel': f.BaseModel ?? '',
      '튜닝방법': f.튜닝방법 ?? '',
      '외부API모델': f.외부API모델 ?? '',
      '타겟HW_OS': f.타겟HW_OS ?? '',
      '배포포맷정밀도': f.배포포맷정밀도 ?? '',
      '혼합구성설명': f.혼합구성설명 ?? '',
      '입력데이터설명': f.입력데이터설명 ?? '',
      '출력데이터설명': f.출력데이터설명 ?? '',
      '기타참고자료파일명': f.기타참고자료파일명 ?? '',
      '데이터흐름도파일명': f.데이터흐름도파일명 ?? '',
    };
    시트.appendRow(헤더행.map(h => 값맵[h] ?? ''));
  });
}


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
function 보고서생성() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getActiveSheet();

  if (시트.getName() !== SHEET.접수대장) {
    SpreadsheetApp.getUi().alert('접수대장 시트에서 실행해주세요.');
    return;
  }

  const 행번호 = 시트.getActiveRange().getRow();
  if (행번호 <= 1) {
    SpreadsheetApp.getUi().alert('데이터 행을 선택해주세요.');
    return;
  }

  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0];
  const 데이터 = 시트.getRange(행번호, 1, 1, 시트.getLastColumn()).getValues()[0];
  const 건 = {};
  헤더.forEach((h, i) => { 건[h] = 데이터[i]; });

  const 접수번호 = 건['접수번호'];
  const 기능목록 = _AI기능상세조회(ss, 접수번호);

  const doc = _Docs보고서생성(건, 기능목록);
  SpreadsheetApp.getUi().alert(
    `보고서가 생성되었습니다!\n\n${doc.getUrl()}`
  );
}

function _AI기능상세조회(ss, 접수번호) {
  const 시트 = ss.getSheetByName(SHEET.AI기능상세);
  const 모든행 = 시트.getDataRange().getValues();
  const 헤더 = 모든행[0];
  return 모든행.slice(1)
    .filter(행 => 행[0] === 접수번호)
    .map(행 => {
      const obj = {};
      헤더.forEach((h, i) => { obj[h] = 행[i]; });
      return obj;
    });
}

function _Docs보고서생성(건, 기능목록) {
  const 제목 = `[기술심사보고서] ${건['접수번호']} — ${건['제품명']}`;
  const doc = DocumentApp.create(제목);
  const body = doc.getBody();

  // 제목
  body.appendParagraph(제목)
    .setHeading(DocumentApp.ParagraphHeading.TITLE);

  body.appendParagraph(
    `작성일: ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 MM월 dd일')}` +
    `   담당심사원: ${건['담당심사원']}`
  ).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);

  body.appendHorizontalRule();

  // 1. 제품 기본 정보
  _섹션제목(body, '1. 제품 기본 정보');
  const 기본표 = body.appendTable([
    ['항목', '내용'],
    ['제품명', 건['제품명']],
    ['제품버전', 건['제품버전']],
    ['제공형태', 건['제공형태']],
    ['제품분류', 건['제품분류']],
    ['세부품명번호', 건['세부품명번호']],
    ['물품식별번호', 건['물품식별번호']],
  ]);
  _표헤더스타일(기본표);

  // 2. 신청기업 정보
  _섹션제목(body, '2. 신청기업 정보');
  const 기업표 = body.appendTable([
    ['항목', '내용'],
    ['기업명', 건['기업명']],
    ['사업자번호', 건['사업자번호']],
    ['대표자', 건['대표자']],
    ['담당자', `${건['담당자명']}  ${건['연락처']}  ${건['이메일']}`],
    ['소재지', 건['소재지']],
  ]);
  _표헤더스타일(기업표);

  // 3. 인공지능 기능 목록
  _섹션제목(body, '3. 인공지능 기능 목록');
  if (기능목록.length > 0) {
    const 기능표헤더 = ['기능번호', '기능명', '인공지능역할(요약)', '구현방식'];
    const 기능표데이터 = 기능목록.map(f => [
      String(f['기능번호'] ?? ''),
      f['기능명'] ?? '',
      String(f['인공지능역할'] ?? '').slice(0, 60) + (String(f['인공지능역할'] ?? '').length > 60 ? '…' : ''),
      f['구현방식'] ?? '',
    ]);
    const 기능표 = body.appendTable([기능표헤더, ...기능표데이터]);
    _표헤더스타일(기능표);
  } else {
    body.appendParagraph('(인공지능기능상세 시트에 데이터 없음 — 별도 입력 필요)');
  }

  // 4. 기능별 인공지능 구현 상세
  _섹션제목(body, '4. 기능별 인공지능 구현 상세');
  기능목록.forEach(f => {
    body.appendParagraph(`4.${f['기능번호']}  ${f['기능명']}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);

    const 상세표 = body.appendTable([
      ['항목', '내용'],
      ['구현방식', _구현방식레이블(f['구현방식'])],
      ['연산자원', f['연산자원'] ?? ''],
      ['실행환경', f['실행환경'] ?? ''],
      ['입력데이터설명', f['입력데이터설명'] ?? ''],
      ['출력데이터설명', f['출력데이터설명'] ?? ''],
      ['Base Model', f['BaseModel'] ?? '-'],
      ['튜닝방법', f['튜닝방법'] ?? '-'],
      ['외부API', f['외부API모델'] ?? '-'],
    ]);
    _표헤더스타일(상세표);
    body.appendParagraph('');
  });

  // 5. 심사 의견 (빈칸)
  _섹션제목(body, '5. 심사 의견');
  ['적합성 검토', '보완 필요 사항', '종합 의견'].forEach(항목 => {
    body.appendParagraph(항목 + ':')
      .setBold(true);
    body.appendParagraph('\n\n');
  });

  // 6. 심사 결과
  _섹션제목(body, '6. 심사 결과');
  body.appendTable([
    ['심사 결과', '□ 적합   □ 조건부 적합   □ 부적합'],
    ['심사 완료일', ''],
    ['심사원 서명', 건['담당심사원']],
  ]);

  doc.saveAndClose();
  return doc;
}

function _섹션제목(body, 텍스트) {
  body.appendParagraph('');
  body.appendParagraph(텍스트)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
}

function _표헤더스타일(표) {
  try {
    const 헤더행 = 표.getRow(0);
    for (let i = 0; i < 헤더행.getNumCells(); i++) {
      헤더행.getCell(i).setBackgroundColor('#1a73e8');
      헤더행.getCell(i).getText(); // 접근 확인
    }
  } catch(e) { /* 스타일 실패는 무시 */ }
}

function _구현방식레이블(코드) {
  const 맵 = {
    A: 'A. 자체 학습 모델',
    B: 'B. 사전학습+파인튜닝',
    C: 'C. 외부 API/모델 활용',
    D: 'D. 엣지 디바이스 추론',
    E: 'E. 혼합형',
  };
  return 맵[코드] ?? (코드 || '미입력');
}


// ─────────────────────────────────────────────
// 6. 사용자 메뉴 등록
// ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 인공지능심사관리')
    .addItem('📝 심사 폼 열기', '심사폼열기')
    .addItem('🔗 접수대장에 심사 링크 채우기', '심사링크채우기')
    .addSeparator()
    .addItem('📁 엑셀 파일 등록 (파싱)', '엑셀파싱등록')
    .addSeparator()
    .addItem('📑 기술심사보고서 생성 (선택 행)', '증적명세서생성')
    .addItem('📦 보고서 3종 산출 (PDF·DOCX·JSON)', '증적3종산출')
    .addSeparator()
    .addItem('⚙️ 초기 설정 (최초 1회)', '초기설정실행')
    .addItem('🔧 헤더만 재설정 (데이터 유지)', '헤더만재설정')
    .addItem('⚠️ 시트 재생성 (데이터 초기화)', '시트재생성_데이터초기화')
    .addItem('💬 Chat 알림 테스트', '챗알림테스트')
    .addToUi();
}

// 심사 폼 웹앱을 새 탭에서 열기 (CONFIG.심사폼URL 사용)
function 심사폼열기() {
  const url = (CONFIG.심사폼URL || '').trim();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      '심사 폼 URL이 설정되지 않았습니다.\n\n' +
      'Apps Script 편집기에서 CONFIG.심사폼URL 칸에 ' +
      '배포된 웹앱 URL을 붙여넣어 주세요.'
    );
    return;
  }
  // 선택된 행이 있으면 그 접수번호를 URL에 붙여 바로 해당 건 열기
  let 최종URL = url;
  try {
    const 시트 = SpreadsheetApp.getActiveSheet();
    const 시트명 = 시트.getName();
    // 메인 시트(일정관리) 또는 접수대장에서 활성 행의 접수번호를 읽음
    if (시트명 === SHEET.일정관리 || 시트명 === SHEET.접수대장) {
      const 행 = 시트.getActiveRange().getRow();
      if (행 > 1) {
        // 두 시트 모두 접수번호가 있는 열을 헤더에서 찾아서 사용
        const H = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0];
        const iNo = H.indexOf('접수번호');
        if (iNo >= 0) {
          const 접수번호 = 시트.getRange(행, iNo + 1).getValue();
          if (접수번호) 최종URL = url + (url.indexOf('?') >= 0 ? '&' : '?') + 'id=' + encodeURIComponent(접수번호);
        }
      }
    }
  } catch (e) {}

  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${최종URL}', '_blank'); google.script.host.close();</script>` +
    `<p style="font-family:sans-serif;font-size:13px">심사 폼을 새 탭에서 엽니다…<br>` +
    `열리지 않으면 <a href="${최종URL}" target="_blank">여기를 클릭</a>하세요.</p>`
  ).setWidth(360).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, '심사 폼 열기');
}

// 접수대장 각 행에 "심사" 하이퍼링크를 채워넣기 (메뉴에서 한 번 실행)
// 마지막 열 다음에 '심사링크' 열을 만들고 건별 링크를 HYPERLINK 수식으로 넣습니다.
function 심사링크채우기() {
  const url = (CONFIG.심사폼URL || '').trim();
  if (!url) { SpreadsheetApp.getUi().alert('CONFIG.심사폼URL을 먼저 설정하세요.'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getSheetByName(SHEET.일정관리);  // 메인 시트
  const D = 시트.getDataRange().getValues();
  const H = D[0];
  let i링크 = H.indexOf('심사링크');
  if (i링크 < 0) {
    i링크 = H.length;
    시트.getRange(1, i링크 + 1).setValue('심사링크')
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  }
  const iNo = H.indexOf('접수번호');

  for (let r = 1; r < D.length; r++) {
    const 접수번호 = D[r][iNo];
    if (!접수번호) continue;
    const 건URL = url + (url.indexOf('?') >= 0 ? '&' : '?') + 'id=' + encodeURIComponent(접수번호);
    시트.getRange(r + 1, i링크 + 1)
      .setFormula(`=HYPERLINK("${건URL}","▶ 심사")`);
  }
  SpreadsheetApp.getUi().alert('일정관리 시트 각 행에 "심사" 링크를 채웠습니다. 클릭하면 해당 건 심사 폼이 열립니다.');
}


// ─────────────────────────────────────────────
// 7. 수동 직접 입력 (엑셀 없이 테스트/긴급 등록)
// ─────────────────────────────────────────────
function 수동직접등록_예시() {
  /**
   * 엑셀 파일 없이 직접 데이터를 입력해 테스트할 수 있습니다.
   * 이 함수를 복사해서 실제 데이터로 수정 후 실행하세요.
   */
  const 건 = {
    제품명: 'Mediview 인공지능 판독 보조 서비스',
    제품버전: 'v1.8.2-pilot',
    제공형태: 'SaaS(클라우드 서비스형)',
    제품분류: '서비스',
    세부품명번호: '00000000',
    물품식별번호: '00000000',
    기업명: '주식회사 메디테크랩',
    사업자번호: '123-45-67890',
    대표자: '김대표',
    담당자명: '이담당',
    연락처: '02-1234-5678',
    이메일: 'pilot@meditechlab.example',
    소재지: '서울특별시 강남구 테헤란로 152',
    개요: '흉부 X-ray 영상에서 이상 소견을 탐지하고 판독 초안을 생성하는 영상판독 보조 서비스입니다.',
    인공지능적용목적: '판독 소요 시간을 단축하고 초기 스크리닝 정확도를 높이는 것을 목적으로 합니다.',
    AI기능명_원본: '흉부 영상 이상 탐지 / 폐렴 위험도 분류 / 판독 소견 요약 생성',
    비고: '',
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 접수번호 = _Sheets에등록(건, '수동입력');

  // 인공지능 기능 상세도 함께 등록
  AI기능상세등록(접수번호, [
    {
      기능번호: 1, 기능명: '흉부 영상 이상 탐지',
      인공지능역할: '흉부 X-ray 이미지에서 이상 소견 후보 탐지',
      입력: 'DICOM 영상, 검사 ID, 메타데이터',
      출력: 'Bounding Box, 위험도 점수, 소견 라벨',
      구현방식: 'A', 연산자원: 'NVIDIA A100 40GB',
      실행환경: 'Ubuntu 22.04, CUDA 12.1, PyTorch 2.6.0',
      입력데이터설명: 'DICOM 영상, 촬영 메타데이터',
      출력데이터설명: 'Bounding Box 및 위험도 점수',
    },
    {
      기능번호: 2, 기능명: '폐렴 위험도 분류',
      인공지능역할: '폐렴 의심 정도와 중증도 등급 분류',
      입력: '흉부 X-ray, 환자 나이대, 촬영 조건',
      출력: '폐렴 의심 등급, 중증도, 확신도',
      구현방식: 'B', BaseModel: 'ResNet-50 v2', 튜닝방법: '전이학습 + LoRA',
      입력데이터설명: '흉부 X-ray, 환자 나이대, 촬영 조건',
      출력데이터설명: '폐렴 의심 등급, 중증도, 확신도',
    },
    {
      기능번호: 3, 기능명: '판독 소견 요약 생성',
      인공지능역할: '탐지 결과와 임상 메모 기반 판독 초안 생성',
      입력: '탐지 결과 JSON, 임상 메모 텍스트',
      출력: '판독 초안, 주요 근거 요약',
      구현방식: 'C', 외부API모델: 'OpenAI gpt-4o-mini',
      입력데이터설명: '탐지 결과 JSON, 임상 메모 텍스트',
      출력데이터설명: '판독 초안 텍스트',
    },
  ]);

  SpreadsheetApp.getUi().alert(`수동 등록 완료!\n접수번호: ${접수번호}`);
}

// ═════════════════════════════════════════════════════════
// 8. 체크 전용 심사 폼 — 웹앱 (우선순위 1)
// ═════════════════════════════════════════════════════════
//
// [배포 방법]
// 1. Apps Script 편집기 > 배포 > 새 배포
// 2. 유형: 웹앱
// 3. 실행 계정: 나 / 액세스: 도메인 내 모든 사용자(또는 본인만)
// 4. 배포 → 생성된 웹앱 URL을 심사원에게 공유
//
// 심사원은 URL 끝에 ?id=접수번호 를 붙이거나, 목록에서 선택해 들어갑니다.
// 원본 데이터(접수대장/AI기능상세)는 읽기만 하고, 체크 결과는
// '심사체크결과' 시트에만 기록됩니다. 원본은 절대 수정되지 않습니다.

function doGet(e) {
  const 접수번호 = e && e.parameter && e.parameter.id ? e.parameter.id : '';
  const t = HtmlService.createTemplateFromFile('심사폼');
  t.초기접수번호 = 접수번호;
  return t.evaluate()
    .setTitle('AI 기술심사 체크리스트')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 웹앱: 심사 대상 목록 조회 (드롭다운용) */
function 웹_접수목록조회() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getSheetByName(SHEET.접수대장);
  if (!시트 || 시트.getLastRow() < 2) return [];

  const 데이터 = 시트.getDataRange().getValues();
  const 헤더 = 데이터[0];
  const iNo = 헤더.indexOf('접수번호');
  const iProd = 헤더.indexOf('제품명');
  const iComp = 헤더.indexOf('기업명');
  const iOwner = 헤더.indexOf('담당심사원');
  const iStat = 헤더.indexOf('상태');

  return 데이터.slice(1)
    .filter(r => r[iNo])
    .map(r => ({
      접수번호: r[iNo],
      제품명: r[iProd],
      기업명: r[iComp],
      담당심사원: r[iOwner],
      상태: r[iStat],
    }));
}

/** 웹앱: 특정 건의 체크 폼 데이터 조회 (원본은 읽기 전용) */
function 웹_체크폼데이터(접수번호) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) 접수대장에서 기본 정보 읽기
  const 대장 = ss.getSheetByName(SHEET.접수대장);
  const D = 대장.getDataRange().getValues();
  const H = D[0];
  const 행 = D.slice(1).find(r => r[H.indexOf('접수번호')] === 접수번호);
  if (!행) return { error: '접수번호를 찾을 수 없습니다: ' + 접수번호 };

  const 건 = {};
  H.forEach((h, i) => { 건[h] = 행[i]; });

  // 2) AI기능상세에서 해당 건 기능 목록
  const 상세시트 = ss.getSheetByName(SHEET.AI기능상세);
  let 기능목록 = [];
  if (상세시트 && 상세시트.getLastRow() >= 2) {
    const SD = 상세시트.getDataRange().getValues();
    const SH = SD[0];
    기능목록 = SD.slice(1)
      .filter(r => r[SH.indexOf('접수번호')] === 접수번호)
      .map(r => {
        const o = {};
        SH.forEach((h, i) => { o[h] = r[i]; });
        return o;
      });
  }

  // 3) 기존 체크 결과 불러오기 (이어서 검토 가능)
  const 기존체크 = _체크결과조회(ss, 접수번호);

  // 4) 항목별로 "엑셀에서 채워진 값" 매핑
  const 항목들 = 체크항목정의.map(정의 => ({
    id: 정의.id,
    no: 정의.no,
    항목: 정의.항목,
    채워진값: _참조값생성(정의.참조, 건, 기능목록),
    기존판정: 기존체크[정의.id] ? 기존체크[정의.id].판정 : '',
    기존사유: 기존체크[정의.id] ? 기존체크[정의.id].사유 : '',
  }));

  // 5) 기능명세 표용 데이터 (ID = 접수번호-기능번호, 증빙 캡처 파일명으로 사용)
  const 기존기능확인 = 웹_기능확인조회(접수번호);
  const 기능표 = 기능목록.map((f, idx) => {
    const ID = `${접수번호}-F${String(f['기능번호'] || (idx + 1)).padStart(2, '0')}`;
    return {
      ID: ID,
      기능번호: f['기능번호'] || (idx + 1),
      기능명: f['기능명'] || '',
      인공지능역할: f['인공지능역할'] || '',
      입력: f['입력'] || '',
      출력: f['출력'] || '',
      구현방식: f['구현방식'] || '',
      설명서참조위치: f['설명서참조위치'] || '',
      확인: 기존기능확인[ID] ? (기존기능확인[ID].확인 === 'Y') : false,
      비고: 기존기능확인[ID] ? 기존기능확인[ID].비고 : '',
    };
  });
  // 기타 행 (ID = 접수번호-ETC)
  const 기타ID = `${접수번호}-ETC`;
  const 기타행 = {
    ID: 기타ID,
    기능번호: '',
    기능명: '기타',
    인공지능역할: '', 입력: '', 출력: '', 구현방식: '', 설명서참조위치: '',
    확인: 기존기능확인[기타ID] ? (기존기능확인[기타ID].확인 === 'Y') : false,
    비고: 기존기능확인[기타ID] ? 기존기능확인[기타ID].비고 : '',
    기타: true,
  };
  기능표.push(기타행);

  return {
    접수번호: 접수번호,
    제품명: 건['제품명'],
    기업명: 건['기업명'],
    담당심사원: 건['담당심사원'],
    상태: 건['상태'],
    항목들: 항목들,
    기능표: 기능표,
  };
}

/** 참조 출처에 따라 화면에 보여줄 "제출값" 문자열 생성 */
function _참조값생성(참조, 건, 기능목록) {
  switch (참조) {
    case '제공형태+제품분류':
      return `제공형태: ${건['제공형태'] || '-'} · 분류: ${건['제품분류'] || '-'}`;
    case '제품명':
      return `제품: ${건['제품명'] || '-'} (${건['제품버전'] || '-'})`;
    case '인공지능기능수': {
      const 이름들 = 기능목록.map(f => f['기능명']).filter(Boolean).join(', ');
      return `기능 ${기능목록.length}개${이름들 ? ' — ' + 이름들 : ''}`;
    }
    case '외부API모델': {
      const api = 기능목록.map(f => f['외부API모델']).filter(Boolean);
      return api.length ? `외부 API: ${api.join(', ')}` : '외부 API 사용 기능 없음';
    }
    case '기업명':
      return `${건['기업명'] || '-'} · 사업자 ${건['사업자번호'] || '-'} · 대표 ${건['대표자'] || '-'} · ${건['연락처'] || '-'}`;
    case '세부품명번호+물품식별번호':
      return `세부품명 ${건['세부품명번호'] || '-'} · 물품식별 ${건['물품식별번호'] || '-'}`;
    case '비고':
      return 건['비고'] || '(비고 없음)';
    case '종합':
      return '전체 항목 검토 후 종합 판정';
    default:
      return 건[참조] || '-';
  }
}

// 체크결과 시트가 없으면 만들어서 반환 (없을 때 에러 방지)
function _체크결과시트확보(ss) {
  let 시트 = ss.getSheetByName(SHEET.체크결과);
  if (!시트) {
    시트 = ss.insertSheet(SHEET.체크결과);
    const 헤더 = ['접수번호', '항목ID', '항목번호', '체크항목', '판정', '사유', '심사원', '검토일시'];
    시트.getRange(1, 1, 1, 헤더.length).setValues([헤더])
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    시트.setFrozenRows(1);
  }
  return 시트;
}

function _체크결과조회(ss, 접수번호) {
  const 시트 = _체크결과시트확보(ss);
  const out = {};
  if (시트.getLastRow() < 2) return out;
  const D = 시트.getDataRange().getValues();
  const H = D[0];
  D.slice(1).forEach(r => {
    if (r[H.indexOf('접수번호')] === 접수번호) {
      out[r[H.indexOf('항목ID')]] = {
        판정: r[H.indexOf('판정')],
        사유: r[H.indexOf('사유')],
      };
    }
  });
  return out;
}

/**
 * 웹앱: 체크 결과 저장 (원본 미수정, 체크결과 시트에만 upsert)
 * results = [{ id, no, 항목, 판정, 사유 }, ...]
 */
function 웹_체크결과저장(접수번호, results, 심사원) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);  // 동시 저장 충돌 방지

  try {
    const 시트 = _체크결과시트확보(ss);
    const D = 시트.getDataRange().getValues();
    const H = D[0];
    const iNo = H.indexOf('접수번호');
    const iId = H.indexOf('항목ID');

    // 기존 (접수번호+항목ID) → 행번호 매핑
    const 기존행 = {};
    for (let r = 1; r < D.length; r++) {
      if (D[r][iNo] === 접수번호) {
        기존행[D[r][iId]] = r + 1; // 시트 행번호 (1-base)
      }
    }

    const 시각 = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    results.forEach(res => {
      const 행값 = [접수번호, res.id, res.no, res.항목, res.판정, res.사유 || '', 심사원 || '', 시각];
      if (기존행[res.id]) {
        // 업데이트 (덮어쓰기지만 '체크결과' 시트 내에서만 — 원본 불변)
        시트.getRange(기존행[res.id], 1, 1, 행값.length).setValues([행값]);
      } else {
        시트.appendRow(행값);
      }
    });

    // 접수대장 상태를 '심사중'으로 (완료 판정이 아니면)
    _상태업데이트(ss, 접수번호, '심사중');

    return { ok: true, saved: results.length, time: 시각 };
  } finally {
    lock.releaseLock();
  }
}

/** 기능확인 시트 확보 */
function _기능확인시트확보(ss) {
  let 시트 = ss.getSheetByName(SHEET.기능확인);
  if (!시트) {
    시트 = ss.insertSheet(SHEET.기능확인);
    const 헤더 = 시트헤더정의[SHEET.기능확인];
    시트.getRange(1, 1, 1, 헤더.length).setValues([헤더])
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    시트.setFrozenRows(1);
  }
  return 시트;
}

/** 웹앱: 기능명세 표의 확인·비고 조회 (이어서 검토 가능) */
function 웹_기능확인조회(접수번호) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = _기능확인시트확보(ss);
  const out = {};
  if (시트.getLastRow() < 2) return out;
  const D = 시트.getDataRange().getValues();
  const H = D[0];
  D.slice(1).forEach(r => {
    if (r[H.indexOf('접수번호')] === 접수번호) {
      out[r[H.indexOf('ID')]] = {
        확인: r[H.indexOf('확인')],
        비고: r[H.indexOf('비고')],
      };
    }
  });
  return out;
}

/**
 * 웹앱: 기능명세 표 확인·비고 저장 (기타 행 포함)
 * items = [{ ID, 기능번호, 기능명, 확인(true/false), 비고 }, ...]
 * '기타' 항목은 ID = 접수번호 + '-ETC' 로 저장됨
 */
function 웹_기능확인저장(접수번호, items, 심사원) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const 시트 = _기능확인시트확보(ss);
    const D = 시트.getDataRange().getValues();
    const H = D[0];
    const iNo = H.indexOf('접수번호');
    const iId = H.indexOf('ID');

    const 기존행 = {};
    for (let r = 1; r < D.length; r++) {
      if (D[r][iNo] === 접수번호) 기존행[D[r][iId]] = r + 1;
    }

    const 시각 = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    (items || []).forEach(it => {
      const 행값 = [
        접수번호, it.ID, it.기능번호 || '', it.기능명 || '',
        it.확인 ? 'Y' : '', it.비고 || '', 심사원 || '', 시각,
      ];
      if (기존행[it.ID]) {
        시트.getRange(기존행[it.ID], 1, 1, 행값.length).setValues([행값]);
      } else {
        시트.appendRow(행값);
      }
    });

    return { ok: true, saved: (items || []).length, time: 시각 };
  } finally {
    lock.releaseLock();
  }
}

function _상태업데이트(ss, 접수번호, 상태) {
  const 시트 = ss.getSheetByName(SHEET.접수대장);
  const D = 시트.getDataRange().getValues();
  const H = D[0];
  const iNo = H.indexOf('접수번호');
  const iStat = H.indexOf('상태');
  for (let r = 1; r < D.length; r++) {
    if (D[r][iNo] === 접수번호) {
      const 현재 = D[r][iStat];
      // 이미 완료/반려면 건드리지 않음
      if (현재 !== '완료' && 현재 !== '반려') {
        시트.getRange(r + 1, iStat + 1).setValue(상태);
      }
      return;
    }
  }
}

// ═════════════════════════════════════════════════════════
// 9. 확인서 자동 생성 (우선순위 2)
// ═════════════════════════════════════════════════════════
//
// 체크 폼에서 저장된 '심사체크결과'를 끌어와 확인서 Docs를 생성합니다.
// 접수대장에서 행을 선택한 뒤 메뉴 > 확인서 생성 을 실행하거나,
// 웹앱 저장 직후 자동 호출할 수도 있습니다.

function 확인서생성() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getActiveSheet();
  if (시트.getName() !== SHEET.접수대장) {
    SpreadsheetApp.getUi().alert('접수대장 시트에서 행을 선택한 뒤 실행해주세요.');
    return;
  }
  const 행번호 = 시트.getActiveRange().getRow();
  if (행번호 <= 1) {
    SpreadsheetApp.getUi().alert('데이터 행을 선택해주세요.');
    return;
  }
  const H = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0];
  const D = 시트.getRange(행번호, 1, 1, 시트.getLastColumn()).getValues()[0];
  const 건 = {};
  H.forEach((h, i) => { 건[h] = D[i]; });

  const doc = _확인서Docs생성(ss, 건);
  SpreadsheetApp.getUi().alert(`확인서가 생성되었습니다.\n\n${doc.getUrl()}`);
}

/** 웹앱/메뉴 공용: 접수번호로 확인서 생성하고 URL 반환 */
function 확인서생성_byId(접수번호) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 대장 = ss.getSheetByName(SHEET.접수대장);
  const D = 대장.getDataRange().getValues();
  const H = D[0];
  const 행 = D.slice(1).find(r => r[H.indexOf('접수번호')] === 접수번호);
  if (!행) throw new Error('접수번호를 찾을 수 없습니다: ' + 접수번호);
  const 건 = {};
  H.forEach((h, i) => { 건[h] = 행[i]; });
  const doc = _확인서Docs생성(ss, 건);
  return { url: doc.getUrl(), name: doc.getName() };
}

function _확인서Docs생성(ss, 건) {
  const 접수번호 = 건['접수번호'];

  // 1) 체크 결과 조회 (항목ID → {판정, 사유})
  const 체크맵 = _체크결과조회(ss, 접수번호);

  // 2) 항목 정의 순서대로 결과 행 구성
  const 결과행 = 체크항목정의
    .filter(d => d.id !== 'C12') // 종합은 별도 처리
    .map(d => {
      const c = 체크맵[d.id] || {};
      return { no: d.no, 항목: d.항목, 판정: c.판정 || '미검토', 사유: c.사유 || '' };
    });

  // 3) 종합 판정 자동 계산
  const 종합 = _종합판정계산(결과행, 체크맵['C12']);

  // 4) Docs 생성
  const 제목 = `[기술심사확인서] ${접수번호} — ${건['제품명']}`;
  const doc = DocumentApp.create(제목);
  const body = doc.getBody();
  const 오늘 = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 MM월 dd일');

  // 제목
  body.appendParagraph('인공지능 제품 기술심사 확인서')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`접수번호 ${접수번호} · ${오늘}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setForegroundColor('#5f6368');
  body.appendParagraph('');

  // 1. 심사 대상
  body.appendParagraph('1. 심사 대상').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const t1 = body.appendTable([
    ['제품명', `${건['제품명'] || '-'} (${건['제품버전'] || '-'})`],
    ['신청기업', `${건['기업명'] || '-'} (사업자 ${건['사업자번호'] || '-'})`],
    ['제공형태', `${건['제공형태'] || '-'} · ${건['제품분류'] || '-'}`],
    ['담당심사원', 건['담당심사원'] || '-'],
  ]);
  _확인서표스타일(t1, true);

  // 2. 항목별 심사 결과
  body.appendParagraph('2. 항목별 심사 결과').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const 헤더 = ['번호', '심사 항목', '판정', '사유'];
  const 표데이터 = [헤더].concat(
    결과행.map(r => [r.no, r.항목, r.판정, r.사유 || '-'])
  );
  const t2 = body.appendTable(표데이터);
  _확인서표스타일(t2, false);
  _판정색칠(t2);

  // 3. 종합 의견
  body.appendParagraph('3. 종합 의견').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const t3 = body.appendTable([
    ['검토 현황', `총 ${결과행.length}개 항목 중 적합 ${종합.적합}건, 보완 ${종합.보완}건, 부적합 ${종합.부적합}건, 미검토 ${종합.미검토}건`],
    ['종합 판정', 종합.판정],
    ['종합 의견', (체크맵['C12'] && 체크맵['C12'].사유) ? 체크맵['C12'].사유 : ''],
  ]);
  _확인서표스타일(t3, true);

  // 서명란
  body.appendParagraph('');
  const sign = body.appendParagraph(`심사일: ${오늘}`);
  sign.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  const sign2 = body.appendParagraph(`담당심사원: ${건['담당심사원'] || ''}　(인)`);
  sign2.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  doc.saveAndClose();
  return doc;
}

/** 종합 판정 자동 계산 */
function _종합판정계산(결과행, 종합체크) {
  const 적합 = 결과행.filter(r => r.판정 === '적합').length;
  const 보완 = 결과행.filter(r => r.판정 === '보완').length;
  const 부적합 = 결과행.filter(r => r.판정 === '부적합').length;
  const 미검토 = 결과행.filter(r => r.판정 === '미검토').length;

  // 심사원이 종합(C12)을 직접 판정했으면 그 값 우선
  let 판정;
  if (종합체크 && 종합체크.판정 && 종합체크.판정 !== '미검토') {
    판정 = { '적합': '적합', '보완': '조건부 적합 (보완 후 재확인)', '부적합': '부적합' }[종합체크.판정];
  } else if (부적합 > 0) {
    판정 = '부적합';
  } else if (보완 > 0) {
    판정 = '조건부 적합 (보완 후 재확인)';
  } else if (미검토 > 0) {
    판정 = '심사 진행 중 (미검토 항목 있음)';
  } else {
    판정 = '적합';
  }
  return { 적합, 보완, 부적합, 미검토, 판정 };
}

function _확인서표스타일(표, 라벨열강조) {
  const 행수 = 표.getNumRows();
  for (let r = 0; r < 행수; r++) {
    const 행 = 표.getRow(r);
    for (let c = 0; c < 행.getNumCells(); c++) {
      const cell = 행.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
    }
    // 라벨열(첫 열) 회색 배경
    if (라벨열강조) {
      행.getCell(0).setBackgroundColor('#f1f3f4');
    }
  }
}

/** 항목별 결과표: 헤더 파랑 + 판정 컬럼 색칠 */
function _판정색칠(표) {
  // 헤더행
  const 헤더 = 표.getRow(0);
  for (let c = 0; c < 헤더.getNumCells(); c++) {
    헤더.getCell(c).setBackgroundColor('#1a73e8');
    const txt = 헤더.getCell(c).editAsText();
    txt.setForegroundColor('#ffffff').setBold(true);
  }
  // 판정 컬럼(인덱스 2) 색칠
  const 색 = { '적합': '#137333', '보완': '#e37400', '부적합': '#c5221f', '미검토': '#9aa0a6' };
  for (let r = 1; r < 표.getNumRows(); r++) {
    const cell = 표.getRow(r).getCell(2);
    const 판정 = cell.getText().trim();
    if (색[판정]) {
      cell.editAsText().setForegroundColor(색[판정]).setBold(true);
    }
  }
}

// ═════════════════════════════════════════════════════════
// 10. 증적 명세서 자동 생성 (증적자료용 — 신청서 전 항목 + 구조도 + 심사결과)
// ═════════════════════════════════════════════════════════
//
// 확인서(요약본)와 달리, 증적 명세서는 신청서의 모든 데이터 항목을
// 누락 없이 수록합니다. 데이터 구조도 이미지는 접수대장의
// '구조도파일명' 칸에 Drive 파일 ID가 있으면 끌어와 삽입합니다.

function 증적명세서생성() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getActiveSheet();
  if (시트.getName() !== SHEET.접수대장) {
    SpreadsheetApp.getUi().alert('접수대장 시트에서 행을 선택한 뒤 실행해주세요.');
    return;
  }
  const 행번호 = 시트.getActiveRange().getRow();
  if (행번호 <= 1) { SpreadsheetApp.getUi().alert('데이터 행을 선택해주세요.'); return; }

  const H = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0];
  const D = 시트.getRange(행번호, 1, 1, 시트.getLastColumn()).getValues()[0];
  const 건 = {};
  H.forEach((h, i) => { 건[h] = D[i]; });

  const doc = _증적명세서Docs생성(ss, 건);
  SpreadsheetApp.getUi().alert(`기술심사보고서가 생성되었습니다.\n\n${doc.getUrl()}`);
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
  const 연도 = new Date().getFullYear();

  // 데이터 수집
  const 기능목록 = _AI기능상세조회(ss, 접수번호);
  const 체크맵 = _체크결과조회(ss, 접수번호);
  const R = CONFIG.보고서 || {};
  const 문서번호 = `${R.문서번호접두 || 'AI심사'}-${접수번호}`;

  // 종합 판정 미리 계산 (첫 페이지에 필요)
  const 결과행 = 체크항목정의.filter(d => d.id !== 'C12').map(d => {
    const c = 체크맵[d.id] || {};
    return [d.no, d.항목, c.판정 || '미검토', c.사유 || '-'];
  });
  const 종합 = _종합판정계산(결과행.map(r => ({ 판정: r[2] })), 체크맵['C12']);

  const 제목문 = `[기술심사보고서] ${접수번호} — ${건['제품명']}`;
  const doc = DocumentApp.create(제목문);
  const body = doc.getBody();

  // ═══════════════ 표지 ═══════════════
  body.appendParagraph('인공지능 제품 기술심사보고서')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const 표지정보 = body.appendParagraph(
    `${R.작성기관 || ''}\n문서번호: ${문서번호}　|　보안등급: ${R.보안등급 || '-'}`
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
  ]);

  // ═══════════════ 2. 심사 대상 제품 ═══════════════
  _명세섹션(body, '2. 심사 대상 제품');
  _명세표(body, [
    ['제품 또는 서비스명', _v(건['제품명'])],
    ['제품 버전', _v(건['제품버전'])],
    ['제공 형태', _v(건['제공형태'])],
    ['인공지능 제품 분류', _v(건['제품분류'])],
    ['세부품명번호', _v(건['세부품명번호'])],
    ['물품식별번호', _v(건['물품식별번호'])],
    ['인공지능 기능 수', _v(건['인공지능기능수'])],
  ]);

  // ═══════════════ 3. 신청기업 현황 ═══════════════
  _명세섹션(body, '3. 신청기업 현황');
  _명세표(body, [
    ['기업명', _v(건['기업명'])],
    ['사업자등록번호', _v(건['사업자번호'])],
    ['대표자명', _v(건['대표자'])],
    ['담당자', `${_v(건['담당자명'])} · ${_v(건['연락처'])} · ${_v(건['이메일'])}`],
    ['소재지', _v(건['소재지'])],
  ]);

  // ═══════════════ 4. 인공지능 기능 명세 ═══════════════
  _명세섹션(body, '4. 인공지능 기능 명세');
  if (기능목록.length) {
    const 표 = [['번호', '기능명', '인공지능역할', '구현방식']];
    기능목록.forEach(f => 표.push([
      _v(f['기능번호']), _v(f['기능명']),
      _v(f['인공지능역할']), _구현방식레이블(f['구현방식']),
    ]));
    _명세표헤더(body.appendTable(표));
  } else {
    body.appendParagraph('(인공지능 기능 데이터 없음)').editAsText().setForegroundColor('#9aa0a6');
  }

  // ═══════════════ 5. 종합 심사 의견 ═══════════════
  _명세섹션(body, '5. 종합 심사 의견');
  _명세표(body, [
    ['검토 현황', `총 ${결과행.length}개 항목 — 적합 ${종합.적합} · 보완 ${종합.보완} · 부적합 ${종합.부적합} · 미검토 ${종합.미검토}`],
    ['종합 판정', 종합.판정],
    ['보완 요구사항', _보완요약(결과행)],
    ['심사원 종합 소견', (체크맵['C12'] && 체크맵['C12'].사유) ? 체크맵['C12'].사유 : ''],
  ]);

  // ═══════════════ 6. 심사 결론 및 서명 ═══════════════
  _명세섹션(body, '6. 심사 결론 및 서명');
  const 결론표 = _명세표(body, [
    ['최종 판정', 종합.판정],
    ['심사일', 오늘],
    ['담당심사원', `${_v(건['담당심사원'])}　　(서명/인)`],
    [`${R.책임자직위 || '심사책임자'}`, '　　(서명/인)'],
  ]);

  // ═══════════════ [붙임] 새 페이지부터 ═══════════════
  body.appendPageBreak();
  body.appendParagraph('붙 임')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // ── 붙임1. 기능별 인공지능 구현 세부 ──
  _명세섹션(body, '붙임1. 기능별 인공지능 구현 세부');
  if (기능목록.length) {
    기능목록.forEach(f => {
      body.appendParagraph(`▸ ${_v(f['기능번호'])}. ${_v(f['기능명'])}`)
        .setHeading(DocumentApp.ParagraphHeading.HEADING3);
      _명세표(body, [
        ['구현 방식', _구현방식레이블(f['구현방식'])],
        ['인공지능역할', _v(f['인공지능역할'])],
        ['입력 데이터', _v(f['입력데이터설명'])],
        ['출력 데이터', _v(f['출력데이터설명'])],
        ['연산 자원', _v(f['연산자원'])],
        ['실행 환경', _v(f['실행환경'])],
        ['학습 데이터 사양', _v(f['학습데이터사양'])],
        ['재현환경·라이선싱', _v(f['재현환경라이선싱'])],
        ['Base Model', _v(f['BaseModel'])],
        ['튜닝 방법', _v(f['튜닝방법'])],
        ['외부 API·모델', _v(f['외부API모델'])],
        ['타겟 HW·OS', _v(f['타겟HW_OS'])],
        ['배포 포맷·정밀도', _v(f['배포포맷정밀도'])],
        ['혼합 구성 설명', _v(f['혼합구성설명'])],
      ]);
    });
  } else {
    body.appendParagraph('(데이터 없음)').editAsText().setForegroundColor('#9aa0a6');
  }

  // ── 붙임2. 데이터 구조도 (새 페이지) ──
  body.appendPageBreak();
  _명세섹션(body, '붙임2. 데이터 구조도');
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

  // ── 붙임3. 심사 항목별 검토 결과 (새 페이지) ──
  body.appendPageBreak();
  _명세섹션(body, '붙임3. 심사 항목별 검토 결과');
  const 결과표 = [['번호', '심사 항목', '판정', '검토 의견']].concat(결과행);
  const t결과 = body.appendTable(결과표);
  _명세표헤더(t결과);
  _판정색칠(t결과);

  // ── 붙임4. 첨부자료 목록 ──
  _명세섹션(body, '붙임4. 첨부자료 목록');
  _명세표(body, [
    ['기존 인증·시험 결과', _v(건['비고'])],
    ['구조도 원본 파일', 구조도원본 ? `${구조도원본.split(/[,\n]+/).filter(Boolean).length}건 (Drive→NAS 보관)` : '(없음)'],
  ]);

  doc.saveAndClose();
  return doc;
}

// 텍스트 요약 (길면 잘라서 …)
function _요약(val, 길이) {
  const s = String(val == null ? '' : val).trim();
  if (s === '') return '(미기재)';
  return s.length > 길이 ? s.slice(0, 길이) + '…' : s;
}

// 보완·부적합 항목만 모아 요약
function _보완요약(결과행) {
  const 보완들 = 결과행.filter(r => r[2] === '보완' || r[2] === '부적합');
  if (!보완들.length) return '없음';
  return 보완들.map(r => `[${r[0]}] ${r[3]}`).join('\n');
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
    }
  }
  return t;
}

function _명세표헤더(t) {
  const h = t.getRow(0);
  for (let c = 0; c < h.getNumCells(); c++) {
    h.getCell(c).setBackgroundColor('#1a73e8');
    h.getCell(c).editAsText().setForegroundColor('#ffffff').setBold(true);
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

// ═════════════════════════════════════════════════════════
// 12. 증적 3종 산출 (DOCX + PDF + JSON) → 보관 폴더 정리
// ═════════════════════════════════════════════════════════
//
// 명세서 Docs를 생성한 뒤 docx/pdf로 변환하고, 신청서 전 항목 +
// 심사 결과를 담은 JSON을 만들어, 접수번호별 보관 폴더에 모읍니다.
// 구조도 원본도 같은 폴더로 복사하여 NAS Cloud Sync가 한 번에 수거하게 합니다.

function 증적3종산출() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getActiveSheet();
  if (시트.getName() !== SHEET.접수대장) {
    SpreadsheetApp.getUi().alert('접수대장 시트에서 행을 선택한 뒤 실행해주세요.');
    return;
  }
  const 행번호 = 시트.getActiveRange().getRow();
  if (행번호 <= 1) { SpreadsheetApp.getUi().alert('데이터 행을 선택해주세요.'); return; }

  const H = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0];
  const D = 시트.getRange(행번호, 1, 1, 시트.getLastColumn()).getValues()[0];
  const 건 = {};
  H.forEach((h, i) => { 건[h] = D[i]; });

  const res = _증적3종산출_처리(ss, 건);
  SpreadsheetApp.getUi().alert(
    `증적 3종 산출 완료\n\n` +
    `폴더: ${res.folderUrl}\n` +
    `DOCX · PDF · JSON + 구조도 원본 ${res.구조도복사}건\n\n` +
    `NAS Cloud Sync가 곧 수거합니다.`
  );
}

function 증적3종산출_byId(접수번호) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 대장 = ss.getSheetByName(SHEET.접수대장);
  const D = 대장.getDataRange().getValues();
  const H = D[0];
  const 행 = D.slice(1).find(r => r[H.indexOf('접수번호')] === 접수번호);
  if (!행) throw new Error('접수번호를 찾을 수 없습니다: ' + 접수번호);
  const 건 = {};
  H.forEach((h, i) => { 건[h] = 행[i]; });
  return _증적3종산출_처리(ss, 건);
}

function _증적3종산출_처리(ss, 건) {
  const 접수번호 = 건['접수번호'];

  // 1) 보관 폴더 준비: 보관루트 / 접수번호 /
  const 폴더 = _보관폴더준비(접수번호);

  // 2) 명세서 Docs 생성 (앞서 만든 함수 재사용)
  const doc = _증적명세서Docs생성(ss, 건);
  const docId = doc.getId();
  const 기본이름 = `증적명세서_${접수번호}`;

  const token = ScriptApp.getOAuthToken();

  // 3) DOCX 변환 (편집용)
  try {
    const docxUrl = `https://www.googleapis.com/drive/v3/files/${docId}/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document`;
    const docxBlob = UrlFetchApp.fetch(docxUrl, {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true,
    }).getBlob().setName(기본이름 + '.docx');
    _폴더에저장(폴더, docxBlob);
  } catch (e) { Logger.log('DOCX 변환 실패: ' + e.message); }

  // 4) PDF 변환 (확정 증적본)
  try {
    const pdfBlob = doc.getAs('application/pdf').setName(기본이름 + '.pdf');
    _폴더에저장(폴더, pdfBlob);
  } catch (e) { Logger.log('PDF 변환 실패: ' + e.message); }

  // 5) JSON 생성 (재활용·체크앱 연동용)
  const json = _증적JSON생성(ss, 건);
  const jsonBlob = Utilities.newBlob(
    JSON.stringify(json, null, 2), 'application/json', 기본이름 + '.json'
  );
  _폴더에저장(폴더, jsonBlob);

  // 6) 명세서 Docs 원본은 임시 — 변환 후 휴지통으로 (편집은 docx로)
  //    편집 가능한 Docs를 남기고 싶으면 아래 줄을 주석 처리하세요.
  DriveApp.getFileById(docId).setTrashed(true);

  // 7) 구조도 원본 복사
  let 구조도복사 = 0;
  const 원본ID들 = String(건['구조도파일명'] || '')
    .split(/[,\n]+/).map(s => _드라이브ID추출(s.trim())).filter(Boolean);
  원본ID들.forEach(id => {
    try {
      const f = DriveApp.getFileById(id);
      if (!폴더.getFilesByName(f.getName()).hasNext()) {
        f.makeCopy(f.getName(), 폴더);
        구조도복사++;
      }
    } catch (e) { Logger.log('구조도 복사 실패: ' + id + ' / ' + e.message); }
  });

  return { folderUrl: 폴더.getUrl(), 구조도복사: 구조도복사 };
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

/** 같은 이름 파일이 있으면 덮어쓰기(이전 휴지통) 후 저장 */
function _폴더에저장(폴더, blob) {
  const 이름 = blob.getName();
  const 기존 = 폴더.getFilesByName(이름);
  while (기존.hasNext()) { 기존.next().setTrashed(true); }
  폴더.createFile(blob);
}

/**
 * 증적 JSON 생성 — 신청서 전 항목 + AI 기능 + 심사 체크 결과
 * (원래 체크 앱이 받던 구조와 호환되도록 구성)
 */
function _증적JSON생성(ss, 건) {
  const 접수번호 = 건['접수번호'];
  const 기능목록 = _AI기능상세조회(ss, 접수번호);
  const 체크맵 = _체크결과조회(ss, 접수번호);

  // 심사 결과 배열
  const 심사결과 = 체크항목정의.map(d => {
    const c = 체크맵[d.id] || {};
    return {
      항목ID: d.id,
      항목번호: d.no,
      항목: d.항목,
      판정: c.판정 || '미검토',
      사유: c.사유 || '',
    };
  });
  const 종합 = _종합판정계산(
    심사결과.filter(r => r.항목ID !== 'C12'),
    체크맵['C12']
  );

  return {
    schemaVersion: '1.1',
    generatedAt: new Date().toISOString(),
    접수번호: 접수번호,
    상태: 건['상태'] || '',

    제품정보: {
      제품명: 건['제품명'] || '',
      제품버전: 건['제품버전'] || '',
      제공형태: 건['제공형태'] || '',
      인공지능기능수: 건['인공지능기능수'] || '',
    },
    기업정보: {
      기업명: 건['기업명'] || '',
      사업자번호: 건['사업자번호'] || '',
      대표자: 건['대표자'] || '',
      담당자명: 건['담당자명'] || '',
      연락처: 건['연락처'] || '',
      이메일: 건['이메일'] || '',
      소재지: 건['소재지'] || '',
    },
    제품분류: {
      제품분류: 건['제품분류'] || '',
      세부품명번호: 건['세부품명번호'] || '',
      물품식별번호: 건['물품식별번호'] || '',
    },
    인공지능기능: 기능목록.map(f => ({
      기능번호: f['기능번호'] || '',
      기능명: f['기능명'] || '',
      인공지능역할: f['인공지능역할'] || '',
      입력: f['입력'] || '',
      출력: f['출력'] || '',
      설명서참조위치: f['설명서참조위치'] || '',
      구현방식: f['구현방식'] || '',
      연산자원: f['연산자원'] || '',
      실행환경: f['실행환경'] || '',
      학습데이터사양: f['학습데이터사양'] || '',
      재현환경라이선싱: f['재현환경라이선싱'] || '',
      BaseModel: f['BaseModel'] || '',
      튜닝방법: f['튜닝방법'] || '',
      외부API모델: f['외부API모델'] || '',
      타겟HW_OS: f['타겟HW_OS'] || '',
      배포포맷정밀도: f['배포포맷정밀도'] || '',
      혼합구성설명: f['혼합구성설명'] || '',
      입력데이터설명: f['입력데이터설명'] || '',
      출력데이터설명: f['출력데이터설명'] || '',
    })),
    기존인증시험: 건['비고'] || '',
    구조도파일명: 건['구조도파일명'] || '',

    심사결과: 심사결과,
    종합판정: {
      적합: 종합.적합, 보완: 종합.보완, 부적합: 종합.부적합, 미검토: 종합.미검토,
      판정: 종합.판정,
      종합의견: (체크맵['C12'] && 체크맵['C12'].사유) ? 체크맵['C12'].사유 : '',
    },
    담당심사원: 건['담당심사원'] || '',
  };
}