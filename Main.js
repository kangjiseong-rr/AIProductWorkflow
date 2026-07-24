/**
 * ============================================================
 *  인공지능 제품 기술심사 관리 시스템 — Google Apps Script
 *  대상: Google Sheets (스프레드시트에 이 스크립트를 붙여넣기)
 * ============================================================
 *
 *  이 파일은 설정값(CONFIG/SHEET)·시트 표준 헤더 정의·초기 설정/컬럼
 *  마이그레이션·메뉴 진입점(onOpen)만 담당합니다. 나머지 기능은 책임별로
 *  분리돼 있습니다(Apps Script는 모든 .js 파일이 전역 스코프를 공유하므로
 *  import 없이 서로 참조합니다):
 *   - FieldDefinition.js    — 신청서 필드 스키마(엑셀 별칭 + JSON path 공용)
 *   - ExcelParser.js        — 엑셀 파일 → 표준 객체 변환
 *   - JsonAdapter.js        — JSON 파일 → 표준 객체 변환
 *   - SheetWriter.js        — 표준 객체 → 접수대장·제품모델·기능상세 시트 기록
 *   - Schedule.js           — 일정관리 동기화 + 공휴일·마감예정일 수식
 *   - ReviewerAssignment.js — 심사원 배정 + Google Chat 알림
 *   - ReportGenerator.js    — 기술심사보고서(Google Docs) 생성
 *
 *  [사용 방법]
 *  1. Google Sheets 새 파일 생성
 *  2. 확장 프로그램 > Apps Script > 이 프로젝트 파일 전체 붙여넣기
 *  3. 저장 후 '초기설정실행' 함수 한 번 실행 (시트 구조 자동 생성)
 *  4. 이후 엑셀 수신 시 '엑셀파싱등록' 함수 실행 (메뉴에서도 실행 가능)
 * ============================================================
 */
// ─────────────────────────────────────────────
// 0. 설정값 (환경에 맞게 수정)
// ─────────────────────────────────────────────
function loadConfig() {
  const baseConfig = {
    // 기본 심사 기간 (일)
    기본심사기간: 14,

    // Google Drive 폴더 ID (엑셀 파일을 업로드할 폴더, 빈 문자열이면 루트)
    드라이브폴더ID: '',

    // ── Google Chat Webhook ──────────────────────
    // Space별 Webhook URL을 설정합니다.
    챗_공통Webhook: '',

    // ── 기술심사보고서 — 기관 정보 (한 번 설정하면 모든 보고서에 적용) ──
    보고서: {
      작성기관: '한국정보통신기술협회(TTA)',
      문서번호접두: 'TTA-AI심사',          // 예: TTA-AI심사-{접수번호}-01
      심사근거: '(과학기술정보통신부 고시 제2026-50호) 인공지능 제품·서비스 확인 절차 운영에 관한 고시',
      심사방법: '제출 기술자료 검토 및 확인 기준에 따른 항목별 적합성 검토',
      보안등급: '대외제한',                 // 표지 표기
      책임자직위: '심사책임자',
    },

    // ── 관리자 이메일 목록 (초기화·재생성 등 위험 메뉴를 볼 수 있는 계정) ──
    관리자이메일: [],
  };

  try {
    // Google Apps Script 스크립트 속성에서 동적으로 로드
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    if (properties.DRIVE_FOLDER_ID) {
      baseConfig.드라이브폴더ID = properties.DRIVE_FOLDER_ID;
    }
    if (properties.CHAT_COMMON_WEBHOOK) {
      baseConfig.챗_공통Webhook = properties.CHAT_COMMON_WEBHOOK;
    }
    if (properties.ADMIN_EMAILS) {
      // 쉼표로 구분된 관리자 이메일 목록 파싱
      baseConfig.관리자이메일 = properties.ADMIN_EMAILS.split(',').map(function(s) { return s.trim(); });
    }
  } catch (e) {
    if (typeof Logger !== 'undefined') {
      Logger.log("설정 로드 실패 (로컬 디버깅 또는 권한 부족): " + e.toString());
    }
  }

  return baseConfig;
}

const CONFIG = loadConfig();

// 시트 이름 상수
const SHEET = {
  접수대장: '접수대장',
  제품모델: '인공지능제품모델',        // 신규 — 세부품명번호/물품식별번호 반복행
  AI기능상세: '인공지능기능상세',      // 표시명만 변경 (코드상 키는 호환 유지)
  일정관리: '일정관리',
  로그: '파싱로그',
  심사원관리: '심사원관리',
  배정알림로그: '배정알림로그',
};


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
    '제품명', '제품수', '제공형태', '제품분류',
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
    // ── 심사 결과 (보고서 5. 종합 심사 의견의 원본 데이터) ──
    '종합판정', '심사완료일', '심사의견',
  ],
  [SHEET.제품모델]: [
    '접수번호', '연번', '모델명', '세부품명번호', '물품식별번호',
  ],
  // AI기능상세 컬럼은 실제 신청 폼(확인신청 웹페이지)의 "핵심 인공지능 기능 명세" +
  // "AI 기능 구현 방식(A~E)" 섹션과 1:1로 맞춘 구성입니다. 신청 시스템의 실제 제출 데이터(JSON) 구조를
  // 대조해 확인한 필드 목록으로, 실제로는 구현방식과 무관하게 모든 하위 필드를 한 행에 함께 갖고
  // 있다가(선택된 구현방식 외 칸은 공란) 보고서에서만 구현방식별로 골라 보여줍니다.
  //  · 공통 요약: 기능번호·기능명·인공지능역할·입력·출력·레퍼런스참조위치 (보고서 4번)
  //  · 구현방식 공통: 구현방식·연산자원요약·실행환경요약(선택)·입력데이터설명·출력데이터설명·기타참고자료파일명 (보고서 붙임1)
  //  · 구현방식별 전용 필드 (폼에서 A~E 선택 시에만 노출되는 항목):
  //      A. 자체 학습 모델   → 학습데이터사양, 개발환경라이브러리알고리즘
  //      B. 오픈소스 모델    → BaseModel명칭, 튜닝방법, 튜닝데이터셋
  //      C. 외부 AI API 연동 → 외부API정보
  //      D. 온디바이스 AI    → 타겟HW_OS, 추론런타임
  //      E. 혼합형 AI 구성   → 혼합구성설명, 모델별역할및입출력흐름
  [SHEET.AI기능상세]: [
    '접수번호', '기능번호', '기능명',
    '인공지능역할', '입력', '출력', '레퍼런스참조위치',
    '구현방식', '연산자원요약', '실행환경요약',
    '학습데이터사양', '개발환경라이브러리알고리즘',
    'BaseModel명칭', '튜닝방법', '튜닝데이터셋',
    '외부API정보',
    '타겟HW_OS', '추론런타임',
    '혼합구성설명', '모델별역할및입출력흐름',
    '입력데이터설명', '출력데이터설명',
    '기타참고자료파일명',
  ],
  [SHEET.일정관리]: [
    // ── 일정·기한 ─────────────────────────────────
    '순번',          // 표시용 일련번호 (자동)
    '접수번호',
    '신청일',        // 기산일 (신청서 원본)
    '심사접수일',    // TTA 인수일
    '마감예정일',    // 신청일 + 15 WD, 주말·공휴일 제외 (수식 자동 생성)
    '상태',          // 대기 / 심사중 / 보완 / 완료(적합) / 종료(부적합)
    '보완요청일',    // 내부 심사원이 직접 입력
    '연장마감일',    // 보완요청일 + 30 WD, 주말·공휴일 제외 (수식 자동 생성)
    '담당심사원',
    '특이사항',      // 모든 사용자가 작성하는 내부 자유 메모·의견
    // ── 신청기업 ──────────────────────────────────
    '기업명', '담당자명', '연락처', '이메일', '소재지',
    // ── 제품·서비스 핵심 정보 ─────────────────────
    '제품명', '제품수', '개요', '제공형태', '제품분류',
    '인공지능적용목적', '인공지능적용범위',
    // ── 제출 현황 ─────────────────────────────────
    '명세서작성방식', '기타제출서류여부', '보유인증',
    // ── 심사 참고 ─────────────────────────────────
    '인공지능기능수',
  ],
  [SHEET.로그]: [
    '일시', '파일명', '접수번호', '결과', '메시지',
  ],
  [SHEET.심사원관리]: [
    '번호', '심사원명', '이메일', '활성여부', '배정순서', 'Chat사용자ID',
  ],
  [SHEET.배정알림로그]: [
    '발송일시', '접수번호', '담당심사원', '이메일', '발송결과', '실행자',
  ],
};

// 구버전 시트명 → 신버전 시트명 (탭 이름만 바꿔서 기존 데이터 그대로 보존)
const 시트이름마이그레이션맵 = {
  'AI기능상세': SHEET.AI기능상세,
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
 * 기존 시트에 새 표준 컬럼이 없으면 표준 헤더 순서의 위치에 삽입합니다.
 * Google Sheets의 열 삽입은 기존 셀·수식 참조를 함께 이동하므로 데이터가 보존됩니다.
 * 행 쓰기는 헤더 이름 기반이라 열 위치가 바뀌어도 안전합니다.
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
    '설명서참조위치': '레퍼런스참조위치',
    '연산자원': '연산자원요약',
    '실행환경': '실행환경요약',
    'BaseModel': 'BaseModel명칭',
    '외부API모델': '외부API정보',
    '배포포맷정밀도': '추론런타임',
    // 지난 마이그레이션에서 잠시 쓰였던 이름 → 실제 신청 시스템 데이터로 확인한 정식 필드명
    '디바이스HW사양': '타겟HW_OS',
    '소프트웨어스택미들웨어': '추론런타임',
    '세부구성요소별설명': '모델별역할및입출력흐름',
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
    누락.forEach(새헤더 => {
      const 표준 = 시트헤더정의[이름];
      const 표준위치 = 표준.indexOf(새헤더);
      기존 = 시트.getRange(1, 1, 1, 시트.getLastColumn())
        .getValues()[0].map(v => String(v).trim());

      // 표준상 가장 가까운 앞 컬럼 뒤에 삽입. 앞 컬럼이 없으면 가장 가까운 뒤 컬럼 앞에 삽입.
      const 앞헤더 = 표준.slice(0, 표준위치).reverse().find(h => 기존.includes(h));
      const 뒤헤더 = 표준.slice(표준위치 + 1).find(h => 기존.includes(h));
      let 삽입열;
      if (앞헤더) {
        삽입열 = 기존.indexOf(앞헤더) + 2; // 1-based: 앞 컬럼 바로 다음
        시트.insertColumnAfter(삽입열 - 1);
      } else if (뒤헤더) {
        삽입열 = 기존.indexOf(뒤헤더) + 1;
        시트.insertColumnBefore(삽입열);
      } else {
        삽입열 = 시트.getLastColumn() + 1;
        시트.insertColumnAfter(시트.getLastColumn());
      }
      시트.getRange(1, 삽입열).setValue(새헤더)
        .setBackground(이름 === SHEET.일정관리 ? 일정관리_헤더색 : '#1a73e8')
        .setFontColor('#ffffff').setFontWeight('bold');
    });
    if (누락.length) {
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
    if (이름 === SHEET.일정관리) {
      _일정관리레거시컬럼삭제_(시트);
      _일정관리헤더색적용_(시트);
    }
  });
  if (추가내역.length) {
    try {
      SpreadsheetApp.getUi().alert('컬럼 마이그레이션 완료:\n\n' + 추가내역.join('\n'));
    } catch (e) { /* UI 없는 환경 */ }
  }
  _전체제품모델요약갱신(ss);
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

  // 1-1) 메인 시트(일정관리)를 맨 앞으로 이동
  const 메인시트 = ss.getSheetByName(SHEET.일정관리);
  if (메인시트) {
    ss.setActiveSheet(메인시트);
    ss.moveActiveSheet(1);
  }

  // 2) 기존 시트에 새 표준 컬럼 자동 추가
  try { 헤더마이그레이션(); } catch (e) { /* UI 없는 환경에서 실행될 수 있음 */ }

  // 3) 컬럼매핑 메타 시트 (엑셀 헤더 별칭을 코드 수정 없이 관리)
  _컬럼매핑시트확보(ss);
  _심사원관리설정_(ss);

  // ── 일정관리 시트 후처리 ─────────────────────────────
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 일정H = 일정시트.getRange(1, 1, 1, 일정시트.getLastColumn())
    .getValues()[0].map(v => String(v).trim());
  const iD마감 = 일정H.indexOf('마감예정일') + 1;
  const iD상태 = 일정H.indexOf('상태') + 1;
  const iD보완요청 = 일정H.indexOf('보완요청일') + 1;
  const iD연장마감 = 일정H.indexOf('연장마감일') + 1;
  const iD담당심사원 = 일정H.indexOf('담당심사원') + 1;
  const 끝열문자 = columnLetter(일정H.length);

  // 마감예정일 수식은 _일정관리행추가()에서 행별로 설정
  // (컬럼 전체를 미리 채우면 appendRow가 빈 행을 못 찾아 밀리므로 사전 채움 안 함)
  _마감예정일수식갱신_(ss);

  // 기존 '완료' 값은 새 상태명으로 마이그레이션한 뒤 상태 드롭다운을 갱신
  const 상태범위 = 일정시트.getRange(2, iD상태, 999, 1);
  상태범위.createTextFinder('완료').matchEntireCell(true).replaceAllWith('완료(적합)');
  상태범위
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['대기', '심사중', '보완', '완료(적합)', '종료(부적합)'], true).build());
  const 심사원관리시트 = ss.getSheetByName(SHEET.심사원관리);
  const 심사원관리H = 심사원관리시트
    ? 심사원관리시트.getRange(1, 1, 1, 심사원관리시트.getLastColumn()).getValues()[0].map(v => String(v).trim())
    : [];
  const 심사원명열 = 심사원관리H.indexOf('심사원명') + 1;
  if (iD담당심사원 > 0 && 심사원관리시트 && 심사원명열 > 0) {
    const 심사원명범위 = 심사원관리시트.getRange(
      2,
      심사원명열,
      Math.max(1, 심사원관리시트.getMaxRows() - 1),
      1
    );
    일정시트.getRange(2, iD담당심사원, 999, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(심사원명범위, true)
        .setAllowInvalid(false)
        .setHelpText('심사원관리 시트에 등록된 심사원을 선택하세요. 배정·발송 시 활성여부가 Y인지 확인합니다.')
        .build());
  }
  const 날짜입력규칙 = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('날짜를 직접 입력하거나 캘린더에서 선택하세요.')
    .build();
  if (iD보완요청 > 0) {
    // 상태 열 오른쪽에 삽입될 때 복제된 상태 드롭다운을 제거하고 날짜 입력 열로 재설정
    일정시트.getRange(2, iD보완요청, 999, 1)
      .clearDataValidations()
      .setDataValidation(날짜입력규칙)
      .setNumberFormat('yy-mm-dd');
  }
  if (iD연장마감 > 0) {
    // 자동 계산 열이므로 수기 입력용 유효성 검사를 두지 않음
    일정시트.getRange(2, iD연장마감, 999, 1)
      .clearDataValidations()
      .setNumberFormat('yy-mm-dd');
  }

  // ── 조건부서식 (톤다운 색상, 행 전체) ──
  // 규칙 우선순위: 위에서부터 먼저 적용됨.
  //   ① 기한 초과(미완료)  → 연빨강   (상태색보다 우선)
  //   ② 완료(적합)         → 연녹색
  //   ③ 종료(부적합)       → 진한 회색
  //   ④ 보완               → 머스터드(짙은 노랑)
  //   ⑤ 심사중             → 연노랑
  //   ⑥ 대기               → 무색 (규칙 없음)
  const 마감열문자 = columnLetter(iD마감);
  const 상태열문자 = columnLetter(iD상태);
  const 보완요청열문자 = columnLetter(iD보완요청);
  const 연장마감열문자 = columnLetter(iD연장마감);
  const 전체범위 = 일정시트.getRange(`A2:${끝열문자}1000`);

  // ① 기한 초과 & 미완료
  const 초과 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      `=AND(NOT(OR($${상태열문자}2="완료",$${상태열문자}2="완료(적합)",$${상태열문자}2="종료(부적합)")),IF($${보완요청열문자}2<>"",AND($${연장마감열문자}2<>"",$${연장마감열문자}2<TODAY()),AND($${마감열문자}2<>"",$${마감열문자}2<TODAY())))`
    )
    .setBackground('#F4CCCC').setFontColor('#990000')
    .setRanges([전체범위]).build();

  // ② 완료(적합) → 연녹색 (구버전 '완료' 값도 호환)
  const 완료 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR($${상태열문자}2="완료(적합)",$${상태열문자}2="완료")`)
    .setBackground('#D9EAD3').setFontColor('#38761D')
    .setRanges([전체범위]).build();

  // ③ 종료(부적합) → 진한 회색
  const 종료 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${상태열문자}2="종료(부적합)"`)
    .setBackground('#666666').setFontColor('#FFFFFF')
    .setRanges([전체범위]).build();

  // ④ 보완 → 머스터드(짙은 노랑)
  const 보완 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${상태열문자}2="보완"`)
    .setBackground('#F9CB9C').setFontColor('#783F04')
    .setRanges([전체범위]).build();

  // ⑤ 심사중 → 연노랑
  const 심사중 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${상태열문자}2="심사중"`)
    .setBackground('#FCE8B2').setFontColor('#7F6000')
    .setRanges([전체범위]).build();

  // 대기(무색)는 규칙 없음
  일정시트.setConditionalFormatRules([초과, 완료, 종료, 보완, 심사중]);

  _일정관리서식적용_(일정시트, true);

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

  // ── 접수대장 종합판정 드롭다운 ────────────────────
  const 대장판정열 = 대장H.indexOf('종합판정') + 1;
  if (대장판정열 > 0) {
    const 대장판정범위 = 대장시트.getRange(2, 대장판정열, 999, 1);
    const 대장판정규칙 = SpreadsheetApp.newDataValidation()
      .requireValueInList(['적합', '보완요청', '부적합(미충족)'], true)
      .build();
    대장판정범위.setDataValidation(대장판정규칙);
  }

  SpreadsheetApp.getUi().alert('초기설정 완료! 심사원관리·배정알림로그를 포함한 시트 구성이 갱신되었습니다.');
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
    .setBackground(이름 === SHEET.일정관리 ? 일정관리_헤더색 : '#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function _드라이브ID추출(입력) {
  const match = 입력.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : 입력;
}


// ─────────────────────────────────────────────
// 6. 사용자 메뉴 등록
// ─────────────────────────────────────────────
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const 관리자 = _관리자여부();

  const menu = ui.createMenu('🔍 인공지능심사관리')
    // ── 심사원 공통 메뉴 ──
    .addItem('📑 기술심사보고서 생성 (선택 행)', '증적명세서생성');

  // ── 관리자 전용 메뉴 (관리자가 아니면 아예 표시 안 함) ──
  if (관리자) {
    menu.addSeparator()
      .addItem('📁 엑셀 파일 등록 (파싱)', '엑셀파싱등록')
      .addSeparator()
      .addItem('👥 심사원 Chat ID 일괄 갱신', '심사원Chat사용자ID일괄갱신')
      .addItem('🔄 선택 행 라운드로빈 추천', '선택행라운드로빈추천')
      .addItem('📨 선택 행 배정 확정·알림 발송', '선택행배정확정알림발송')
      .addItem('🔁 선택 행 배정 알림 재발송', '선택행배정알림재발송')
      .addSeparator()
      .addItem('⚙️ 초기 설정·컬럼 갱신', '초기설정실행')
      .addItem('📅 공휴일·마감예정일 갱신', '마감예정일갱신')
      .addItem('💬 Chat 알림 테스트', '챗알림테스트');
  }

  menu.addToUi();
}

/**
 * 실행 계정이 관리자인지 판단.
 * CONFIG.관리자이메일이 비어 있으면(초기 상태) 스프레드시트 소유자만 관리자로 간주.
 * ※ 메뉴 숨김용 편의 판정이며, 완전한 보안은 아님.
 */
function _관리자여부() {
  let 이메일 = '';
  try { 이메일 = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}

  const 목록 = (CONFIG.관리자이메일 || []).map(e => String(e).toLowerCase().trim()).filter(Boolean);

  // 관리자 목록이 지정돼 있으면 그 목록으로 판정
  if (목록.length > 0) {
    return 이메일 !== '' && 목록.indexOf(이메일) >= 0;
  }

  // 목록이 비어 있으면(설정 전) 스프레드시트 소유자를 관리자로 간주
  try {
    const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
    const ownerEmail = owner ? (owner.getEmail() || '').toLowerCase() : '';
    if (ownerEmail && 이메일) return 이메일 === ownerEmail;
  } catch (e) {}

  // 판정 불가(도메인 밖 계정 등)면 안전하게 비관리자 처리
  return false;
}
