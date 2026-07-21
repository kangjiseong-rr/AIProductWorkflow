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
function loadConfig() {
  const baseConfig = {
    // 담당자 목록 (라운드로빈 자동 배분)
    담당자목록: ['홍길동', '김심사', '이검토', '박분석'],

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
    if (properties.ASSIGNEES) {
      // 쉼표로 구분된 담당자 목록 파싱
      baseConfig.담당자목록 = properties.ASSIGNEES.split(',').map(function(s) { return s.trim(); });
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

const 일정관리_헤더색 = '#1f3a5f';
const 일정관리_특이사항헤더색 = '#dbe7f3';
const 일정관리_특이사항헤더글자색 = '#1f3a5f';
const 보고서_표헤더색 = '#f1f3f3';
const 보고서_표헤더글자색 = '#202124';
const 공휴일시트명 = '공휴일';
const 대한민국공휴일캘린더URL = 'https://calendar.google.com/calendar/ical/ko.south_korea%23holiday%40group.v.calendar.google.com/public/basic.ics';

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
    '심사원명', '이메일', '활성여부', '배정순서', 'Chat사용자ID',
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

/** 더 이상 쓰지 않는 옛 컬럼(심사링크, 보고서생성)이 남아있으면 완전히 삭제 */
function _일정관리레거시컬럼삭제_(시트) {
  ['심사링크', '보고서생성'].forEach(옛컬럼명 => {
    const lastCol = 시트.getLastColumn();
    if (lastCol < 1) return;
    const 헤더 = 시트.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
    const iOld = 헤더.indexOf(옛컬럼명);
    if (iOld >= 0) 시트.deleteColumn(iOld + 1);
  });
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
  const 활성심사원명 = _활성심사원목록_(ss).map(심사원 => 심사원.심사원명);
  if (iD담당심사원 > 0 && 활성심사원명.length) {
    일정시트.getRange(2, iD담당심사원, 999, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(활성심사원명, true)
        .setAllowInvalid(false)
        .setHelpText('심사원관리 시트에서 활성화된 심사원을 선택하세요.')
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

function _일정관리헤더색적용_(시트) {
  const lastCol = Math.max(1, 시트.getLastColumn());
  const 헤더 = 시트.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

  // 기존 시트도 갱신 시 전체 헤더를 남색으로 통일하고, 특이사항만 연한 톤으로 구분합니다.
  시트.getRange(1, 1, 1, lastCol)
    .setBackground(일정관리_헤더색)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  const 특이사항열 = 헤더.indexOf('특이사항') + 1;
  if (특이사항열 > 0) {
    시트.getRange(1, 특이사항열)
      .setBackground(일정관리_특이사항헤더색)
      .setFontColor(일정관리_특이사항헤더글자색)
      .setFontWeight('bold');
  }
}

function _일정관리서식적용_(시트, 요약뷰) {
  _일정관리레거시컬럼삭제_(시트);
  const lastCol = Math.max(1, 시트.getLastColumn());
  const 헤더 = 시트.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

  // 일정관리 날짜는 실제 날짜값을 유지하고 화면에는 두 자리 연도로 간결하게 표시
  ['신청일', '심사접수일', '마감예정일', '보완요청일', '연장마감일'].forEach(날짜헤더 => {
    const 열 = 헤더.indexOf(날짜헤더) + 1;
    if (열 > 0) 시트.getRange(2, 열, Math.max(1, 시트.getMaxRows() - 1), 1).setNumberFormat('yy-mm-dd');
  });

  시트.setHiddenGridlines(false);
  시트.setFrozenRows(1);
  시트.setFrozenColumns(Math.min(2, lastCol));

  시트.setRowHeight(1, 34);
  시트.getRange(1, 1, Math.max(1, 시트.getMaxRows()), lastCol)
    .setVerticalAlignment('middle');
  시트.getRange(2, 1, Math.max(1, 시트.getMaxRows() - 1), lastCol)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  try {
    _일정관리구글표적용_(시트, 헤더);
  } catch (e) {
    Logger.log('일정관리 Google Sheets 표 적용 실패: ' + e.message);
  }

  _일정관리헤더색적용_(시트);

  const 너비맵 = {
    '순번': 45, '접수번호': 115,
    '신청일': 90, '심사접수일': 95, '마감예정일': 95,
    '상태': 75, '보완요청일': 95, '연장마감일': 95, '담당심사원': 95, '특이사항': 260,
    '기업명': 155, '담당자명': 85, '연락처': 115, '이메일': 180, '소재지': 200,
    '제품명': 190, '제품수': 65, '개요': 260,
    '제공형태': 110, '제품분류': 100,
    '인공지능적용목적': 260, '인공지능적용범위': 260,
    '명세서작성방식': 110, '기타제출서류여부': 120, '보유인증': 130,
    '인공지능기능수': 90,
  };
  헤더.forEach((h, idx) => {
    if (너비맵[h]) 시트.setColumnWidth(idx + 1, 너비맵[h]);
  });

  시트.showColumns(1, lastCol);
  if (요약뷰) {
    const 표시컬럼 = new Set([
      '순번', '접수번호',
      '신청일', '심사접수일', '마감예정일',
      '상태', '보완요청일', '연장마감일', '담당심사원', '특이사항',
      '기업명', '담당자명', '연락처', '소재지',
      '제품명', '제품수', '제공형태', '제품분류',
      '인공지능기능수',
    ]);
    헤더.forEach((h, idx) => {
      if (h && !표시컬럼.has(h)) 시트.hideColumns(idx + 1);
    });
  }
}

function _일정관리구글표적용_(시트, 헤더) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const sheetId = 시트.getSheetId();
  const tableName = '일정관리_표';
  const lastCol = 헤더.length;
  const lastRow = Math.max(2, 시트.getLastRow());
  const 기존표목록 = _시트표목록조회_(spreadsheetId, sheetId);

  const requests = 기존표목록
    .filter(t => t.name === tableName || (t.range && t.range.sheetId === sheetId))
    .map(t => ({ deleteTable: { tableId: t.tableId } }));

  requests.push({
    addTable: {
      table: {
        name: tableName,
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: lastRow,
          startColumnIndex: 0,
          endColumnIndex: lastCol,
        },
        rowsProperties: {
          headerColorStyle: { rgbColor: _hexToRgb_(일정관리_헤더색) },
          firstBandColorStyle: { rgbColor: _hexToRgb_('#ffffff') },
          secondBandColorStyle: { rgbColor: _hexToRgb_('#f8fbfb') },
        },
        columnProperties: 헤더.map((h, idx) => ({
          columnIndex: idx,
          columnName: h || `Column ${idx + 1}`,
          columnType: _일정관리표컬럼타입_(h),
        })),
      },
    },
  });

  _sheetsBatchUpdate_(spreadsheetId, requests);
}

function _일정관리표컬럼타입_(헤더명) {
  if (['순번', '제품수', '인공지능기능수'].indexOf(헤더명) >= 0) return 'DOUBLE';
  if (['신청일', '심사접수일', '마감예정일', '보완요청일', '연장마감일'].indexOf(헤더명) >= 0) return 'DATE';
  return 'TEXT';
}

function _시트표목록조회_(spreadsheetId, sheetId) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties(sheetId),tables(tableId,name,range))`;
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Sheets API 표 조회 실패 (${code}): ${res.getContentText()}`);
  }
  const data = JSON.parse(res.getContentText());
  const sheet = (data.sheets || []).find(s => s.properties && s.properties.sheetId === sheetId);
  return sheet && sheet.tables ? sheet.tables : [];
}

function _sheetsBatchUpdate_(spreadsheetId, requests) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ requests: requests }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Sheets API batchUpdate 실패 (${code}): ${res.getContentText()}`);
  }
}

function _hexToRgb_(hex) {
  const value = String(hex || '').replace('#', '');
  const n = parseInt(value, 16);
  return {
    red: ((n >> 16) & 255) / 255,
    green: ((n >> 8) & 255) / 255,
    blue: (n & 255) / 255,
  };
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

/** 심사원관리 시트 기본값·표시 형식 설정 */
function _심사원관리설정_(ss) {
  const 시트 = ss.getSheetByName(SHEET.심사원관리);
  if (!시트) return;

  if (시트.getLastRow() < 2) {
    const 기본심사원 = (CONFIG.담당자목록 || []).map((이름, i) => [이름, '', 'Y', i + 1, '']);
    if (기본심사원.length) 시트.getRange(2, 1, 기본심사원.length, 5).setValues(기본심사원);
  }

  시트.setFrozenRows(1);
  시트.setColumnWidth(1, 110);
  시트.setColumnWidth(2, 210);
  시트.setColumnWidth(3, 85);
  시트.setColumnWidth(4, 85);
  시트.setColumnWidth(5, 170);
  시트.getRange(2, 3, Math.max(1, 시트.getMaxRows() - 1), 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Y', 'N'], true).setAllowInvalid(false).build());

  const 로그시트 = ss.getSheetByName(SHEET.배정알림로그);
  if (로그시트) {
    로그시트.setFrozenRows(1);
    [145, 120, 110, 210, 100, 210].forEach((너비, i) => 로그시트.setColumnWidth(i + 1, 너비));
  }
}

/** 심사원관리 시트의 활성 심사원을 배정순서대로 반환 */
function _활성심사원목록_(ss) {
  const 시트 = ss.getSheetByName(SHEET.심사원관리);
  if (!시트 || 시트.getLastRow() < 2) {
    return (CONFIG.담당자목록 || []).filter(Boolean).map((이름, i) => ({
      심사원명: String(이름).trim(), 이메일: '', 활성여부: 'Y', 배정순서: i + 1, Chat사용자ID: '',
    }));
  }

  const 값 = 시트.getRange(2, 1, 시트.getLastRow() - 1, 5).getValues();
  return 값.map((행, i) => ({
    심사원명: String(행[0] || '').trim(),
    이메일: String(행[1] || '').trim(),
    활성여부: String(행[2] || '').trim().toUpperCase(),
    배정순서: Number(행[3]) || (i + 1),
    Chat사용자ID: String(행[4] || '').trim(),
  }))
    .filter(심사원 => 심사원.심사원명 && 심사원.활성여부 === 'Y')
    .sort((a, b) => a.배정순서 - b.배정순서 || a.심사원명.localeCompare(b.심사원명));
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
    ['제품명', '제품 또는 서비스 모델명, 제품 또는 서비스명, 제품·서비스 모델명, 제품명, 서비스명'],
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
    ['레퍼런스참조위치', '레퍼런스 참조위치, 설명서참조위치, 설명서 참조 위치'],
    ['구현방식', 'AI 구현 방식, 인공지능 구현 방식, 구현방식, 구현 방식'],
    ['연산자원요약', 'AI 연산 자원 요약, 연산자원요약, 연산자원, 연산 자원'],
    ['실행환경요약', 'AI 실행 환경 요약, 실행환경요약, 실행환경, 실행 환경'],
    ['학습데이터사양', '학습 데이터 사양, 학습데이터사양, 데이터명 라이선싱'],
    ['개발환경라이브러리알고리즘', '개발환경·라이브러리·알고리즘, 개발환경라이브러리알고리즘'],
    ['BaseModel명칭', 'Base Model 명칭, BaseModel명칭, BaseModel, Base Model'],
    ['튜닝방법', '튜닝 방법, 튜닝방법'],
    ['튜닝데이터셋', '튜닝 데이터셋, 튜닝데이터셋'],
    ['외부API정보', '외부 API 정보, 외부API정보, 외부 API, 외부API모델'],
    ['타겟HW_OS', '타겟 하드웨어·OS, 타겟HW_OS, 타겟, HW, OS'],
    ['추론런타임', '추론 런타임, 추론런타임, 배포포맷정밀도'],
    ['혼합구성설명', '혼합 구성 설명, 혼합구성설명'],
    ['모델별역할및입출력흐름', '모델별 역할 및 입출력 흐름, 모델별역할및입출력흐름, 세부구성요소별설명'],
    ['모델명', '제품 또는 서비스 모델명, 제품·서비스 모델명, 모델명'],
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
  let 등록건수 = 0;
  let 이미존재건수 = 0;
  const 건너뛴행 = [];

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
    건목록.forEach((건, idx) => {
      try {
        const 결과 = _Sheets에등록(건, 파일명);
        const 접수번호 = 결과.접수번호;
        if (건.제품명) 제품명별접수번호[String(건.제품명).trim()] = 접수번호;
        // 접수번호 자기참조도 등록 (기능탭에 접수번호 컬럼이 있으면 직접 매칭)
        제품명별접수번호['__접수__' + 접수번호] = 접수번호;
        if (결과.신규) {
          try { _보관폴더준비(접수번호); } catch (e2) { Logger.log('접수번호 폴더 생성 실패: ' + e2.message); }
          등록건수++;
        } else {
          이미존재건수++;  // append-only 정책 — 기존 데이터는 덮어쓰지 않고 건너뜀
        }
      } catch (e) {
        // 접수번호 없는 행 등은 등록하지 않고 건너뜀 (로그에 기록)
        건너뛴행.push(`${idx + 1}행: ${e.message}${건.제품명 ? ' (제품: ' + 건.제품명 + ')' : ''}`);
      }
    });

    if (건너뛴행.length) {
      try {
        ss.getSheetByName(SHEET.로그).appendRow([
          Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
          파일명, '', '건너뜀',
          `접수번호 누락 등으로 ${건너뛴행.length}건 미등록 — ${건너뛴행.join(' / ')}`,
        ]);
      } catch (e2) {}
    }

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

  // 신규행뿐 아니라 구버전 수식이 남은 기존 행도 WD 15일 + 공휴일 기준으로 갱신
  _마감예정일수식갱신_(ss);

  // 파싱 결과 알림
  let msg = `엑셀 파싱 완료\n\n신규 등록: ${등록건수}건`;
  if (이미존재건수) {
    msg += `\n이미 등록됨(건너뜀): ${이미존재건수}건 — append-only 정책으로 기존 데이터 유지`;
  }
  if (건너뛴행.length) {
    msg += `\n건너뜀: ${건너뛴행.length}건 (접수번호 누락 등)\n\n` + 건너뛴행.join('\n');
    msg += `\n\n※ 건너뛴 행은 KOSA 접수번호가 없어 등록되지 않았습니다. 파싱로그를 확인하세요.`;
  }
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
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
    제품명: idx('제품명', ['제품 또는 서비스 모델명', '제품명', '제품 또는 서비스명', '서비스명']),
    기능번호: idx('기능번호', ['기능번호', '번호']),
    기능명: idx('기능명', ['기능명']),
    역할: idx('인공지능역할', ['인공지능역할', '인공지능 역할', '역할', 'AI 역할']),
    입력: idx('입력', ['입력']),
    출력: idx('출력', ['출력']),
    레퍼런스참조위치: idx('레퍼런스참조위치', ['레퍼런스참조위치', '레퍼런스 참조위치', '설명서참조위치', '설명서 참조 위치']),
    구현방식: idx('구현방식', ['AI 구현 방식', '구현방식', '구현 방식']),
    연산자원요약: idx('연산자원요약', ['AI 연산 자원 요약', '연산자원요약', '연산자원', '연산 자원']),
    실행환경요약: idx('실행환경요약', ['AI 실행 환경 요약', '실행환경요약', '실행환경', '실행 환경']),
    학습데이터사양: idx('학습데이터사양', ['학습데이터사양', '학습 데이터 사양', '학습데이터']),
    개발환경라이브러리알고리즘: idx('개발환경라이브러리알고리즘', ['개발환경·라이브러리·알고리즘', '개발환경라이브러리알고리즘']),
    BaseModel명칭: idx('BaseModel명칭', ['Base Model 명칭', 'BaseModel명칭', 'BaseModel', 'Base Model', '베이스']),
    튜닝방법: idx('튜닝방법', ['튜닝 방법', '튜닝방법']),
    튜닝데이터셋: idx('튜닝데이터셋', ['튜닝 데이터셋', '튜닝데이터셋']),
    외부API정보: idx('외부API정보', ['외부 API 정보', '외부API정보', '외부 API', '외부API']),
    타겟HW_OS: idx('타겟HW_OS', ['타겟 하드웨어·OS', '타겟HW_OS', '타겟', 'HW', 'OS']),
    추론런타임: idx('추론런타임', ['추론 런타임', '추론런타임', '배포포맷정밀도']),
    혼합구성설명: idx('혼합구성설명', ['혼합 구성 설명', '혼합구성설명', '혼합']),
    모델별역할및입출력흐름: idx('모델별역할및입출력흐름', ['모델별 역할 및 입출력 흐름', '모델별역할및입출력흐름']),
    입력데이터설명: idx('입력데이터설명', ['입력데이터설명']),
    출력데이터설명: idx('출력데이터설명', ['출력데이터설명']),
    기타참고자료: idx('기타참고자료파일명', ['기타참고자료', '기타 참고자료']),
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
      레퍼런스참조위치: get(I.레퍼런스참조위치),
      구현방식: get(I.구현방식),
      연산자원요약: get(I.연산자원요약),
      실행환경요약: get(I.실행환경요약),
      학습데이터사양: get(I.학습데이터사양),
      개발환경라이브러리알고리즘: get(I.개발환경라이브러리알고리즘),
      BaseModel명칭: get(I.BaseModel명칭),
      튜닝방법: get(I.튜닝방법),
      튜닝데이터셋: get(I.튜닝데이터셋),
      외부API정보: get(I.외부API정보),
      타겟HW_OS: get(I.타겟HW_OS),
      추론런타임: get(I.추론런타임),
      혼합구성설명: get(I.혼합구성설명),
      모델별역할및입출력흐름: get(I.모델별역할및입출력흐름),
      입력데이터설명: get(I.입력데이터설명),
      출력데이터설명: get(I.출력데이터설명),
      기타참고자료파일명: get(I.기타참고자료),
    };
    (묶음[접수번호] = 묶음[접수번호] || []).push(기능);
  }

  // 접수번호별로 기능상세 등록 + 접수대장 기능수 갱신 (이미 등록된 접수번호는 건너뜀)
  Object.keys(묶음).forEach(접수번호 => {
    const 등록됨 = AI기능상세등록(접수번호, 묶음[접수번호]);
    if (등록됨) _접수대장기능수갱신(접수번호, 묶음[접수번호]);
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
    제품명: idx('제품명', ['제품 또는 서비스 모델명', '제품명', '제품 또는 서비스명', '서비스명']),
    연번: idx('연번', ['연번']),
    모델명: idx('모델명', ['제품 또는 서비스 모델명', '모델명', '제품·서비스 모델명']),
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
 *  · 마감예정일 → 신청일로부터 15 WD(주말·공휴일 제외) 수식
 *
 * VLOOKUP은 접수번호(A열)를 키로 하므로, 접수대장 행이 정렬·이동돼도
 * 항상 올바른 값을 따라갑니다. 접수대장에서 회사·제품정보를 고치면
 * 일정관리에 자동 반영됩니다.
 */
function _일정관리행추가(ss, 접수번호, 직접값) {
  const 일정시트 = ss.getSheetByName(SHEET.일정관리);
  const 신청연도 = _날짜연도_(직접값.신청일) || new Date().getFullYear();
  _공휴일연도확보_(ss, [신청연도, 신청연도 + 1]);
  // ⚠️ 실제 시트의 라이브 헤더를 읽음 (정적 정의가 아니라).
  // 심사원이 컬럼을 수동 추가·삽입해도 값이 밀리지 않고 이름 기준으로 배치됨.
  const 일정H = 일정시트.getRange(1, 1, 1, 일정시트.getLastColumn())
    .getValues()[0].map(v => String(v).trim());
  const 대장H = 시트헤더정의[SHEET.접수대장];

  // 접수대장을 원본으로 참조할 컬럼 (일정관리 컬럼명 → 접수대장 컬럼명)
  // 대부분 이름이 같지만, 명세서/기타제출서류는 가공이 필요해 별도 처리
  const 참조맵 = {
    '기업명': '기업명',
    '담당자명': '담당자명',
    '연락처': '연락처',
    '이메일': '이메일',
    '소재지': '소재지',
    '제품명': '제품명',
    '제품수': '제품수',
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

    // 1) 직접 입력값 (상태·담당심사원·신청일·심사접수일)
    if (h === '접수번호') return 접수번호;
    if (h === '상태') return '대기';
    if (Object.prototype.hasOwnProperty.call(직접값, h)) return 직접값[h];

    // 2) 마감예정일 = 신청일로부터 15 WD (토·일·공휴일 제외)
    if (h === '마감예정일') {
      const iD신청 = 일정H.indexOf('신청일') + 1;
      const 신청셀 = `${columnLetter(iD신청)}${새행번호}`;
      return `=IF(${신청셀}="","",WORKDAY(${신청셀},15,'${공휴일시트명}'!$A$2:$A))`;
    }

    // 2-1) 연장마감일 = 보완요청일로부터 30 WD (토·일·공휴일 제외)
    if (h === '연장마감일') {
      const iD보완 = 일정H.indexOf('보완요청일') + 1;
      const 보완셀 = `${columnLetter(iD보완)}${새행번호}`;
      return `=IF(${보완셀}="","",WORKDAY(${보완셀},30,'${공휴일시트명}'!$A$2:$A))`;
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
  try {
    _일정관리서식적용_(일정시트, true);
  } catch (e) {
    Logger.log('일정관리 표 갱신 실패: ' + e.message);
  }
}

/** 날짜 값(Date 또는 문자열)에서 연도를 추출 */
function _날짜연도_(값) {
  if (값 instanceof Date && !isNaN(값.getTime())) return 값.getFullYear();
  const m = String(값 || '').match(/(20\d{2})/);
  return m ? Number(m[1]) : null;
}

/** 대한민국 공휴일을 자동 보관하는 보조 시트 확보 */
function _공휴일시트확보_(ss) {
  let 시트 = ss.getSheetByName(공휴일시트명);
  if (!시트) {
    시트 = ss.insertSheet(공휴일시트명);
    시트.getRange(1, 1, 1, 2).setValues([['날짜', '공휴일명']])
      .setBackground('#5f6368').setFontColor('#ffffff').setFontWeight('bold');
    시트.setFrozenRows(1);
    시트.setColumnWidth(1, 100);
    시트.setColumnWidth(2, 180);
  }
  return 시트;
}

/**
 * 공개 대한민국 공휴일 캘린더에서 필요한 연도의 법정·대체·임시 공휴일을 자동 동기화한다.
 * 기념일(어버이날·식목일 등)은 WORKDAY 제외일이 아니므로 포함하지 않는다.
 */
function _공휴일연도확보_(ss, 연도목록) {
  const 시트 = _공휴일시트확보_(ss);
  const 필요연도 = Array.from(new Set(연도목록.map(Number).filter(y => y >= 2000 && y <= 2100)));
  if (!필요연도.length) return 시트;

  const 기존값 = 시트.getLastRow() > 1
    ? 시트.getRange(2, 1, 시트.getLastRow() - 1, 2).getValues()
    : [];
  const 기존연도 = new Set(기존값.map(r => _날짜연도_(r[0])).filter(Boolean));
  const 누락연도 = 필요연도.filter(y => !기존연도.has(y));
  if (!누락연도.length) return 시트;

  try {
    const ics = UrlFetchApp.fetch(대한민국공휴일캘린더URL, { muteHttpExceptions: false })
      .getContentText('UTF-8').replace(/\r?\n[ \t]/g, '');
    const 공휴일명패턴 = /(새해|신정|설날|삼일절|3·1절|어린이날|부처님오신날|석가탄신일|현충일|광복절|추석|개천절|한글날|성탄절|크리스마스|선거일|임시공휴일|대체공휴일|쉬는 날)/;
    const 추가값 = [];

    ics.split('BEGIN:VEVENT').slice(1).forEach(block => {
      const 날짜매치 = block.match(/DTSTART;VALUE=DATE:(\d{4})(\d{2})(\d{2})/);
      const 이름매치 = block.match(/(?:^|\n)SUMMARY(?:;[^:]*)?:(.*)/);
      if (!날짜매치 || !이름매치) return;
      const 연도 = Number(날짜매치[1]);
      const 이름 = 이름매치[1].trim().replace(/\\,/g, ',');
      if (누락연도.indexOf(연도) < 0 || !공휴일명패턴.test(이름)) return;
      추가값.push([new Date(연도, Number(날짜매치[2]) - 1, Number(날짜매치[3])), 이름]);
    });

    const 날짜맵 = {};
    기존값.concat(추가값).forEach(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (isNaN(d.getTime())) return;
      const key = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
      날짜맵[key] = [d, r[1]];
    });
    const 전체값 = Object.keys(날짜맵).sort().map(k => 날짜맵[k]);
    if (시트.getLastRow() > 1) 시트.getRange(2, 1, 시트.getLastRow() - 1, 2).clearContent();
    if (전체값.length) {
      시트.getRange(2, 1, 전체값.length, 2).setValues(전체값);
      시트.getRange(2, 1, 전체값.length, 1).setNumberFormat('yy-mm-dd');
    }
  } catch (e) {
    Logger.log('공휴일 자동 동기화 실패(기존 공휴일 목록으로 계산): ' + e.message);
  }
  return 시트;
}

/** 일정관리의 기본 마감(15 WD)과 보완 연장마감(30 WD) 수식을 모두 갱신 */
function _마감예정일수식갱신_(ss) {
  const 시트 = ss.getSheetByName(SHEET.일정관리);
  if (!시트 || 시트.getLastRow() < 2) {
    _공휴일연도확보_(ss, [new Date().getFullYear(), new Date().getFullYear() + 1]);
    return 0;
  }
  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 신청열 = 헤더.indexOf('신청일') + 1;
  const 마감열 = 헤더.indexOf('마감예정일') + 1;
  const 보완요청열 = 헤더.indexOf('보완요청일') + 1;
  const 연장마감열 = 헤더.indexOf('연장마감일') + 1;
  if (신청열 < 1 || 마감열 < 1) return 0;

  const 행수 = 시트.getLastRow() - 1;
  const 신청값 = 시트.getRange(2, 신청열, 행수, 1).getValues().flat();
  const 보완요청값 = 보완요청열 > 0
    ? 시트.getRange(2, 보완요청열, 행수, 1).getValues().flat()
    : [];
  // 기본 마감뿐 아니라 연장마감 계산에 필요한 연도의 공휴일도 확보한다.
  // 보완요청일이 신청일의 다음 해 이후인 경우 신청일만 보면 해당 공휴일이 누락될 수 있다.
  const 연도목록 = 신청값.concat(보완요청값).map(_날짜연도_).filter(Boolean);
  const 현재연도 = new Date().getFullYear();
  연도목록.push(현재연도, 현재연도 + 1);
  _공휴일연도확보_(ss, 연도목록.concat(연도목록.map(y => y + 1)));

  const 신청열문자 = columnLetter(신청열);
  const 수식 = 신청값.map((_, i) => {
    const 행 = i + 2;
    const 신청셀 = `${신청열문자}${행}`;
    return [`=IF(${신청셀}="","",WORKDAY(${신청셀},15,'${공휴일시트명}'!$A$2:$A))`];
  });
  시트.getRange(2, 마감열, 행수, 1).setFormulas(수식).setNumberFormat('yy-mm-dd');
  if (보완요청열 > 0 && 연장마감열 > 0) {
    const 보완요청열문자 = columnLetter(보완요청열);
    const 연장수식 = 신청값.map((_, i) => {
      const 행 = i + 2;
      const 보완셀 = `${보완요청열문자}${행}`;
      return [`=IF(${보완셀}="","",WORKDAY(${보완셀},30,'${공휴일시트명}'!$A$2:$A))`];
    });
    // 구버전에서 남아 있을 수 있는 수기 날짜 입력 검사를 먼저 제거해야
    // 자동 수식 입력 시 "날짜를 직접 선택" 유효성 검사 예외가 발생하지 않는다.
    시트.getRange(2, 연장마감열, 행수, 1)
      .clearDataValidations()
      .setFormulas(연장수식)
      .setNumberFormat('yy-mm-dd');
  }
  return 행수;
}

/** 관리자 수동 실행용: 공휴일과 기존 마감예정일 수식을 즉시 갱신 */
function 마감예정일갱신() {
  const 갱신건수 = _마감예정일수식갱신_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert(`마감일 갱신 완료: ${갱신건수}건\n기본: 신청일 + 15 WD\n보완: 보완요청일 + 30 WD\n(주말·공휴일 제외)`);
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

  // 이미 등록된 접수번호는 건너뜀 (append-only 정책 — 기존 제품모델 행을 지우고 다시 쓰지 않음)
  const D = 시트.getDataRange().getValues();
  if (D.length > 1) {
    const iNo = D[0].indexOf('접수번호');
    const 이미존재 = D.slice(1).some(행 => String(행[iNo]).trim() === String(접수번호).trim());
    if (이미존재) {
      Logger.log(`제품모델 이미 등록됨 - 건너뜀: ${접수번호}`);
      return false;
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

  _접수대장제품모델요약갱신(ss, 접수번호, 모델목록);
  return true;
}

/**
 * 제품모델 반복행 기준으로 접수대장의 제품수를 갱신합니다.
 * 제품명은 더 이상 여기서 덮어쓰지 않습니다 — 신청 엑셀에는 원래부터 '제품명'이라는
 * 별도 항목이 없고 '제품 또는 서비스 모델명'(조달청 기준 모델명)이 유일한 기준값이라,
 * 기본정보 탭 파싱 시 채운 값이 이미 최종값입니다.
 */
function _접수대장제품모델요약갱신(ss, 접수번호, 모델목록) {
  const 대장시트 = ss.getSheetByName(SHEET.접수대장);
  if (!대장시트 || !모델목록 || !모델목록.length) return;

  const D = 대장시트.getDataRange().getValues();
  const H = D[0].map(v => String(v).trim());
  const iNo = H.indexOf('접수번호');
  const i제품명 = H.indexOf('제품명');
  const i제품수 = H.indexOf('제품수');
  if (iNo < 0) return;

  // 신청 폼에는 "대표 제품명"이라는 별도 항목이 없다 — 조달청 기준 모델명(제품모델
  // 반복 목록)만 존재하므로, 접수대장 제품명은 항상 1번(연번 최소) 모델명을 그대로 쓴다.
  const 정렬목록 = 모델목록.slice().sort((a, b) => Number(a.연번 || 0) - Number(b.연번 || 0));
  const 대표모델명 = String(정렬목록[0].모델명 || '').trim();

  for (let r = 1; r < D.length; r++) {
    if (String(D[r][iNo]).trim() !== String(접수번호).trim()) continue;
    if (i제품명 >= 0 && 대표모델명) 대장시트.getRange(r + 1, i제품명 + 1).setValue(대표모델명);
    if (i제품수 >= 0) 대장시트.getRange(r + 1, i제품수 + 1).setValue(모델목록.length);
    break;
  }
}

/** 기존 제품모델 시트 전체를 기준으로 접수대장 제품명/제품수를 재동기화 */
function _전체제품모델요약갱신(ss) {
  const 모델시트 = ss.getSheetByName(SHEET.제품모델);
  if (!모델시트 || 모델시트.getLastRow() < 2) return;

  const D = 모델시트.getDataRange().getValues();
  const H = D[0].map(v => String(v).trim());
  const iNo = H.indexOf('접수번호');
  if (iNo < 0) return;

  const 묶음 = {};
  D.slice(1).forEach(행 => {
    const 접수번호 = String(행[iNo] || '').trim();
    if (!접수번호) return;
    const obj = {};
    H.forEach((h, i) => { obj[h] = 행[i]; });
    if (!묶음[접수번호]) 묶음[접수번호] = [];
    묶음[접수번호].push({
      연번: obj['연번'],
      모델명: obj['모델명'],
      세부품명번호: obj['세부품명번호'],
      물품식별번호: obj['물품식별번호'],
    });
  });

  Object.keys(묶음).forEach(접수번호 => {
    _접수대장제품모델요약갱신(ss, 접수번호, 묶음[접수번호]);
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
    제품명:         V('제품명', ['제품 또는 서비스 모델명', '제품 또는 서비스명', '제품·서비스명', '제품명', '서비스명']),
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

  // 접수번호는 KOSA가 부여한 신청번호(예: AI202600001)를 그대로 키로 사용.
  // TTA가 임의 생성하면 KOSA 번호 체계와 어긋나므로, 없으면 오류 처리(폴백 없음).
  const 접수번호 = String(건.접수번호 || '').trim();
  if (!접수번호) {
    throw new Error('접수번호(신청번호)가 없습니다. KOSA가 부여한 접수번호가 있는 행만 등록할 수 있습니다.');
  }

  const 오늘 = new Date();
  const 마감일 = new Date(오늘);
  마감일.setDate(마감일.getDate() + CONFIG.기본심사기간);

  const 대장시트 = ss.getSheetByName(SHEET.접수대장);
  const D = 대장시트.getDataRange().getValues();
  const H = D[0];
  const iNo = H.indexOf('접수번호');

  // 이미 등록된 접수번호는 절대 덮어쓰지 않고 무조건 건너뜀 (append-only 정책).
  // 재파싱·중복 파일 업로드로 같은 접수번호가 다시 들어와도 기존 데이터는 그대로 유지됩니다.
  const 이미존재 = D.slice(1).some(행 => String(행[iNo]).trim() === 접수번호);
  if (이미존재) {
    try {
      ss.getSheetByName(SHEET.로그).appendRow([
        Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
        파일명, 접수번호, '건너뜀',
        `이미 등록된 접수번호 — append-only 정책으로 덮어쓰지 않음 (${건.제품명 || ''})`,
      ]);
    } catch (e2) {}
    Logger.log(`이미 등록됨 - 건너뜀: ${접수번호}`);
    return { 접수번호, 신규: false };
  }

  const 담당자 = _담당자배분(ss);
  const 착수일 = Utilities.formatDate(오늘, 'Asia/Seoul', 'yyyy-MM-dd');
  const 상태 = '접수';

  // ── 헤더 이름 기반 쓰기 ──
  // 컬럼 순서가 바뀌거나 시트에 컬럼이 추가/삽입돼도 값이 밀리지 않습니다.
  const 값맵 = {
    // TTA 관리
    '접수번호': 접수번호,
    '심사접수일': Utilities.formatDate(오늘, 'Asia/Seoul', 'yyyy-MM-dd'),
    '상태': 상태,
    // 기업 정보
    '기업명': 건.기업명, '사업자번호': 건.사업자번호, '대표자': 건.대표자, '소재지': 건.소재지,
    '담당자명': 건.담당자명, '연락처': 건.연락처, '이메일': 건.이메일,
    // 제품·서비스 정보
    '제품명': 건.제품명, '제품수': 건.제품수 || 1, '제공형태': 건.제공형태, '제품분류': 건.제품분류,
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

  const 행값 = H.map(h => 값맵[h] ?? '');
  대장시트.appendRow(행값);

  // ── 일정관리 추가 ──
  // 하이브리드 싱크:
  //   · 신청정보(회사·연락처·제품 등) = 접수대장 VLOOKUP 수식 참조 (접수대장이 원본)
  //   · 상태·담당심사원 = 일정관리에서 직접 편집 (일정관리가 원본)
  //   · 마감예정일 = 신청일 + 15 WD (주말·공휴일 제외) 수식
  _일정관리행추가(ss, 접수번호, {
    신청일: 건.신청일 || '',
    심사접수일: Utilities.formatDate(오늘, 'Asia/Seoul', 'yyyy-MM-dd'),
    담당심사원: 담당자,   // 최초 배분값 (이후 일정관리에서 직접 수정 가능)
  });

  // ── 로그 ──
  ss.getSheetByName(SHEET.로그).appendRow([
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    파일명, 접수번호, '신규등록',
    `${건.기업명} / ${건.제품명} 등록 → 담당: ${담당자}`,
  ]);

  Logger.log(`등록: ${접수번호} / ${건.제품명}`);
  return { 접수번호, 신규: true };
}


function _담당자배분(ss) {
  const 심사원목록 = _활성심사원목록_(ss);
  if (!심사원목록.length) throw new Error('심사원관리 시트에 활성 심사원이 없습니다.');
  const 시트 = ss.getSheetByName(SHEET.접수대장);
  const 마지막행 = 시트.getLastRow() - 1; // 헤더 제외
  const idx = 마지막행 % 심사원목록.length;
  return 심사원목록[idx].심사원명;
}

// ─────────────────────────────────────────────
// 심사원 배정 확정 및 Google Chat 알림
// ─────────────────────────────────────────────

/** 일정관리에서 선택한 데이터 행을 헤더 기반 객체로 반환 */
function _선택일정행목록_(ss) {
  const 시트 = ss.getActiveSheet();
  if (시트.getName() !== SHEET.일정관리) throw new Error('일정관리 시트에서 데이터 행을 선택하세요.');
  const 선택 = 시트.getActiveRange();
  const 시작행 = Math.max(2, 선택.getRow());
  const 끝행 = 선택.getLastRow();
  if (끝행 < 2) throw new Error('헤더가 아닌 데이터 행을 선택하세요.');
  if (끝행 - 시작행 + 1 > 50) throw new Error('한 번에 최대 50개 행까지 발송할 수 있습니다.');

  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 값 = 시트.getRange(시작행, 1, 끝행 - 시작행 + 1, 헤더.length).getDisplayValues();
  return 값.map((행값, i) => {
    const 건 = { _행번호: 시작행 + i };
    헤더.forEach((h, j) => { 건[h] = 행값[j]; });
    return 건;
  }).filter(건 => String(건['접수번호'] || '').trim());
}

/** 선택 행 중 담당심사원이 비어 있는 행에만 라운드로빈 초안을 입력 */
function 선택행라운드로빈추천() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시트 = ss.getActiveSheet();
  const 건목록 = _선택일정행목록_(ss);
  const 심사원목록 = _활성심사원목록_(ss);
  if (!심사원목록.length) throw new Error('심사원관리 시트에 활성 심사원이 없습니다.');
  const 헤더 = 시트.getRange(1, 1, 1, 시트.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const 담당열 = 헤더.indexOf('담당심사원') + 1;
  let 입력수 = 0;
  const 시작순서 = Math.max(0, ss.getSheetByName(SHEET.접수대장).getLastRow() - 1);
  건목록.forEach(건 => {
    if (String(건['담당심사원'] || '').trim()) return;
    const 심사원 = 심사원목록[(시작순서 + 입력수) % 심사원목록.length];
    시트.getRange(건._행번호, 담당열).setValue(심사원.심사원명);
    입력수++;
  });
  SpreadsheetApp.getUi().alert(`라운드로빈 추천 완료: ${입력수}건\n기존 담당심사원 값은 변경하지 않았습니다.`);
}

function _성공발송키목록_(ss) {
  const 시트 = ss.getSheetByName(SHEET.배정알림로그);
  if (!시트 || 시트.getLastRow() < 2) return new Set();
  return new Set(
    시트.getRange(2, 1, 시트.getLastRow() - 1, 6).getValues()
      .filter(행 => String(행[4]).trim() === '발송완료')
      .map(행 => `${String(행[1]).trim()}|${String(행[2]).trim()}`)
  );
}

function _Chat멘션_(사용자ID, 심사원명) {
  const id = String(사용자ID || '').trim();
  if (!id) return `*${심사원명}*`;
  return `<${id.indexOf('users/') === 0 ? id : 'users/' + id}>`;
}

/** 선택 행을 담당자별로 묶어 공용 Chat에 한 번 발송 */
function _선택행배정알림발송_(재발송) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const 건목록 = _선택일정행목록_(ss);
  if (!건목록.length) throw new Error('발송할 접수 건이 없습니다.');
  if (!CONFIG.챗_공통Webhook) throw new Error('Script Properties의 CHAT_COMMON_WEBHOOK을 먼저 설정하세요.');

  const 심사원맵 = {};
  _활성심사원목록_(ss).forEach(심사원 => { 심사원맵[심사원.심사원명] = 심사원; });
  const 누락 = 건목록.filter(건 => !심사원맵[String(건['담당심사원'] || '').trim()]);
  if (누락.length) {
    throw new Error('활성 심사원이 지정되지 않은 행이 있습니다: ' + 누락.map(건 => 건['접수번호']).join(', '));
  }
  const 이메일누락 = Array.from(new Set(건목록.map(건 => String(건['담당심사원']).trim())))
    .filter(이름 => !심사원맵[이름].이메일);
  if (이메일누락.length) throw new Error('심사원관리 시트에 이메일을 입력하세요: ' + 이메일누락.join(', '));

  if (!재발송) {
    const 성공키 = _성공발송키목록_(ss);
    const 중복 = 건목록.filter(건 => 성공키.has(`${건['접수번호']}|${건['담당심사원']}`));
    if (중복.length) {
      throw new Error('이미 발송된 건이 있습니다: ' + 중복.map(건 => 건['접수번호']).join(', ') + '\n재발송 메뉴를 사용하세요.');
    }
  }

  const 그룹 = {};
  건목록.forEach(건 => {
    const 이름 = String(건['담당심사원']).trim();
    if (!그룹[이름]) 그룹[이름] = [];
    그룹[이름].push(건);
  });
  const 요약 = Object.keys(그룹).map(이름 => `${이름} ${그룹[이름].length}건`).join(', ');
  if (ui.alert('배정 알림 발송 확인', `${요약}\n\n선택한 ${건목록.length}건을 공용 Google Chat 방에 발송할까요?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const 본문 = ['📨 *AI 기술심사 배정 확정*'];
  Object.keys(그룹).forEach(이름 => {
    const 심사원 = 심사원맵[이름];
    본문.push('', `${_Chat멘션_(심사원.Chat사용자ID, 이름)} — ${그룹[이름].length}건`);
    그룹[이름].forEach(건 => {
      본문.push(`• *${건['접수번호']}* | ${건['기업명'] || '-'} | ${건['제품명'] || '-'} | 마감 ${건['마감예정일'] || '-'}`);
    });
  });
  본문.push('', `<${ss.getUrl()}|일정관리 바로가기>`);

  const 결과 = _챗POST(CONFIG.챗_공통Webhook, { text: 본문.join('\n') });
  const 실행자 = (() => { try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; } })();
  const 발송일시 = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const 로그값 = 건목록.map(건 => {
    const 이름 = String(건['담당심사원']).trim();
    const 심사원 = 심사원맵[이름];
    return [발송일시, 건['접수번호'], 이름, 심사원.이메일, 결과.success ? '발송완료' : '발송실패', 실행자];
  });
  const 로그시트 = ss.getSheetByName(SHEET.배정알림로그);
  로그시트.getRange(로그시트.getLastRow() + 1, 1, 로그값.length, 6).setValues(로그값);
  if (!결과.success) throw new Error(`Google Chat 발송 실패 (${결과.code}): ${결과.message}`);
  ui.alert(`배정 알림 발송 완료: ${건목록.length}건`);
}

function 선택행배정확정알림발송() { _선택행배정알림발송_(false); }
function 선택행배정알림재발송() { _선택행배정알림발송_(true); }

/** Webhook POST 공통 함수 */
function _챗POST(url, payload) {
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    return { success: code >= 200 && code < 300, code: code, message: res.getContentText() };
  } catch (e) {
    Logger.log('Chat 알림 실패: ' + e.message);
    return { success: false, code: 0, message: e.message };
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

  const 결과 = _챗POST(url, {
    text: '✅ *AI 기술심사 관리 시스템* 연결 테스트\n알림이 정상적으로 도착했습니다! 🎉',
  });
  ui.alert(결과.success
    ? '테스트 메시지를 발송했습니다. Google Chat을 확인하세요.'
    : `테스트 발송 실패 (${결과.code}): ${결과.message}`);
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

  // 이미 등록된 접수번호는 건너뜀 (append-only 정책 — 기존 기능상세 행을 지우고 다시 쓰지 않음)
  const D = 시트.getDataRange().getValues();
  if (D.length > 1) {
    const iNo = D[0].indexOf('접수번호');
    const 이미존재 = D.slice(1).some(행 => String(행[iNo]).trim() === String(접수번호).trim());
    if (이미존재) {
      Logger.log(`AI기능상세 이미 등록됨 - 건너뜀: ${접수번호}`);
      return false;
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
      '레퍼런스참조위치': f.레퍼런스참조위치 ?? '',
      '구현방식': f.구현방식 ?? '',
      '연산자원요약': f.연산자원요약 ?? '',
      '실행환경요약': f.실행환경요약 ?? '',
      '학습데이터사양': f.학습데이터사양 ?? '',
      '개발환경라이브러리알고리즘': f.개발환경라이브러리알고리즘 ?? '',
      'BaseModel명칭': f.BaseModel명칭 ?? '',
      '튜닝방법': f.튜닝방법 ?? '',
      '튜닝데이터셋': f.튜닝데이터셋 ?? '',
      '외부API정보': f.외부API정보 ?? '',
      '타겟HW_OS': f.타겟HW_OS ?? '',
      '추론런타임': f.추론런타임 ?? '',
      '혼합구성설명': f.혼합구성설명 ?? '',
      '모델별역할및입출력흐름': f.모델별역할및입출력흐름 ?? '',
      '입력데이터설명': f.입력데이터설명 ?? '',
      '출력데이터설명': f.출력데이터설명 ?? '',
      '기타참고자료파일명': f.기타참고자료파일명 ?? '',
    };
    시트.appendRow(헤더행.map(h => 값맵[h] ?? ''));
  });
  return true;
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

/**
 * 접수번호로 접수대장에서 건 데이터(원본)를 조회해 객체로 반환.
 * 어느 탭에서 실행하든 항상 접수대장의 정본 데이터를 사용하게 한다.
 */
function _건조회(ss, 접수번호) {
  const 시트 = ss.getSheetByName(SHEET.접수대장);
  const D = 시트.getDataRange().getValues();
  const H = D[0];
  const iNo = H.indexOf('접수번호');
  for (let r = 1; r < D.length; r++) {
    if (String(D[r][iNo]).trim() === String(접수번호).trim()) {
      const 건 = {};
      H.forEach((h, i) => { 건[h] = D[r][i]; });
      return 건;
    }
  }
  return null;
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

function _제품모델목록조회(ss, 접수번호) {
  const 시트 = ss.getSheetByName(SHEET.제품모델);
  if (!시트 || 시트.getLastRow() < 2) return [];
  const 모든행 = 시트.getDataRange().getValues();
  const 헤더 = 모든행[0];
  const iNo = 헤더.indexOf('접수번호');
  return 모든행.slice(1)
    .filter(행 => String(행[iNo]).trim() === String(접수번호).trim())
    .map(행 => {
      const obj = {};
      헤더.forEach((h, i) => { obj[h] = 행[i]; });
      return obj;
    });
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

  const { 접수번호, 신규 } = _Sheets에등록(건, '수동입력');
  if (!신규) {
    SpreadsheetApp.getUi().alert(`이미 등록된 접수번호입니다 (건너뜀): ${접수번호}`);
    return;
  }

  // 인공지능 기능 상세도 함께 등록
  AI기능상세등록(접수번호, [
    {
      기능번호: 1, 기능명: '흉부 영상 이상 탐지',
      인공지능역할: '흉부 X-ray 이미지에서 이상 소견 후보 탐지',
      입력: 'DICOM 영상, 검사 ID, 메타데이터',
      출력: 'Bounding Box, 위험도 점수, 소견 라벨',
      구현방식: 'A', 연산자원요약: 'NVIDIA A100 40GB',
      실행환경요약: 'Ubuntu 22.04, CUDA 12.1, PyTorch 2.6.0',
      입력데이터설명: 'DICOM 영상, 촬영 메타데이터',
      출력데이터설명: 'Bounding Box 및 위험도 점수',
    },
    {
      기능번호: 2, 기능명: '폐렴 위험도 분류',
      인공지능역할: '폐렴 의심 정도와 중증도 등급 분류',
      입력: '흉부 X-ray, 환자 나이대, 촬영 조건',
      출력: '폐렴 의심 등급, 중증도, 확신도',
      구현방식: 'B', BaseModel명칭: 'ResNet-50 v2', 튜닝방법: '전이학습 + LoRA',
      입력데이터설명: '흉부 X-ray, 환자 나이대, 촬영 조건',
      출력데이터설명: '폐렴 의심 등급, 중증도, 확신도',
    },
    {
      기능번호: 3, 기능명: '판독 소견 요약 생성',
      인공지능역할: '탐지 결과와 임상 메모 기반 판독 초안 생성',
      입력: '탐지 결과 JSON, 임상 메모 텍스트',
      출력: '판독 초안, 주요 근거 요약',
      구현방식: 'C', 외부API정보: 'OpenAI gpt-4o-mini',
      입력데이터설명: '탐지 결과 JSON, 임상 메모 텍스트',
      출력데이터설명: '판독 초안 텍스트',
    },
  ]);

  SpreadsheetApp.getUi().alert(`수동 등록 완료!\n접수번호: ${접수번호}`);
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
