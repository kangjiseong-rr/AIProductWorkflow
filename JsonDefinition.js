/**
 * ============================================================
 *  JSON 신청서 정의 — KOSA JSON Data Schema v1.0 필드 정의 전용 모듈
 * ============================================================
 *
 *  이 파일은 순수 정의(데이터)만 담습니다. 검증·변환·시트 등록 같은 처리
 *  로직은 전혀 두지 않고, JsonAdapter.js가 이 정의들을 참조해 동작합니다.
 *
 *  담당 범위:
 *   - JSON 최상위 필수 구조 (JSON필수구조정의)
 *   - "건" 객체(접수대장/일정관리 공용) 필드 정의 (JSON필드정의)
 *   - 제품모델 반복행 필드 정의 (JSON제품모델필드정의)
 *   - 기능상세 반복행 필드 정의 (JSON기능상세필드정의)
 *   - 코드값 → 한글 라벨 매핑표 (코드값매핑표) — 미등록 코드값 정책은
 *     JsonAdapter.js의 _코드값라벨매핑()에서 처리 (여기는 데이터만)
 *   - 범용 type으로 표현하기 어려운 개별 포매터 (JSON포매터)
 *
 *  필드 정의 스키마:
 *   { sheetColumn, path, type, required, formatter, ... }
 *   - sheetColumn : 결과 객체의 키이자 시트 헤더명
 *   - path        : JSON 경로(dot 표기). 배열이면 앞에서부터 값이 있는 것을 채택(fallback)
 *   - type        : 'string'(기본) | 'enum' | 'boolean' | 'arrayLength' | 'arrayPresence'
 *                   | 'date' | 'joinField' | 'index1'
 *   - required    : 문서화 목적 메타데이터 (입력폼_명세서.md 기준 필수여부). 현재는
 *                   구조 검증에는 쓰이지 않고 향후 필드 단위 검증을 붙일 때 참조하는 용도
 *   - formatter   : type 처리 후 추가로 적용할 JSON포매터의 키 이름 (예: 사업자번호 재포맷)
 *   - fixedValue  : path 없이 항상 고정값을 쓰는 경우(첨부파일 등, 관리자 결정으로 제외)
 * ============================================================
 */

// ─────────────────────────────────────────────
// 1. JSON 최상위 필수 구조
//    ※ 여기 없는 실패(코드값 미등록 등)는 등록을 막지 않는 별개 정책 —
//      JSON_ADAPTER_GUIDE.md "미등록 코드값 정책" 참고.
// ─────────────────────────────────────────────
const JSON필수구조정의 = [
  { path: 'applicant', type: 'object' },
  { path: 'serviceInfo', type: 'object' },
  { path: 'agreements', type: 'object' },
  { path: 'productModels', type: 'array' },
  { path: 'keyFunctions', type: 'array' },
  { path: 'aiImplementations', type: 'array' },
];

// ─────────────────────────────────────────────
// 2. 코드값 매핑표
//    관리자 결정: 매핑표에 없는 코드값은 등록을 막지 않고 원본값을 그대로
//    저장 + Logger.log 경고(적용 로직은 JsonAdapter.js _코드값라벨매핑 참고).
// ─────────────────────────────────────────────
const 코드값매핑표 = {
  제공형태: {
    SAAS: 'SaaS(클라우드 서비스형)',
    ONPREMISE: '사내 구축형(온프레미스)',
    // TODO: CLOUD, ETC 등 KOSA 코드값 확정되는 대로 추가
  },
  제품분류: {
    SOFTWARE: '소프트웨어',
    SERVICE: '서비스',
    // TODO: PRODUCT, ETC 등 KOSA 코드값 확정되는 대로 추가
  },
  명세서작성방식: {
    WRITE: '직접제출',
    FILE: '파일제출',
  },
};

// ─────────────────────────────────────────────
// 3. 개별 포매터 (범용 type으로 표현하기 어려운 경우만)
// ─────────────────────────────────────────────
const JSON포매터 = {
  사업자번호정규화(값) {
    const 숫자만 = String(값 || '').replace(/[^0-9]/g, '');
    if (숫자만.length !== 10) return String(값 || '').trim();
    return `${숫자만.slice(0, 3)}-${숫자만.slice(3, 5)}-${숫자만.slice(5)}`;
  },
};

// ─────────────────────────────────────────────
// 4. "건" 객체 필드 정의 — 접수대장/일정관리 공용 (_Sheets에등록이 참조하는 키와 동일)
// ─────────────────────────────────────────────
const JSON필드정의 = [
  { sheetColumn: '기업명', path: 'applicant.companyNm', required: true },
  { sheetColumn: '사업자번호', path: 'applicant.businessNo', required: true, formatter: '사업자번호정규화' },
  { sheetColumn: '대표자', path: 'applicant.ceoNm', required: true },
  { sheetColumn: '소재지', path: 'applicant.companyAddr', required: true },
  { sheetColumn: '담당자명', path: 'applicant.managerNm', required: true },
  // 관리자 결정: 휴대전화 우선, 없으면 일반전화 — path의 fallback 배열로 표현
  { sheetColumn: '연락처', path: ['applicant.managerMobile', 'applicant.managerTel'], required: true },
  { sheetColumn: '이메일', path: 'applicant.managerEmail', required: true },
  // 관리자 결정: keyId(Base36)를 접수번호로 그대로 채택
  { sheetColumn: '접수번호', path: 'keyId', required: true },
  { sheetColumn: '제품명', path: 'serviceInfo.productNm', required: true },
  { sheetColumn: '제품수', path: 'productModels', type: 'arrayLength', fallback: 1, required: false },
  { sheetColumn: '제공형태', path: 'serviceInfo.serviceType', type: 'enum', enumCategory: '제공형태', required: true },
  { sheetColumn: '제품분류', path: 'serviceInfo.aiTechCategory', type: 'enum', enumCategory: '제품분류', required: true },
  { sheetColumn: '개요', path: 'serviceInfo.productSummary', required: true },
  { sheetColumn: '인공지능적용목적', path: 'serviceInfo.aiTechPurpose', required: true },
  { sheetColumn: '인공지능적용범위', path: 'serviceInfo.aiTechAppliedArea', required: true },

  // 관리자 결정: 첨부파일(KOSA fileId ↔ Drive ID)은 JSON 경로에서 처리 제외 — 빈 값 유지
  { sheetColumn: '구조도파일명', fixedValue: '', required: false },
  { sheetColumn: '명세서파일명', fixedValue: '', required: false },
  { sheetColumn: '기타제출서류파일명', fixedValue: '', required: false },
  { sheetColumn: '인증서파일명', fixedValue: '', required: false },

  { sheetColumn: '명세서작성방식', path: 'attachments.submitMethodWrite', type: 'enum', enumCategory: '명세서작성방식', required: false },
  // 관리자 결정: AuthType 정책 확정 전까지 파일 존재 여부로만 단순화
  {
    sheetColumn: '보유인증',
    path: 'certification.certFiles',
    type: 'arrayPresence',
    presentValue: '보유(상세 확인 필요)',
    absentValue: '해당없음',
    required: false,
  },

  { sheetColumn: '열람이용동의여부', path: 'agreements.documentConsent', type: 'boolean', required: true },
  { sheetColumn: '개인정보수집이용동의여부', path: 'agreements.privacyConsent', type: 'boolean', required: true },
  { sheetColumn: '개인정보3자제공동의여부', path: 'agreements.thirdPartyConsent', type: 'boolean', required: true },
  { sheetColumn: '특기사항', path: 'serviceInfo.remark', required: false },

  // 관리자 결정: JSON에 신청일 필드가 없으므로 TTA 파싱 시각으로 대체.
  // TODO: KOSA가 submittedAt(또는 동일 의미 필드)을 추가하면 path 배열에 필드명만 추가.
  { sheetColumn: '신청일', path: ['submittedAt', 'applicationDate', 'appliedAt'], type: 'date', fallback: 'now', required: true },

  // 등록 후 _접수대장기능수갱신이 재계산 — 초기값 용도
  { sheetColumn: 'AI기능명_원본', path: 'keyFunctions', type: 'joinField', field: 'funcNm', delimiter: '/', required: false },
  { sheetColumn: '비고', fixedValue: '', required: false },
];

// ─────────────────────────────────────────────
// 5. 제품모델 반복행 필드 정의 (productModels[] → 인공지능제품모델 시트)
//    연번은 배열 인덱스 기반이라 여기 두지 않고 JsonAdapter.js에서 직접 채움.
// ─────────────────────────────────────────────
const JSON제품모델필드정의 = [
  { sheetColumn: '모델명', path: 'modelNm', required: true },
  { sheetColumn: '세부품명번호', path: 'detailItemNo', required: true },
  { sheetColumn: '물품식별번호', path: 'productIdNo', required: true },
];

// ─────────────────────────────────────────────
// 6. 기능상세 반복행 필드 정의 (keyFunctions[] + aiImplementations[] → 인공지능기능상세 시트)
//    두 배열은 스키마상 분리돼 있어 공통 컨텍스트 { f: keyFunctions[i], impl: aiImplementations[i] }로
//    합친 뒤 path를 'f.xxx' / 'impl.xxx'로 참조한다 (조인 자체는 JsonAdapter.js의 배열 인덱스 로직).
// ─────────────────────────────────────────────
const JSON기능상세필드정의 = [
  { sheetColumn: '기능번호', path: 'impl.aiFunctionNumber', type: 'index1', required: true },
  { sheetColumn: '기능명', path: ['impl.aiFunctionName', 'f.funcNm'], required: true },
  { sheetColumn: '인공지능역할', path: 'f.funcRole', required: true },
  { sheetColumn: '입력', path: 'f.inputData', required: true },
  { sheetColumn: '출력', path: 'f.outputData', required: true },
  { sheetColumn: '레퍼런스참조위치', path: 'f.reference', required: true },
  { sheetColumn: '구현방식', path: 'impl.aiImplementationMethod', required: true },
  { sheetColumn: '연산자원요약', path: 'impl.aiComputeResource', required: true },
  { sheetColumn: '실행환경요약', path: 'impl.aiRuntimeDetail', required: false },
  { sheetColumn: '학습데이터사양', path: 'impl.learningDataSpec', required: false },
  { sheetColumn: '개발환경라이브러리알고리즘', path: 'impl.developmentEnvironment', required: false },
  { sheetColumn: 'BaseModel명칭', path: 'impl.baseModelName', required: false },
  { sheetColumn: '튜닝방법', path: 'impl.tuningMethod', required: false },
  { sheetColumn: '튜닝데이터셋', path: 'impl.tuningDataSet', required: false },
  { sheetColumn: '외부API정보', path: 'impl.outApiModel', required: false },
  { sheetColumn: '타겟HW_OS', path: 'impl.targetDeviceHardware', required: false },
  { sheetColumn: '추론런타임', path: 'impl.aiModelRuntimeEngine', required: false },
  { sheetColumn: '혼합구성설명', path: 'impl.mixCompositionExplain', required: false },
  { sheetColumn: '모델별역할및입출력흐름', path: 'impl.modelSpecificRoles', required: false },
  { sheetColumn: '입력데이터설명', path: 'impl.inputDataDescription', required: true },
  { sheetColumn: '출력데이터설명', path: 'impl.outputDataDescription', required: true },
  { sheetColumn: '기타참고자료파일명', fixedValue: '', required: false },
];
