/**
 * ============================================================
 *  신청서 필드 정의 — 엑셀·JSON 공용 단일 정의 모듈
 * ============================================================
 *
 *  이 파일은 순수 정의(데이터)만 담습니다. 검증·변환·시트 등록 같은 처리
 *  로직은 전혀 두지 않습니다.
 *   - ExcelParser.js(엑셀 파싱)는 각 정의의 excelAliases만 참조합니다.
 *   - JsonAdapter.js(JSON 파싱)는 각 정의의 path/type/formatter 등만 참조합니다.
 *  즉 "이 sheetColumn은 엑셀에서 이런 헤더로, JSON에서 이런 path로 온다"를
 *  한 엔트리에서 같이 관리해, 엑셀/JSON 두 입력경로가 스키마를 따로
 *  들고 있지 않도록 합니다.
 *
 *  담당 범위:
 *   - JSON 최상위 필수 구조 (JSON필수구조정의) — JSON 전용 검증
 *   - "건" 객체(접수대장/일정관리 공용) 필드 정의 (필드정의)
 *   - 제품모델 반복행 필드 정의 (제품모델필드정의)
 *   - 기능상세 반복행 필드 정의 (기능상세필드정의)
 *   - 코드값 → 한글 라벨 매핑표 (코드값매핑표) — JSON 코드값(SAAS 등)에만
 *     적용. 엑셀은 이미 사람이 입력한 한글 값이 오므로 매핑 대상이 아님.
 *     미등록 코드값 정책은 JsonAdapter.js의 _코드값라벨매핑()에서 처리
 *   - 범용 type으로 표현하기 어려운 개별 포매터 (포매터)
 *
 *  필드 정의 스키마:
 *   { sheetColumn, excelAliases, path, type, required, formatter, ... }
 *   - sheetColumn   : 결과 객체의 키이자 시트 헤더명 (엑셀·JSON 공용 canonical 이름)
 *   - excelAliases  : 엑셀 헤더에서 이 필드를 찾을 때 시도할 후보 이름 목록.
 *                     정확일치 우선 → 부분일치 폴백(ExcelParser.js의 _정의헤더인덱스맵).
 *                     스프레드시트의 "컬럼매핑" 시트에서 사용자가 별칭을 추가하면
 *                     여기 목록과 합쳐진다(_별칭합치기). 엑셀에 대응 항목이
 *                     없는 필드(TTA 내부관리 등)는 생략.
 *   - path          : JSON 경로(dot 표기). 배열이면 앞에서부터 값이 있는 것을 채택(fallback).
 *                     JSON에 대응 항목이 없는 필드는 생략(fixedValue 등으로 대체).
 *   - type          : 'string'(기본) | 'enum' | 'boolean' | 'arrayLength' | 'arrayPresence'
 *                     | 'date' | 'joinField' | 'index1' — JSON 변환에만 적용.
 *   - required      : 문서화 목적 메타데이터 (입력폼_명세서.md 기준 필수여부). 현재는
 *                     구조 검증에는 쓰이지 않고 향후 필드 단위 검증을 붙일 때 참조하는 용도
 *   - formatter     : type 처리 후 추가로 적용할 포매터의 키 이름 (예: 사업자번호 재포맷).
 *                     JSON 변환에만 적용.
 *   - fixedValue    : path 없이 항상 고정값을 쓰는 경우(첨부파일 등, 관리자 결정으로 제외).
 *                     JSON 변환에만 적용 — 엑셀은 excelAliases로 별도 값을 읽을 수 있다.
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
// 2. 코드값 매핑표 (JSON 코드값 전용 — 엑셀은 이미 한글 라벨이 입력되어 있어 대상 아님)
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
// 3. 개별 포매터 (범용 type으로 표현하기 어려운 경우만, JSON 변환 전용)
// ─────────────────────────────────────────────
const 포매터 = {
  사업자번호정규화(값) {
    const 숫자만 = String(값 || '').replace(/[^0-9]/g, '');
    if (숫자만.length !== 10) return String(값 || '').trim();
    return `${숫자만.slice(0, 3)}-${숫자만.slice(3, 5)}-${숫자만.slice(5)}`;
  },
};

// ─────────────────────────────────────────────
// 4. "건" 객체 필드 정의 — 접수대장/일정관리 공용 (_Sheets에등록이 참조하는 키와 동일)
//    excelAliases는 기존 4곳(엑셀 세로형 파싱·기능탭 파싱·제품모델탭 파싱·컬럼매핑
//    기본시드)에 흩어져 있던 후보 목록을 합집합으로 병합한 것 — 표현이 갈리던
//    별칭도 여기 한 곳에 모여 어디서 파싱하든 동일하게 인식된다.
// ─────────────────────────────────────────────
const 필드정의 = [
  {
    sheetColumn: '기업명', path: 'applicant.companyNm', required: true,
    excelAliases: ['상호(사업자명)', '기업명'],
  },
  {
    sheetColumn: '사업자번호', path: 'applicant.businessNo', required: true, formatter: '사업자번호정규화',
    excelAliases: ['법인등록번호', '사업자등록번호', '사업자번호'],
  },
  {
    sheetColumn: '대표자', path: 'applicant.ceoNm', required: true,
    excelAliases: ['대표자 성명', '대표자명', '대표자'],
  },
  {
    sheetColumn: '소재지', path: 'applicant.companyAddr', required: true,
    excelAliases: ['소재지'],
  },
  {
    sheetColumn: '담당자명', path: 'applicant.managerNm', required: true,
    excelAliases: ['업무담당자', '담당자명', '담당자'],
  },
  // 관리자 결정: 휴대전화 우선, 없으면 일반전화 — path의 fallback 배열로 표현
  {
    sheetColumn: '연락처', path: ['applicant.managerMobile', 'applicant.managerTel'], required: true,
    excelAliases: ['담당자 전화번호', '연락처', '전화번호'],
  },
  {
    sheetColumn: '이메일', path: 'applicant.managerEmail', required: true,
    excelAliases: ['담당자 이메일주소', '이메일', 'email', 'e-mail'],
  },
  // 관리자 결정: keyId(Base36)를 접수번호로 그대로 채택
  {
    sheetColumn: '접수번호', path: 'keyId', required: true,
    excelAliases: ['접수번호', '접수 번호', '관리번호', '관리 번호', '심사번호', '신청번호'],
  },
  {
    sheetColumn: '제품명', path: 'serviceInfo.productNm', required: true,
    excelAliases: ['제품 또는 서비스 모델명', '제품 또는 서비스명', '제품·서비스명', '제품·서비스 모델명', '제품명', '서비스명'],
  },
  { sheetColumn: '제품수', path: 'productModels', type: 'arrayLength', fallback: 1, required: false },
  {
    sheetColumn: '제공형태', path: 'serviceInfo.serviceType', type: 'enum', enumCategory: '제공형태', required: true,
    excelAliases: ['제품 또는 서비스 제공 형태', '제공 형태', '제공형태'],
  },
  {
    sheetColumn: '제품분류', path: 'serviceInfo.aiTechCategory', type: 'enum', enumCategory: '제품분류', required: true,
    excelAliases: ['인공지능 제품·서비스 분류', 'AI 제품 분류', 'AI제품분류', '제품분류'],
  },
  {
    sheetColumn: '개요', path: 'serviceInfo.productSummary', required: true,
    excelAliases: ['개요'],
  },
  {
    sheetColumn: '인공지능적용목적', path: 'serviceInfo.aiTechPurpose', required: true,
    excelAliases: ['인공지능 적용 목적', 'AI 적용 목적', '인공지능적용목적'],
  },
  {
    sheetColumn: '인공지능적용범위', path: 'serviceInfo.aiTechAppliedArea', required: true,
    excelAliases: ['인공지능 적용 범위', 'AI 적용 범위', '인공지능적용범위'],
  },

  // 관리자 결정: 첨부파일(KOSA fileId ↔ Drive ID)은 JSON 경로에서 처리 제외 — 빈 값 유지.
  // 엑셀은 이 컬럼들에 파일명 텍스트가 그대로 들어오는 경우가 있어 excelAliases로 읽는다.
  { sheetColumn: '구조도파일명', fixedValue: '', required: false, excelAliases: ['제품 구조도', '인공지능 제품·서비스 구조도', '구조도파일명'] },
  { sheetColumn: '명세서파일명', fixedValue: '', required: false, excelAliases: ['명세서 파일', '명세서파일명'] },
  { sheetColumn: '기타제출서류파일명', fixedValue: '', required: false, excelAliases: ['기타 제출서류', '기타제출서류파일명'] },
  { sheetColumn: '인증서파일명', fixedValue: '', required: false, excelAliases: ['인증서·시험성적서', '인증서파일명'] },

  {
    sheetColumn: '명세서작성방식', path: 'attachments.submitMethodWrite', type: 'enum', enumCategory: '명세서작성방식', required: false,
    excelAliases: ['명세서 작성 방식', '명세서작성방식'],
  },
  // 관리자 결정: AuthType 정책 확정 전까지 파일 존재 여부로만 단순화
  {
    sheetColumn: '보유인증',
    path: 'certification.certFiles',
    type: 'arrayPresence',
    presentValue: '보유(상세 확인 필요)',
    absentValue: '해당없음',
    required: false,
    excelAliases: ['보유 인증', '보유인증'],
  },

  {
    sheetColumn: '열람이용동의여부', path: 'agreements.documentConsent', type: 'boolean', required: true,
    excelAliases: ['열람·이용동의', '열람이용동의여부'],
  },
  {
    sheetColumn: '개인정보수집이용동의여부', path: 'agreements.privacyConsent', type: 'boolean', required: true,
    excelAliases: ['개인정보 수집·이용', '개인정보수집이용동의여부'],
  },
  {
    sheetColumn: '개인정보3자제공동의여부', path: 'agreements.thirdPartyConsent', type: 'boolean', required: true,
    excelAliases: ['개인정보 제3자 제공', '개인정보3자제공동의여부'],
  },
  {
    sheetColumn: '특기사항', path: 'serviceInfo.remark', required: false,
    excelAliases: ['특기사항'],
  },

  // 관리자 결정: JSON에 신청일 필드가 없으므로 TTA 파싱 시각으로 대체.
  // TODO: KOSA가 submittedAt(또는 동일 의미 필드)을 추가하면 path 배열에 필드명만 추가.
  {
    sheetColumn: '신청일', path: ['submittedAt', 'applicationDate', 'appliedAt'], type: 'date', fallback: 'now', required: true,
    excelAliases: ['신청일자', '신청일'],
  },

  // 등록 후 _접수대장기능수갱신이 재계산 — 초기값 용도
  {
    sheetColumn: 'AI기능명_원본', path: 'keyFunctions', type: 'joinField', field: 'funcNm', delimiter: '/', required: false,
    excelAliases: ['AI 기능명', '기능명'],
  },
  { sheetColumn: '비고', fixedValue: '', required: false, excelAliases: ['비고', '기타'] },
];

// ─────────────────────────────────────────────
// 5. 제품모델 반복행 필드 정의 (productModels[] → 인공지능제품모델 시트)
//    연번은 배열 인덱스 기반이라 여기 두지 않고 JsonAdapter.js(JSON)와
//    ExcelParser.js의 _제품모델탭파싱등록(엑셀) 양쪽에서 각자 직접 채운다.
// ─────────────────────────────────────────────
const 제품모델필드정의 = [
  {
    sheetColumn: '모델명', path: 'modelNm', required: true,
    excelAliases: ['제품 또는 서비스 모델명', '제품·서비스 모델명', '모델명'],
  },
  {
    sheetColumn: '세부품명번호', path: 'detailItemNo', required: true,
    excelAliases: ['세부품명번호'],
  },
  {
    sheetColumn: '물품식별번호', path: 'productIdNo', required: true,
    excelAliases: ['물품식별번호'],
  },
];

// ─────────────────────────────────────────────
// 6. 기능상세 반복행 필드 정의 (keyFunctions[] + aiImplementations[] → 인공지능기능상세 시트)
//    JSON은 두 배열이 스키마상 분리돼 있어 공통 컨텍스트 { f: keyFunctions[i], impl: aiImplementations[i] }로
//    합친 뒤 path를 'f.xxx' / 'impl.xxx'로 참조한다(조인 자체는 JsonAdapter.js의 배열 인덱스 로직).
//    엑셀은 기능탭 한 행에 이미 모든 컬럼이 나열돼 있어 excelAliases만으로 바로 찾는다.
// ─────────────────────────────────────────────
const 기능상세필드정의 = [
  {
    sheetColumn: '기능번호', path: 'impl.aiFunctionNumber', type: 'index1', required: true,
    excelAliases: ['기능번호', '번호'],
  },
  {
    sheetColumn: '기능명', path: ['impl.aiFunctionName', 'f.funcNm'], required: true,
    excelAliases: ['기능명'],
  },
  {
    sheetColumn: '인공지능역할', path: 'f.funcRole', required: true,
    excelAliases: ['인공지능역할', '인공지능 역할', '역할', 'AI 역할'],
  },
  {
    sheetColumn: '입력', path: 'f.inputData', required: true,
    excelAliases: ['입력'],
  },
  {
    sheetColumn: '출력', path: 'f.outputData', required: true,
    excelAliases: ['출력'],
  },
  {
    sheetColumn: '레퍼런스참조위치', path: 'f.reference', required: true,
    excelAliases: ['레퍼런스참조위치', '레퍼런스 참조위치', '설명서참조위치', '설명서 참조 위치'],
  },
  {
    sheetColumn: '구현방식', path: 'impl.aiImplementationMethod', required: true,
    excelAliases: ['AI 구현 방식', '인공지능 구현 방식', '구현방식', '구현 방식'],
  },
  {
    sheetColumn: '연산자원요약', path: 'impl.aiComputeResource', required: true,
    excelAliases: ['AI 연산 자원 요약', '연산자원요약', '연산자원', '연산 자원'],
  },
  {
    sheetColumn: '실행환경요약', path: 'impl.aiRuntimeDetail', required: false,
    excelAliases: ['AI 실행 환경 요약', '실행환경요약', '실행환경', '실행 환경'],
  },
  {
    sheetColumn: '학습데이터사양', path: 'impl.learningDataSpec', required: false,
    excelAliases: ['학습데이터사양', '학습 데이터 사양', '학습데이터', '데이터명 라이선싱'],
  },
  {
    sheetColumn: '개발환경라이브러리알고리즘', path: 'impl.developmentEnvironment', required: false,
    excelAliases: ['개발환경·라이브러리·알고리즘', '개발환경라이브러리알고리즘'],
  },
  {
    sheetColumn: 'BaseModel명칭', path: 'impl.baseModelName', required: false,
    excelAliases: ['Base Model 명칭', 'BaseModel명칭', 'BaseModel', 'Base Model', '베이스'],
  },
  {
    sheetColumn: '튜닝방법', path: 'impl.tuningMethod', required: false,
    excelAliases: ['튜닝 방법', '튜닝방법'],
  },
  {
    sheetColumn: '튜닝데이터셋', path: 'impl.tuningDataSet', required: false,
    excelAliases: ['튜닝 데이터셋', '튜닝데이터셋'],
  },
  {
    sheetColumn: '외부API정보', path: 'impl.outApiModel', required: false,
    excelAliases: ['외부 API 정보', '외부API정보', '외부 API', '외부API', '외부API모델'],
  },
  {
    sheetColumn: '타겟HW_OS', path: 'impl.targetDeviceHardware', required: false,
    excelAliases: ['타겟 하드웨어·OS', '타겟HW_OS', '타겟', 'HW', 'OS'],
  },
  {
    sheetColumn: '추론런타임', path: 'impl.aiModelRuntimeEngine', required: false,
    excelAliases: ['추론 런타임', '추론런타임', '배포포맷정밀도'],
  },
  {
    sheetColumn: '혼합구성설명', path: 'impl.mixCompositionExplain', required: false,
    excelAliases: ['혼합 구성 설명', '혼합구성설명', '혼합'],
  },
  {
    sheetColumn: '모델별역할및입출력흐름', path: 'impl.modelSpecificRoles', required: false,
    excelAliases: ['모델별 역할 및 입출력 흐름', '모델별역할및입출력흐름', '세부구성요소별설명'],
  },
  {
    sheetColumn: '입력데이터설명', path: 'impl.inputDataDescription', required: true,
    excelAliases: ['입력데이터설명'],
  },
  {
    sheetColumn: '출력데이터설명', path: 'impl.outputDataDescription', required: true,
    excelAliases: ['출력데이터설명'],
  },
  { sheetColumn: '기타참고자료파일명', fixedValue: '', required: false, excelAliases: ['기타참고자료', '기타 참고자료'] },
];
