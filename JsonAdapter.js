/**
 * ============================================================
 *  JSON 신청서 어댑터 — KOSA JSON Data Schema v1.0 → 기존 시트 등록 파이프라인 연결
 * ============================================================
 *
 *  이 파일은 Code.js를 수정하지 않고 JSON 입력을 지원하기 위한 별도 모듈입니다.
 *  Apps Script는 프로젝트 내 모든 .js/.gs 파일이 하나의 전역 스코프를 공유하므로,
 *  아래 함수들은 Code.js의 SHEET/CONFIG 상수와 _Sheets에등록/제품모델등록/
 *  AI기능상세등록/_접수대장기능수갱신/_드라이브ID추출을 import 없이 그대로 재사용합니다.
 *
 *  변환 규칙·정책 결정 사항의 근거는 JSON_ADAPTER_GUIDE.md를 참고하세요.
 * ============================================================
 */

// ─────────────────────────────────────────────
// 0. 메뉴 진입점
// ─────────────────────────────────────────────
function JSON파싱등록() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'JSON 파일 등록',
    'JSON 신청서 파일의 Drive 파일 ID 또는 공유 URL을 입력하세요.\n' +
    '(.json 파일만 지원합니다.)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const 입력값 = response.getResponseText().trim();
  const 파일ID = _드라이브ID추출(입력값);

  try {
    const file = DriveApp.getFileById(파일ID);
    const json = JSON.parse(file.getBlob().getDataAsString());
    const 결과 = JSON파싱실행_내부(json, file.getName());

    if (결과.신규) {
      ui.alert(`JSON 등록 완료\n\n접수번호: ${결과.접수번호}`);
    } else {
      ui.alert(`이미 등록된 접수번호입니다 — append-only 정책으로 건너뜀\n\n접수번호: ${결과.접수번호}`);
    }
  } catch (e) {
    ui.alert('JSON 등록 실패: ' + e.message);
  }
}

// ─────────────────────────────────────────────
// 1. 오케스트레이터 — _엑셀파싱처리와 동일한 순서로 기존 등록 함수 호출
// ─────────────────────────────────────────────
function JSON파싱실행_내부(json, 파일명) {
  _JSON구조검증(json);

  const 접수번호 = _접수번호결정(json);
  const 건 = _JSON을건객체로변환(json, 접수번호);

  const 결과 = _Sheets에등록(건, 파일명);
  if (!결과.신규) {
    return { 접수번호, 신규: false };
  }

  const 모델목록 = _JSON에서제품모델목록추출(json);
  if (모델목록.length) 제품모델등록(접수번호, 모델목록);

  const 기능목록 = _JSON에서기능목록추출(json);
  if (기능목록.length) {
    const 등록됨 = AI기능상세등록(접수번호, 기능목록);
    if (등록됨) _접수대장기능수갱신(접수번호, 기능목록);
  }

  return { 접수번호, 신규: true };
}

// ─────────────────────────────────────────────
// 2. 구조 검증
//    ※ 핵심 객체(applicant/serviceInfo/agreements)나 배열 자체가 없는 경우는
//      "건 객체를 만들 수 없는" 치명적 결함이므로 등록을 중단합니다.
//      아래 3번의 "미등록 코드값" 정책(자동화 계속 진행)과는 다른 성격입니다.
// ─────────────────────────────────────────────
function _JSON구조검증(json) {
  const 오류목록 = [];
  if (!json || typeof json !== 'object') 오류목록.push('JSON 최상위가 객체가 아닙니다.');
  if (!json.applicant) 오류목록.push('applicant 객체가 없습니다.');
  if (!json.serviceInfo) 오류목록.push('serviceInfo 객체가 없습니다.');
  if (!json.agreements) 오류목록.push('agreements 객체가 없습니다.');
  if (!Array.isArray(json.productModels)) 오류목록.push('productModels 배열이 없습니다.');
  if (!Array.isArray(json.keyFunctions)) 오류목록.push('keyFunctions 배열이 없습니다.');
  if (!Array.isArray(json.aiImplementations)) 오류목록.push('aiImplementations 배열이 없습니다.');

  if (오류목록.length) {
    throw new Error('JSON 구조 검증 실패:\n' + 오류목록.join('\n'));
  }
}

// ─────────────────────────────────────────────
// 3. 코드값 매핑표 + 미등록 코드값 정책
//    정책: 매핑표에 없는 코드값을 만나도 자동화를 중단하지 않는다.
//         원본 코드값을 그대로 저장하고 Logger.log에 경고를 남긴다.
//         운영자는 로그를 보고 아래 매핑표에 신규 코드를 추가한다.
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

function _코드값라벨매핑(카테고리, 원본코드, 로그용필드명) {
  const 원본 = String(원본코드 || '').trim();
  if (!원본) return '';

  const 테이블 = 코드값매핑표[카테고리] || {};
  const 라벨 = 테이블[원본.toUpperCase()];
  if (라벨) return 라벨;

  Logger.log(
    `[경고] 미등록 코드값 — 필드: ${로그용필드명 || 카테고리}, 값: "${원본}" ` +
    `(매핑표에 없어 원본값을 그대로 저장합니다. 코드값매핑표.${카테고리}에 추가 필요)`
  );
  return 원본;
}

// ─────────────────────────────────────────────
// 4. 개별 변환 함수 (자동 변환 허용 항목)
// ─────────────────────────────────────────────
function _Boolean동의매핑(값) {
  return 값 === true ? '동의' : '미동의';
}

function _연락처병합(applicant) {
  applicant = applicant || {};
  return String(applicant.managerMobile || applicant.managerTel || '').trim();
}

function _사업자번호정규화(원본) {
  const 숫자만 = String(원본 || '').replace(/[^0-9]/g, '');
  if (숫자만.length !== 10) return String(원본 || '').trim();
  return `${숫자만.slice(0, 3)}-${숫자만.slice(3, 5)}-${숫자만.slice(5)}`;
}

// ── 관리자 결정 반영 ──────────────────────────
function _접수번호결정(json) {
  // 관리자 결정: keyId(Base36)를 접수번호로 그대로 채택. 엑셀 경로(AI2026...)와
  // JSON 경로(Base36)가 서로 다른 형식이어도 각자 원형을 그대로 노출한다.
  return String((json && json.keyId) || '').trim();
}

function _신청일결정(json) {
  // 관리자 결정: JSON에 신청일 필드가 없으므로 TTA 파싱 시각으로 대체.
  // TODO: KOSA가 향후 submittedAt(또는 동일 의미 필드)을 추가하면
  //       아래 후보 목록에 필드명만 추가하면 된다.
  const 후보 = (json && (json.submittedAt || json.applicationDate || json.appliedAt)) || '';
  if (후보) return String(후보).trim();
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function _보유인증단순화(json) {
  // 관리자 결정: AuthType 정책 확정 전까지 파일 존재 여부로만 단순화.
  // TODO: KOSA의 AuthType(GS/CAT/해당없음 등) 필드 위치가 확정되면 이 함수만 교체.
  const 파일목록 = (json && json.certification && json.certification.certFiles) || [];
  return 파일목록.length > 0 ? '보유(상세 확인 필요)' : '해당없음';
}

// ─────────────────────────────────────────────
// 5. JSON → "건" 객체 변환 (_파싱_세로형과 동일한 키 구조)
// ─────────────────────────────────────────────
function _JSON을건객체로변환(json, 접수번호) {
  const applicant = json.applicant || {};
  const serviceInfo = json.serviceInfo || {};
  const agreements = json.agreements || {};
  const attachments = json.attachments || {};

  return {
    // 기업 정보
    기업명: String(applicant.companyNm || '').trim(),
    사업자번호: _사업자번호정규화(applicant.businessNo),
    대표자: String(applicant.ceoNm || '').trim(),
    소재지: String(applicant.companyAddr || '').trim(),
    담당자명: String(applicant.managerNm || '').trim(),
    연락처: _연락처병합(applicant),
    이메일: String(applicant.managerEmail || '').trim(),

    // 제품·서비스 정보
    접수번호,
    제품명: String(serviceInfo.productNm || '').trim(),
    제품수: Array.isArray(json.productModels) ? json.productModels.length : 1,
    제공형태: _코드값라벨매핑('제공형태', serviceInfo.serviceType, 'serviceInfo.serviceType'),
    제품분류: _코드값라벨매핑('제품분류', serviceInfo.aiTechCategory, 'serviceInfo.aiTechCategory'),
    개요: String(serviceInfo.productSummary || '').trim(),
    인공지능적용목적: String(serviceInfo.aiTechPurpose || '').trim(),
    인공지능적용범위: String(serviceInfo.aiTechAppliedArea || '').trim(),

    // 첨부파일 — 관리자 결정: JSON 경로에서는 처리 제외(빈 값 유지)
    구조도파일명: '',
    명세서파일명: '',
    기타제출서류파일명: '',
    인증서파일명: '',

    명세서작성방식: _코드값라벨매핑('명세서작성방식', attachments.submitMethodWrite, 'attachments.submitMethodWrite'),
    보유인증: _보유인증단순화(json),

    // 동의서·특기사항
    열람이용동의여부: _Boolean동의매핑(agreements.documentConsent),
    개인정보수집이용동의여부: _Boolean동의매핑(agreements.privacyConsent),
    개인정보3자제공동의여부: _Boolean동의매핑(agreements.thirdPartyConsent),
    특기사항: String(serviceInfo.remark || '').trim(),

    신청일: _신청일결정(json),

    // 파서 내부 보조 (인공지능기능명(요약) 초기값 — 등록 후 _접수대장기능수갱신이 재계산)
    AI기능명_원본: (Array.isArray(json.keyFunctions) ? json.keyFunctions : [])
      .map(f => String((f && f.funcNm) || '').trim())
      .filter(Boolean)
      .join('/'),
    비고: '',
  };
}

// ─────────────────────────────────────────────
// 6. 제품모델 / 기능상세 배열 추출
// ─────────────────────────────────────────────
function _JSON에서제품모델목록추출(json) {
  const 목록 = Array.isArray(json.productModels) ? json.productModels : [];
  return 목록.map((m, i) => ({
    연번: i + 1,
    모델명: String((m && m.modelNm) || '').trim(),
    세부품명번호: String((m && m.detailItemNo) || '').trim(),
    물품식별번호: String((m && m.productIdNo) || '').trim(),
  }));
}

/**
 * keyFunctions[](기능명·역할·입출력 요약)와 aiImplementations[](구현방식 세부)는
 * JSON 스키마상 별도 배열로 분리돼 있고 공통 순번 필드가 없어, 배열 인덱스
 * 순서로 조인한다(이름 텍스트 매칭은 표기 차이에 취약해 사용하지 않음).
 * 두 배열 길이가 다르면 경고만 남기고 짧은 쪽은 빈 값으로 채워 계속 진행한다.
 */
function _JSON에서기능목록추출(json) {
  const keyFunctions = Array.isArray(json.keyFunctions) ? json.keyFunctions : [];
  const aiImplementations = Array.isArray(json.aiImplementations) ? json.aiImplementations : [];

  if (keyFunctions.length !== aiImplementations.length) {
    Logger.log(
      `[경고] keyFunctions(${keyFunctions.length}건)와 aiImplementations(${aiImplementations.length}건) ` +
      `개수가 일치하지 않습니다 — 배열 인덱스 기준으로 가능한 만큼만 조인합니다.`
    );
  }

  const 개수 = Math.max(keyFunctions.length, aiImplementations.length);
  const 결과 = [];

  for (let i = 0; i < 개수; i++) {
    const f = keyFunctions[i] || {};
    const impl = aiImplementations[i] || {};

    결과.push({
      기능번호: impl.aiFunctionNumber || (i + 1),
      기능명: String(impl.aiFunctionName || f.funcNm || '').trim(),
      인공지능역할: String(f.funcRole || '').trim(),
      입력: String(f.inputData || '').trim(),
      출력: String(f.outputData || '').trim(),
      레퍼런스참조위치: String(f.reference || '').trim(),
      구현방식: String(impl.aiImplementationMethod || '').trim(),
      연산자원요약: String(impl.aiComputeResource || '').trim(),
      실행환경요약: String(impl.aiRuntimeDetail || '').trim(),
      학습데이터사양: String(impl.learningDataSpec || '').trim(),
      개발환경라이브러리알고리즘: String(impl.developmentEnvironment || '').trim(),
      BaseModel명칭: String(impl.baseModelName || '').trim(),
      튜닝방법: String(impl.tuningMethod || '').trim(),
      튜닝데이터셋: String(impl.tuningDataSet || '').trim(),
      외부API정보: String(impl.outApiModel || '').trim(),
      타겟HW_OS: String(impl.targetDeviceHardware || '').trim(),
      추론런타임: String(impl.aiModelRuntimeEngine || '').trim(),
      혼합구성설명: String(impl.mixCompositionExplain || '').trim(),
      모델별역할및입출력흐름: String(impl.modelSpecificRoles || '').trim(),
      입력데이터설명: String(impl.inputDataDescription || '').trim(),
      출력데이터설명: String(impl.outputDataDescription || '').trim(),
      기타참고자료파일명: '',
    });
  }

  return 결과;
}

// ─────────────────────────────────────────────
// 7. DEV 전용 테스트 하네스
//    ⚠ 실제 활성 스프레드시트에 행을 씁니다. 반드시 dev 프로젝트에서만 실행하세요.
// ─────────────────────────────────────────────
function testJsonImport() {
  const 샘플 = {
    keyId: '92EM5EO4MV',
    status: 'SUBMITTED',
    applicant: {
      companyNm: '(주)에이아이테크',
      businessNo: '1234567890',
      ceoNm: '홍길동',
      companyAddr: '서울특별시 강남구 테헤란로 123, 에이아이빌딩 5층',
      managerNm: '김철수',
      managerTel: '02-123-4567',
      managerMobile: '010-1234-5678',
      managerEmail: 'chulsoo.kim@aitech.com',
    },
    serviceInfo: {
      productNm: 'VisionAI-Detector v2.0',
      productSummary: '딥러닝 기반 실시간 객체 인식 및 분류 솔루션',
      aiTechPurpose: 'CCTV 영상을 분석하여 이상 징후를 자동으로 감지하고 관리자에게 알림 제공',
      serviceType: 'SAAS',
      serviceTypeEtc: '',
      aiTechAppliedArea: '보안, 관제, 스마트시티',
      aiTechCategory: 'SOFTWARE',
      aiTechCategoryEtc: '',
      serviceDomain: '공공안전',
      remark: '본 제품은 기존 GS인증 1등급을 획득한 엔진을 기반으로 함',
    },
    productModels: [
      { modelNm: 'VAI-DET-100', detailItemNo: '4322269101', productIdNo: '23456789' },
      { modelNm: 'VAI-DET-200 (High-End)', detailItemNo: '4322269102', productIdNo: '23456790' },
    ],
    keyFunctions: [
      {
        funcNm: '실시간 객체 탐지',
        funcRole: '입력 영상에서 사람, 차량, 화재 등 80여 종의 객체를 실시간으로 식별',
        inputData: 'RTSP 스트리밍 영상 (Full HD)',
        outputData: '객체 좌표(Bounding Box) 및 클래스 정보 (JSON)',
        reference: '사용자 매뉴얼 p.15~20',
      },
      {
        funcNm: '행동 패턴 분석',
        funcRole: '객체의 움직임을 추적하여 배회, 침입, 쓰러짐 등의 행위 판별',
        inputData: '식별된 객체 트래킹 데이터',
        outputData: '이벤트 발생 로그 및 알람',
        reference: '기술사양서 p.8',
      },
    ],
    aiImplementations: [
      {
        aiFunctionNumber: 1,
        aiFunctionName: '실시간 객체 탐지',
        aiImplementationMethod: 'A',
        aiComputeResource: 'NVIDIA RTX 4090 24GB / RAM 64GB',
        aiRuntimeDetail: 'Ubuntu 22.04, CUDA 11.8, TensorRT 8.5',
        learningDataSpec: 'COCO Dataset 및 자체 수집 보안 영상 데이터 50만 장',
        developmentEnvironment: 'PyTorch 2.0, Python 3.10',
        baseModelName: 'YOLOv8-Custom',
        tuningMethod: 'Fine-tuning',
        tuningDataSet: '국내 야간 환경 특화 데이터셋 5만 장',
        outApiModel: '',
        targetDeviceHardware: 'Edge AI Box (Jetson Orin)',
        aiModelRuntimeEngine: 'ONNX Runtime',
        mixCompositionExplain: '',
        modelSpecificRoles: '',
        inputDataDescription: '640x640 RGB Image',
        outputDataDescription: 'Class ID, Confidence Score, BBox Coordinates',
      },
    ],
    attachments: {
      submitMethodWrite: 'WRITE',
      structureFiles: [
        { fileId: 'FILE_001', LogicalFileNm: '시스템구성도.pdf', PhysicalFileNm: '20240125_92EM5_1.pdf', FileSize: 1024550, GrpId: '1' },
      ],
      specFiles: [],
      etcFiles: [],
    },
    certification: {
      certFiles: [
        { fileId: 'FILE_002', LogicalFileNm: 'GS인증서_1등급.jpg', PhysicalFileNm: '20240125_92EM5_2.jpg', FileSize: 512000, GrpId: '7' },
      ],
    },
    agreements: {
      documentConsent: true,
      privacyConsent: true,
      thirdPartyConsent: true,
      finalSubmissionConsent: true,
    },
  };

  const 결과 = JSON파싱실행_내부(샘플, 'testJsonImport() 수동 실행');
  Logger.log('테스트 결과: ' + JSON.stringify(결과));
}
