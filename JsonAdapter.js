/**
 * ============================================================
 *  JSON 신청서 어댑터 — KOSA JSON Data Schema v1.0 → 기존 시트 등록 파이프라인 연결
 * ============================================================
 *
 *  이 파일은 검증·변환·시트 등록 등 "처리 로직"만 담당합니다. JSON 필드
 *  경로·타입·시트컬럼·enum 매핑·formatter 같은 "정의"는 전부 FieldDefinition.js에
 *  있습니다(엑셀 파싱의 excelAliases도 같은 파일에서 같이 관리 — ExcelParser.js가 참조).
 *  Apps Script는 프로젝트 내 모든 .js 파일이 하나의 전역 스코프를 공유하므로,
 *  이 파일들은 import 없이 서로 참조합니다.
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
// 1. 오케스트레이터
//    구조 검증 → 전체 변환(건/제품모델/기능목록) → 쓰기 가능 여부 사전 검사(A안)
//    → 여기까지 통과해야 실제 쓰기 시작. 즉 쓰기가 시작된 이후에는 실패하지 않는다.
// ─────────────────────────────────────────────
function JSON파싱실행_내부(json, 파일명) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _JSON구조검증(json);

  const 건 = _JSON을건객체로변환(json);
  const 접수번호 = String(건.접수번호 || '').trim();
  if (!접수번호) {
    throw new Error('접수번호(keyId)가 없습니다. keyId가 있는 JSON만 등록할 수 있습니다.');
  }

  const 모델목록 = _JSON에서제품모델목록추출(json);
  const 기능목록 = _JSON에서기능목록추출(json);

  // ── 이슈 2, A안: 등록 시작 전 4개 시트 모두 쓰기 가능 여부 사전 검사 ──
  _쓰기가능여부검사(ss, {
    [SHEET.접수대장]: 1,
    [SHEET.일정관리]: 1,
    [SHEET.제품모델]: Math.max(1, 모델목록.length),
    [SHEET.AI기능상세]: Math.max(1, 기능목록.length),
  });

  const 결과 = _Sheets에등록(건, 파일명);
  if (!결과.신규) {
    return { 접수번호, 신규: false };
  }

  if (모델목록.length) 제품모델등록(접수번호, 모델목록);

  if (기능목록.length) {
    const 등록됨 = AI기능상세등록(접수번호, 기능목록);
    if (등록됨) _접수대장기능수갱신(접수번호, 기능목록);
  }

  return { 접수번호, 신규: true };
}

// ─────────────────────────────────────────────
// 2. 등록 전 사전 검증 (이슈 2, A안 — rollback 방식은 채택하지 않음)
//    새로 추가될 행(들)의 전체 컬럼 범위가 현재 실행 계정 기준으로 편집
//    가능한지 4개 시트 모두 확인한다. 하나라도 막히면 아무 시트에도 쓰지 않고
//    어느 시트의 어느 범위 때문인지 명시한 오류로 등록 자체를 시작하지 않는다.
// ─────────────────────────────────────────────
function _쓰기가능여부검사(ss, 시트별추가행수) {
  const 실패목록 = [];

  Object.keys(시트별추가행수).forEach(이름 => {
    const 시트 = ss.getSheetByName(이름);
    if (!시트) {
      실패목록.push(`"${이름}" 시트를 찾을 수 없습니다.`);
      return;
    }

    const lastCol = Math.max(1, 시트.getLastColumn());
    const 시작행 = 시트.getLastRow() + 1;
    const 끝행 = 시작행 + Math.max(1, 시트별추가행수[이름]) - 1;

    // 시트 전체 보호 — 예외(비보호) 범위에 우리가 쓸 행 전체가 포함되면 통과
    시트.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(보호 => {
      if (보호.canEdit()) return;
      const 예외범위목록 = 보호.getUnprotectedRanges() || [];
      const 예외로완전히덮임 = 예외범위목록.some(r =>
        r.getRow() <= 시작행 && r.getLastRow() >= 끝행 &&
        r.getColumn() <= 1 && r.getLastColumn() >= lastCol
      );
      if (!예외로완전히덮임) {
        실패목록.push(`"${이름}" 시트 전체가 보호되어 있어 새 행(${시작행}~${끝행}행)을 쓸 수 없습니다.`);
      }
    });

    // 컬럼 단위 범위 보호 — 새로 쓸 행 구간과 겹치는 보호 범위만 확인
    시트.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(보호 => {
      if (보호.canEdit()) return;
      const p = 보호.getRange();
      const 행겹침 = p.getRow() <= 끝행 && p.getLastRow() >= 시작행;
      const 열겹침 = p.getColumn() <= lastCol && p.getLastColumn() >= 1;
      if (행겹침 && 열겹침) {
        const 시작열 = Math.max(1, p.getColumn());
        const 끝열 = Math.min(lastCol, p.getLastColumn());
        const 컬럼명 = 시트.getRange(1, 시작열, 1, 끝열 - 시작열 + 1).getValues()[0]
          .filter(String).join(', ') || `${시작열}~${끝열}열`;
        실패목록.push(`"${이름}" 시트의 "${컬럼명}" 컬럼이 보호되어 있어 새 행(${시작행}~${끝행}행)을 쓸 수 없습니다.`);
      }
    });
  });

  if (실패목록.length) {
    throw new Error(
      '등록을 시작할 수 없습니다 — 아래 시트/범위에 쓰기 권한이 없습니다:\n\n' +
      실패목록.join('\n') +
      '\n\n관리자에게 문의해 해당 범위의 보호 설정(편집 권한)을 확인한 뒤 다시 시도하세요.'
    );
  }
}

// ─────────────────────────────────────────────
// 3. 구조 검증 — FieldDefinition.js의 JSON필수구조정의를 그대로 순회
//    ※ 필수 객체/배열 누락은 "건 객체를 만들 수 없는" 치명적 결함이라 등록을
//      중단시킨다. 코드값 미등록(4번)과는 다른 성격.
// ─────────────────────────────────────────────
function _JSON구조검증(json) {
  const 오류목록 = [];

  if (!json || typeof json !== 'object') {
    오류목록.push('JSON 최상위가 객체가 아닙니다.');
  } else {
    JSON필수구조정의.forEach(정의 => {
      const 값 = json[정의.path];
      if (정의.type === 'array' && !Array.isArray(값)) {
        오류목록.push(`${정의.path} 배열이 없습니다.`);
      } else if (정의.type === 'object' && (!값 || typeof 값 !== 'object' || Array.isArray(값))) {
        오류목록.push(`${정의.path} 객체가 없습니다.`);
      }
    });
  }

  if (오류목록.length) {
    throw new Error('JSON 구조 검증 실패:\n' + 오류목록.join('\n'));
  }
}

// ─────────────────────────────────────────────
// 4. 코드값 매핑 (미등록 코드값 정책)
//    매핑표에 없는 코드값을 만나도 자동화는 중단하지 않는다. 원본값을 그대로
//    저장하고 Logger.log에 경고를 남긴다 — 운영자는 로그를 보고
//    FieldDefinition.js의 코드값매핑표에 신규 코드를 추가한다.
// ─────────────────────────────────────────────
function _코드값라벨매핑(카테고리, 원본코드, 로그용필드명) {
  const 원본 = String(원본코드 || '').trim();
  if (!원본) return '';

  const 테이블 = 코드값매핑표[카테고리] || {};
  const 라벨 = 테이블[원본.toUpperCase()];
  if (라벨) return 라벨;

  Logger.log(
    `[경고] 미등록 코드값 — 필드: ${로그용필드명 || 카테고리}, 값: "${원본}" ` +
    `(매핑표에 없어 원본값을 그대로 저장합니다. FieldDefinition.js의 코드값매핑표.${카테고리}에 추가 필요)`
  );
  return 원본;
}

// ─────────────────────────────────────────────
// 5. 필드 정의 해석 엔진 — FieldDefinition.js의 정의 배열을 실제 값으로 변환
// ─────────────────────────────────────────────

/** dot 경로로 값을 조회한다. path가 배열이면 앞에서부터 값이 있는 첫 후보를 채택(fallback). */
function _경로값(obj, path) {
  if (!path) return undefined;
  const 후보목록 = Array.isArray(path) ? path : [path];
  for (const 경로 of 후보목록) {
    const 값 = String(경로).split('.').reduce((acc, key) => (acc == null ? undefined : acc[key]), obj);
    if (값 !== undefined && 값 !== null && String(값).trim() !== '') return 값;
  }
  return undefined;
}

/** dot 경로의 원시값을 그대로 조회한다(배열/불리언 등 falsy-but-valid 값 판별용, fallback 없음). */
function _원시경로값(obj, path) {
  if (!path) return undefined;
  return String(path).split('.').reduce((acc, key) => (acc == null ? undefined : acc[key]), obj);
}

/**
 * 정의 하나를 obj(컨텍스트)에 적용해 시트에 쓸 값을 계산한다.
 * index는 반복행(제품모델/기능상세)에서 'index1' 타입에만 사용된다.
 */
function _필드값추출(obj, 정의, index) {
  if (Object.prototype.hasOwnProperty.call(정의, 'fixedValue')) return 정의.fixedValue;

  let 값;
  switch (정의.type) {
    case 'arrayLength': {
      const 배열 = _원시경로값(obj, 정의.path);
      값 = Array.isArray(배열) ? 배열.length : (정의.fallback ?? 0);
      break;
    }
    case 'arrayPresence': {
      const 배열 = _원시경로값(obj, 정의.path);
      값 = Array.isArray(배열) && 배열.length > 0 ? 정의.presentValue : 정의.absentValue;
      break;
    }
    case 'joinField': {
      const 배열 = _원시경로값(obj, 정의.path);
      값 = Array.isArray(배열)
        ? 배열.map(item => String((item && item[정의.field]) || '').trim()).filter(Boolean).join(정의.delimiter || '/')
        : '';
      break;
    }
    case 'boolean': {
      값 = _원시경로값(obj, 정의.path) === true ? '동의' : '미동의';
      break;
    }
    case 'enum': {
      값 = _코드값라벨매핑(정의.enumCategory, _경로값(obj, 정의.path), 정의.sheetColumn);
      break;
    }
    case 'date': {
      const 원시값 = _경로값(obj, 정의.path);
      값 = 원시값
        ? String(원시값).trim()
        : (정의.fallback === 'now' ? Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd') : '');
      break;
    }
    case 'index1': {
      const 원시값 = _원시경로값(obj, 정의.path);
      값 = (원시값 !== undefined && 원시값 !== null && 원시값 !== '') ? 원시값 : (Number(index) + 1);
      break;
    }
    default: {
      값 = String(_경로값(obj, 정의.path) ?? '').trim();
    }
  }

  if (정의.formatter && 포매터[정의.formatter]) {
    값 = 포매터[정의.formatter](값, obj, 정의);
  }
  return 값;
}

// ─────────────────────────────────────────────
// 6. JSON → "건" 객체 / 반복행 배열 변환 — 전부 정의표를 순회하는 범용 로직
// ─────────────────────────────────────────────
function _JSON을건객체로변환(json) {
  const 건 = {};
  필드정의.forEach(정의 => {
    건[정의.sheetColumn] = _필드값추출(json, 정의);
  });
  return 건;
}

function _JSON에서제품모델목록추출(json) {
  const 목록 = Array.isArray(json.productModels) ? json.productModels : [];
  return 목록.map((m, i) => {
    const 행 = { 연번: i + 1 };
    제품모델필드정의.forEach(정의 => {
      행[정의.sheetColumn] = _필드값추출(m || {}, 정의, i);
    });
    return 행;
  });
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
    const context = { f: keyFunctions[i] || {}, impl: aiImplementations[i] || {} };
    const 행 = {};
    기능상세필드정의.forEach(정의 => {
      행[정의.sheetColumn] = _필드값추출(context, 정의, i);
    });
    결과.push(행);
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
