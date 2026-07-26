/**
 * ============================================================
 *  시트 등록 — 표준 "건"/반복행 객체 → 접수대장·제품모델·기능상세 시트 기록
 * ============================================================
 *
 *  ExcelParser.js·JsonAdapter.js가 만든 표준 객체(sheetColumn을 키로 갖는)를
 *  실제 스프레드시트 행으로 쓰는 처리만 담당합니다. append-only 정책(이미
 *  등록된 접수번호는 덮어쓰지 않고 건너뜀)을 여기서 일괄 적용합니다.
 * ============================================================
 */

// 테이블 친화 행 추가
// appendRow는 시트 맨 끝에 붙어 테이블 범위 밖으로 나갈 수 있으므로,
// 마지막 데이터 행 바로 아래(= 테이블 확장 영역)에 값을 써서 테이블이 자동 흡수하게 합니다.
function _테이블행추가(sheet, 값배열) {
  const 마지막 = sheet.getLastRow();
  const 대상행 = 마지막 + 1;
  sheet.getRange(대상행, 1, 1, 값배열.length).setValues([값배열]);
  return 대상행;
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
