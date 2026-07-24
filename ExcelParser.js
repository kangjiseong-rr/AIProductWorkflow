/**
 * ============================================================
 *  엑셀 신청서 파싱 — 엑셀 파일 → 표준 "건"/반복행 객체 변환
 * ============================================================
 *
 *  KOSA 엑셀(세로형 1건 / 가로형 N건 두 포맷 모두 대응)을 파싱해 SheetWriter.js가
 *  기대하는 형태의 건 객체·반복행 배열로 변환합니다. 필드 경로·별칭 등 "정의"는
 *  FieldDefinition.js에만 있고, 여기는 그 정의를 실제 엑셀 헤더에 매칭하는
 *  "처리 로직"만 담당합니다(JsonAdapter.js와 대칭 구조).
 *
 *  담당 범위:
 *   - 컬럼매핑 메타 시트 (엑셀 헤더 별칭을 코드 수정 없이 관리)
 *   - 엑셀 파일 등록 메뉴 진입점 + Drive → Sheets 변환
 *   - 세로형(항목-값)/가로형(헤더+N행) 자동 감지 및 파싱
 *   - 기능상세 탭·제품모델 탭 파싱 (접수번호/제품명으로 건에 연결)
 * ============================================================
 */

// ─────────────────────────────────────────────
// 1-1. 컬럼매핑 메타 시트 — 엑셀 헤더 별칭을 코드 수정 없이 관리
// ─────────────────────────────────────────────
// KOSA 엑셀의 컬럼명이 바뀌면 코드가 아니라 이 시트의 '별칭' 칸만 고치면 됩니다.
// A열: 내부키 (코드가 쓰는 표준 이름)  B열: 별칭 (쉼표로 구분, 부분일치 허용)

const 컬럼매핑시트명 = '컬럼매핑';

/**
 * 기본값 시드는 FieldDefinition.js(필드정의·제품모델필드정의·기능상세필드정의)의
 * excelAliases에서 그대로 생성한다 — 여기서 별도로 별칭 목록을 하드코딩하지 않는다.
 * (세 정의 배열은 sheetColumn이 서로 겹치지 않아 단순 concat으로 충분하다.)
 */
function _컬럼매핑기본값_() {
  const 행목록 = [['내부키', '별칭 (쉼표 구분 — 엑셀 헤더가 바뀌면 여기만 수정)']];
  [].concat(필드정의, 제품모델필드정의, 기능상세필드정의).forEach(정의 => {
    if (정의.excelAliases && 정의.excelAliases.length) {
      행목록.push([정의.sheetColumn, 정의.excelAliases.join(', ')]);
    }
  });
  return 행목록;
}

function _컬럼매핑시트확보(ss) {
  let 시트 = ss.getSheetByName(컬럼매핑시트명);
  if (시트) return 시트;
  시트 = ss.insertSheet(컬럼매핑시트명);
  const 기본값 = _컬럼매핑기본값_();
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

/**
 * 헤더 배열에서 정의.sheetColumn 하나의 위치를 찾는다.
 * 정확일치를 모두 시도한 뒤에만 부분일치로 폴백 — '입력'이 '입력데이터설명'을
 * 잘못 잡는 것처럼, 짧은 별칭이 긴 컬럼명을 오매칭하는 사고를 방지한다.
 */
function _헤더위치찾기(정의, 헤더) {
  if (!정의 || !정의.excelAliases) return -1;
  const 후보 = _별칭합치기(정의.sheetColumn, 정의.excelAliases);
  for (const c of 후보) {
    const i = 헤더.findIndex(h => h === c);
    if (i >= 0) return i;
  }
  for (const c of 후보) {
    const i = 헤더.findIndex(h => h.includes(c));
    if (i >= 0) return i;
  }
  return -1;
}

/** FieldDefinition.js 정의배열(엑셀 대상 항목만) → { sheetColumn: 헤더인덱스(-1이면 없음) } */
function _정의헤더인덱스맵(정의배열, 헤더) {
  const 맵 = {};
  정의배열.forEach(정의 => {
    if (!정의.excelAliases) return;
    맵[정의.sheetColumn] = _헤더위치찾기(정의, 헤더);
  });
  return 맵;
}

/** 정의배열에서 sheetColumn으로 정의 하나를 찾는다 (접수번호/제품명처럼 다른 시트 파싱에서도 재사용할 때) */
function _정의찾기(정의배열, sheetColumn) {
  return 정의배열.find(d => d.sheetColumn === sheetColumn);
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

  // 기능상세 컬럼은 FieldDefinition.js의 기능상세필드정의(excelAliases)로 찾는다.
  // 행 연결용 접수번호·제품명은 "건" 정의(필드정의)의 같은 항목을 그대로 재사용.
  const I = _정의헤더인덱스맵(기능상세필드정의, 헤더);
  const i접수번호 = _헤더위치찾기(_정의찾기(필드정의, '접수번호'), 헤더);
  const i제품명 = _헤더위치찾기(_정의찾기(필드정의, '제품명'), 헤더);

  // 접수번호별로 기능 목록 묶기
  const 묶음 = {};  // 접수번호 → [기능객체...]
  for (let r = 1; r < 데이터.length; r++) {
    const 행 = 데이터[r];
    if (행.every(v => v === '' || v == null)) continue;

    // 접수번호 컬럼이 있으면 그걸로 직접 연결 (가장 정확), 없으면 제품명으로
    let 접수번호 = '';
    if (i접수번호 >= 0) {
      const 직접 = String(행[i접수번호] ?? '').trim();
      if (직접 && 제품명별접수번호['__접수__' + 직접]) 접수번호 = 직접;
    }
    if (!접수번호 && i제품명 >= 0) {
      const 제품명 = String(행[i제품명]).trim();
      접수번호 = 제품명별접수번호[제품명];
    }
    if (!접수번호) continue;  // 연결 안 되는 행은 스킵

    const 기능 = {};
    기능상세필드정의.forEach(정의 => {
      const i = I[정의.sheetColumn];
      기능[정의.sheetColumn] = i >= 0 ? String(행[i] ?? '').trim() : '';
    });
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

  // 제품모델 컬럼은 제품모델필드정의로 찾는다. 연번은 JSON 쪽과 마찬가지로
  // 배열 인덱스 기반이라 정의 배열에 두지 않고 여기서 직접 찾는다(FieldDefinition.js 5번 참고).
  const I = _정의헤더인덱스맵(제품모델필드정의, 헤더);
  const i연번 = _헤더위치찾기({ sheetColumn: '연번', excelAliases: ['연번'] }, 헤더);
  const i접수번호 = _헤더위치찾기(_정의찾기(필드정의, '접수번호'), 헤더);
  const i제품명 = _헤더위치찾기(_정의찾기(필드정의, '제품명'), 헤더);

  const 묶음 = {};
  for (let r = 1; r < 데이터.length; r++) {
    const 행 = 데이터[r];
    if (행.every(v => v === '' || v == null)) continue;

    let 접수번호 = '';
    if (i접수번호 >= 0) {
      const 직접 = String(행[i접수번호] ?? '').trim();
      if (직접 && 제품명별접수번호['__접수__' + 직접]) 접수번호 = 직접;
    }
    if (!접수번호 && i제품명 >= 0) {
      const 제품명 = String(행[i제품명]).trim();
      접수번호 = 제품명별접수번호[제품명];
    }
    if (!접수번호) continue;

    const get = (i) => (i >= 0 ? String(행[i] ?? '').trim() : '');
    const 모델 = { 연번: get(i연번) };
    제품모델필드정의.forEach(정의 => { 모델[정의.sheetColumn] = get(I[정의.sheetColumn]); });
    (묶음[접수번호] = 묶음[접수번호] || []).push(모델);
  }

  Object.keys(묶음).forEach(접수번호 => 제품모델등록(접수번호, 묶음[접수번호]));
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

  // 건 정의(필드정의)를 그대로 순회 — excelAliases가 있는 항목만 읽는다.
  // (제품수처럼 엑셀에서 직접 읽지 않고 별도 계산되는 컬럼은 자동으로 제외됨)
  const 건 = {};
  필드정의.forEach(정의 => {
    if (!정의.excelAliases) return;
    건[정의.sheetColumn] = _맵값(맵, _별칭합치기(정의.sheetColumn, 정의.excelAliases));
  });
  건.원본맵 = 맵;
  return 건;
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
