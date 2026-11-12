/**
 * ============================================================
 *  사업자번호·법인등록번호 일괄 복구 (일회성 관리자 메뉴)
 * ============================================================
 *  예전 파서는 엑셀의 법인등록번호를 사업자번호보다 먼저 읽어, 접수대장 '사업자번호'에
 *  법인등록번호(13자리)가 섞여 들어갔고 법인등록번호는 따로 남기지 않았다.
 *  이 메뉴는 각 접수 폴더의 01.신청서에 보관된 원본 엑셀을 다시 읽어
 *  접수대장의 '사업자번호'·'법인등록번호' 두 열만 다시 채운다. 다른 열은 건드리지 않는다.
 *
 *  · 엑셀 변환이 건당 수 초 걸려 Apps Script 실행 제한(6분)에 걸릴 수 있으므로,
 *    시간 예산 안에서 처리한 만큼 기록하고 멈춘다. 다시 실행하면 남은 건부터 이어서 처리한다.
 *  · 처리한 접수번호는 문서 속성에 기록해 재실행 시 건너뛴다.
 *  · 사업자번호를 바꾼 건은 이전값 → 새값을 파싱로그에 남긴다.
 */

const 사업자번호복구_완료키 = 'BIZNO_BACKFILL_DONE';
const 사업자번호복구_시간예산ms = 4.5 * 60 * 1000;

function 사업자번호법인번호일괄복구() {
  const ui = SpreadsheetApp.getUi();
  if (!_관리자여부()) {
    ui.alert('관리자만 실행할 수 있습니다.');
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 대장 = ss.getSheetByName(SHEET.접수대장);
  const 헤더 = 대장.getRange(1, 1, 1, 대장.getLastColumn()).getValues()[0].map(v => String(v).trim());
  if (헤더.indexOf('법인등록번호') < 0) {
    ui.alert('접수대장에 법인등록번호 컬럼이 없습니다. 먼저 [누락 컬럼 오른쪽 끝 추가]를 실행하세요.');
    return;
  }
  const i번호 = 헤더.indexOf('접수번호');
  const 마지막행 = 대장.getLastRow();
  if (i번호 < 0 || 마지막행 < 2) {
    ui.alert('접수대장에 처리할 행이 없습니다.');
    return;
  }

  const 속성 = PropertiesService.getDocumentProperties();
  const 완료 = new Set(JSON.parse(속성.getProperty(사업자번호복구_완료키) || '[]'));
  const 접수번호목록 = 대장.getRange(2, i번호 + 1, 마지막행 - 1, 1).getDisplayValues()
    .map(r => String(r[0]).trim()).filter(Boolean);
  const 대상 = [...new Set(접수번호목록)].filter(no => !완료.has(no));
  if (!대상.length) {
    ui.alert('모든 건을 이미 처리했습니다.');
    return;
  }

  const 응답 = ui.alert(
    '사업자번호·법인등록번호 일괄 복구',
    `접수 폴더의 원본 엑셀(01.신청서)을 다시 읽어 접수대장의 '사업자번호'·'법인등록번호' 두 열만 채웁니다.\n` +
    `대상: ${대상.length}건 (이미 처리한 ${완료.size}건 제외)\n\n` +
    `· 다른 열은 바꾸지 않습니다.\n` +
    `· 약 ${Math.round(사업자번호복구_시간예산ms / 60000 * 10) / 10}분마다 멈춥니다. 남은 건이 있으면 다시 실행하세요.\n` +
    `· 사업자번호를 바꾼 건은 이전값과 함께 파싱로그에 남깁니다.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (응답 !== ui.Button.OK) return;

  const 시작 = Date.now();
  const 루트 = CONFIG.드라이브폴더ID
    ? DriveApp.getFolderById(CONFIG.드라이브폴더ID)
    : (it => (it.hasNext() ? it.next() : null))(DriveApp.getFoldersByName('심사증적_원본보관'));
  if (!루트) {
    ui.alert('접수 폴더 루트(심사증적_원본보관)를 찾지 못했습니다.');
    return;
  }

  const 결과맵 = {};   // 접수번호 → { 사업자번호, 법인등록번호 }
  const 실패 = [];
  let 처리수 = 0;
  for (const 접수번호 of 대상) {
    if (Date.now() - 시작 > 사업자번호복구_시간예산ms) break;
    처리수++;
    try {
      const 건 = _원본엑셀에서건찾기_(루트, 접수번호);
      if (!건) {
        실패.push(`${접수번호}: 원본 엑셀에서 같은 접수번호를 찾지 못함`);
      } else {
        결과맵[접수번호] = { 사업자번호: 건.사업자번호 || '', 법인등록번호: 건.법인등록번호 || '' };
      }
    } catch (e) {
      실패.push(`${접수번호}: ${e.message}`);
    }
    완료.add(접수번호);
  }

  const 변경 = _사업자번호복구결과기록_(ss, 결과맵);
  속성.setProperty(사업자번호복구_완료키, JSON.stringify([...완료]));

  const 로그 = ss.getSheetByName(SHEET.로그);
  const 시각 = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  if (로그) {
    const 행들 = 변경.사업자번호변경.map(([no, 전, 후]) =>
      [시각, '사업자번호 일괄 복구', no, '사업자번호 변경', `${전} → ${후}`]);
    실패.forEach(msg => 행들.push([시각, '사업자번호 일괄 복구', msg.split(':')[0], '실패', msg]));
    if (행들.length) 로그.getRange(로그.getLastRow() + 1, 1, 행들.length, 5).setValues(행들);
  }

  const 남은수 = 대상.length - 처리수;
  ui.alert(
    '사업자번호·법인등록번호 일괄 복구 결과',
    `이번 실행: ${처리수}건 확인\n` +
    `· 사업자번호 변경: ${변경.사업자번호변경.length}건\n` +
    `· 법인등록번호 채움: ${변경.법인등록번호채움}건\n` +
    `· 실패: ${실패.length}건${실패.length ? '\n  ' + 실패.slice(0, 10).join('\n  ') + (실패.length > 10 ? '\n  …(파싱로그 참고)' : '') : ''}\n\n` +
    (남은수 > 0 ? `남은 건 ${남은수}건 — 메뉴를 한 번 더 실행하세요.` : '모든 건 처리를 마쳤습니다.'),
    ui.ButtonSet.OK
  );
}

/** 접수 폴더 → 01.신청서 → 원본 엑셀(최신순)을 읽어 접수번호가 같은 건을 반환한다. */
function _원본엑셀에서건찾기_(루트, 접수번호) {
  const 검색어 = _폴더명값정리_(접수번호).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  const 폴더검색 = 루트.searchFolders(`title contains '(${검색어})'`);
  if (!폴더검색.hasNext()) throw new Error('접수 폴더 없음');
  const 신청서검색 = 폴더검색.next().getFoldersByName('01.신청서');
  if (!신청서검색.hasNext()) throw new Error('01.신청서 폴더 없음');

  const 엑셀목록 = [];
  const 파일들 = 신청서검색.next().getFiles();
  while (파일들.hasNext()) {
    const f = 파일들.next();
    const mime = f.getMimeType();
    if (mime === MimeType.GOOGLE_SHEETS || mime === MimeType.MICROSOFT_EXCEL ||
        /\.xlsx?$/i.test(f.getName())) {
      엑셀목록.push(f);
    }
  }
  if (!엑셀목록.length) throw new Error('01.신청서에 엑셀 파일 없음');
  엑셀목록.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());

  for (const 파일 of 엑셀목록) {
    const 건 = _엑셀파일에서건찾기_(파일, 접수번호);
    if (건) return 건;
  }
  return null;
}

function _엑셀파일에서건찾기_(파일, 접수번호) {
  let 임시ID = null;
  try {
    const 시트ID = 파일.getMimeType() === MimeType.GOOGLE_SHEETS
      ? 파일.getId()
      : (임시ID = _엑셀을Sheets로변환(파일.getBlob(), '__임시복구__' + Date.now()));
    const 시트들 = SpreadsheetApp.openById(시트ID).getSheets();
    // 기본정보 탭을 우선 보고, 없으면 나머지 탭도 확인한다.
    시트들.sort((a, b) => (/^기본정보$|basic/i.test(b.getName()) ? 1 : 0) - (/^기본정보$|basic/i.test(a.getName()) ? 1 : 0));
    for (const sh of 시트들) {
      // 전화번호·사업자번호의 선행 0을 보존하기 위해 표시 문자열로 읽는다(업로드와 동일).
      const 데이터 = sh.getDataRange().getDisplayValues();
      if (!데이터.length) continue;
      const 건목록 = _세로형감지(데이터) ? [_파싱_세로형(데이터)] : _파싱_가로형(데이터);
      const 건 = 건목록.find(g => String(g.접수번호 || '').trim() === 접수번호);
      if (건) return 건;
    }
    return null;
  } finally {
    if (임시ID) { try { DriveApp.getFileById(임시ID).setTrashed(true); } catch (e) {} }
  }
}

/**
 * 결과를 접수대장 두 열에만 한 번에 쓴다. 쓰는 시점에 접수번호 열을 다시 읽어
 * 행 위치를 맞추므로, 실행 중 행이 추가·삭제돼도 다른 건에 값이 들어가지 않는다.
 */
function _사업자번호복구결과기록_(ss, 결과맵) {
  const 결과 = { 사업자번호변경: [], 법인등록번호채움: 0 };
  if (!Object.keys(결과맵).length) return 결과;

  const 잠금 = LockService.getDocumentLock();
  if (!잠금.tryLock(30000)) throw new Error('다른 작업이 실행 중이어서 기록하지 못했습니다. 잠시 후 다시 실행하세요.');
  try {
    const 대장 = ss.getSheetByName(SHEET.접수대장);
    const 헤더 = 대장.getRange(1, 1, 1, 대장.getLastColumn()).getValues()[0].map(v => String(v).trim());
    const 열번호 = 헤더.indexOf('접수번호') + 1;
    const 열사업자 = 헤더.indexOf('사업자번호') + 1;
    const 열법인 = 헤더.indexOf('법인등록번호') + 1;
    const n = 대장.getLastRow() - 1;
    const 번호들 = 대장.getRange(2, 열번호, n, 1).getDisplayValues();
    const 사업자범위 = 대장.getRange(2, 열사업자, n, 1);
    const 법인범위 = 대장.getRange(2, 열법인, n, 1);
    const 사업자값 = 사업자범위.getDisplayValues();
    const 법인값 = 법인범위.getDisplayValues();

    번호들.forEach(([no], i) => {
      const r = 결과맵[String(no).trim()];
      if (!r) return;
      const 전 = String(사업자값[i][0]).trim();
      if (r.사업자번호 && r.사업자번호 !== 전) {
        결과.사업자번호변경.push([String(no).trim(), 전 || '(빈칸)', r.사업자번호]);
        사업자값[i][0] = r.사업자번호;
      }
      if (r.법인등록번호 && r.법인등록번호 !== String(법인값[i][0]).trim()) {
        법인값[i][0] = r.법인등록번호;
        결과.법인등록번호채움++;
      }
    });

    // 선행 0 보존을 위해 텍스트 서식으로 쓴다(표 열 타입 충돌 시 서식만 건너뜀).
    [사업자범위, 법인범위].forEach(범위 => { try { 범위.setNumberFormat('@'); } catch (e) {} });
    사업자범위.setValues(사업자값);
    법인범위.setValues(법인값);
    SpreadsheetApp.flush();
  } finally {
    잠금.releaseLock();
  }
  return 결과;
}
