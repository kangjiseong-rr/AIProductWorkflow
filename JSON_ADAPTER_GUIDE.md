# JSON 신청서 어댑터 가이드

KOSA의 `인공지능_제품서비스_확인_신청서_JSON_Data_Schema_v1.0.pdf`에 정의된 JSON 신청서를
기존 XLSX 인수 파이프라인과 **동일한 정책**으로 시트에 등록하기 위한 변환 규칙 문서입니다.

- 실제 변환 코드: [JsonAdapter.js](./JsonAdapter.js)
- 대상 시트/등록 함수: `Code.js`의 `_Sheets에등록` / `제품모델등록` / `AI기능상세등록` /
  `_접수대장기능수갱신` (수정 없이 그대로 재사용)
- Code.js에 대한 변경은 `onOpen()` 메뉴에 **"📥 JSON 파일 등록 (파싱)"** 항목 한 줄을 추가한 것뿐입니다.

---

## 1. 아키텍처 개요

```
onOpen() [Code.js, 1줄 추가]
 └─ 메뉴 클릭 → JSON파싱등록() [JsonAdapter.js]
       ├─ _드라이브ID추출()                       ⟶ Code.js 재사용
       ├─ JSON.parse(파일)
       └─ JSON파싱실행_내부(json, 파일명)
             ├─ _JSON구조검증()
             ├─ _접수번호결정() / _신청일결정()
             ├─ _JSON을건객체로변환()
             │     ├─ _Boolean동의매핑()
             │     ├─ _코드값라벨매핑()            ⚠ 미등록 코드값 정책 적용 지점
             │     ├─ _연락처병합()
             │     ├─ _사업자번호정규화()
             │     └─ _보유인증단순화()
             ├─ _Sheets에등록(건, 파일명)           ⟶ Code.js 재사용 (접수대장+일정관리+파싱로그)
             ├─ _JSON에서제품모델목록추출() → 제품모델등록()        ⟶ Code.js 재사용
             └─ _JSON에서기능목록추출()   → AI기능상세등록()
                                          → _접수대장기능수갱신()   ⟶ Code.js 재사용
```

Apps Script는 프로젝트 내 모든 `.js`/`.gs` 파일이 하나의 전역 스코프를 공유하므로,
`JsonAdapter.js`에서 `Code.js`의 함수를 import 없이 그대로 호출합니다.

**배포 시 유의**: `.claspignore`는 기본적으로 `Code.js`와 `appsscript.json`만 push 대상으로
허용하는 화이트리스트 방식이었습니다. `JsonAdapter.js`도 push되도록 `!JsonAdapter.js` 줄을
추가해두었습니다 — 신규 파일을 또 추가할 경우 `.claspignore`도 함께 확인하세요.

---

## 2. 변환 규칙 (자동 변환 — 질문 없이 처리)

| XLSX 필드 | JSON 경로 | 변환 규칙 | 구현 위치 |
|---|---|---|---|
| 열람이용동의여부 등 3종 | `agreements.*` (Boolean) | `true→'동의'`, 그 외(`false`/누락)→`'미동의'` | `_Boolean동의매핑()` |
| 연락처 | `applicant.managerMobile` / `managerTel` | 휴대전화 우선, 없으면 일반전화 | `_연락처병합()` |
| 사업자번호 | `applicant.businessNo` | 숫자만 추출해 10자리면 `000-00-00000` 재포맷, 아니면 원본 유지 | `_사업자번호정규화()` |
| 제공형태 / 제품분류 / 명세서작성방식 | `serviceInfo.serviceType` / `aiTechCategory`, `attachments.submitMethodWrite` | 코드값 → 한글 라벨 매핑표 조회 (3번 참조) | `_코드값라벨매핑()` |
| 제품수 / 인공지능기능수 / 구현방식(요약) 등 | `productModels[]` / `keyFunctions[]` / `aiImplementations[]` | 배열 길이·내용 기반 자동 계산 — 기존 `제품모델등록`/`_접수대장기능수갱신` 로직 그대로 재사용 | `Code.js` 기존 함수 |

---

## 3. 관리자 결정 사항 (확정 반영됨)

| 항목 | 결정 내용 | 구현 위치 |
|---|---|---|
| **접수번호(receiptNo)** | JSON의 `keyId`(Base36)를 접수번호로 그대로 채택. 엑셀 경로는 기존 `AI2026...` 형식을 유지하고, JSON 경로는 Base36 원형을 그대로 사용 — 소스별로 서로 다른 형식이 공존함 | `_접수번호결정()` |
| **신청일(submittedAt)** | 현재 JSON에 날짜 필드가 없어 TTA 파싱 시각으로 대체. ⚠️ 실제 KOSA 신청일보다 항상 같거나 늦어지므로 법정 처리기한(15일) 판단 시 유의할 것 — 후보 필드 목록에 한 줄만 추가하면 향후 KOSA가 필드를 추가했을 때 즉시 반영 가능하도록 설계 | `_신청일결정()` |
| **첨부파일(fileId ↔ Drive ID)** | JSON 경로에서는 첨부파일 처리를 제외(구조도파일명 등은 빈 값 유지). 기존 보고서 생성 코드(`_증적명세서Docs생성`)가 빈 값을 이미 "(미등록)" 문구로 정상 처리하므로 추가 예외처리 불필요 | `_JSON을건객체로변환()` |
| **보유인증(AuthType)** | KOSA 스키마 자체가 `AuthType` 필드를 "매핑 필요" 상태로 미확정 상태로 남겨둠. `certification.certFiles` 존재 여부로만 단순화(`'보유(상세 확인 필요)'` / `'해당없음'`). 정책 확정 시 이 함수 내부만 교체하면 됨 | `_보유인증단순화()` |

---

## 미등록 코드값 정책

- 매핑 테이블에 없는 코드값이 들어오면 자동화는 중단하지 않는다.
- 원본 코드값을 시트에 그대로 저장한다.
- Logger.log에 `[경고] 미등록 코드값` 형태로 기록한다.
- 운영자는 로그를 확인하여 매핑 테이블에 신규 코드를 추가한다.

적용 대상: `serviceInfo.serviceType`(제공형태), `serviceInfo.aiTechCategory`(제품분류),
`attachments.submitMethodWrite`(명세서작성방식) 3개 필드. 로그 형식 예시:

```
[경고] 미등록 코드값 — 필드: serviceInfo.serviceType, 값: "CLOUD" (매핑표에 없어 원본값을 그대로 저장합니다. 코드값매핑표.제공형태에 추가 필요)
```

배열 개수 불일치(`keyFunctions` vs `aiImplementations`) 역시 같은 철학으로, 등록을 막지 않고
경고 로그만 남긴 뒤 가능한 만큼(배열 인덱스 기준)만 조인해 계속 진행합니다.

이 정책은 **구조 검증 실패(`_JSON구조검증`)와는 다른 층위**입니다 — `applicant`/`serviceInfo`/
`agreements` 객체나 필수 배열 자체가 아예 없는 경우는 "건 객체를 만들 수조차 없는" 치명적
결함으로 간주해 등록을 중단시킵니다. 반면 코드값 미등록은 "값의 표현 형식을 모를 뿐 등록 자체는
가능한" 경우라 자동화를 계속 진행합니다.

---

## 4. TODO 목록 (코드값 매핑표 보완 대상)

- [ ] `코드값매핑표.제공형태`에 `CLOUD`, `ETC` 등 KOSA 실사용 코드값 전체 목록 확보 후 추가
- [ ] `코드값매핑표.제품분류`에 `PRODUCT`, `ETC` 등 추가
- [ ] `코드값매핑표.명세서작성방식`에 `WRITE`/`FILE` 외 다른 코드값이 실제로 쓰이는지 확인
- [ ] `_신청일결정()`의 후보 필드 목록에 KOSA가 향후 추가할 신청일 필드명 반영
- [ ] `_보유인증단순화()` — KOSA의 `AuthType` 필드 위치·허용값이 확정되면 로직 교체
- [ ] 첨부파일(KOSA fileId ↔ Google Drive ID) 처리 파이프라인 별도 설계 — 현재는 JSON 경로에서 완전 제외 상태

---

## 5. 샘플 파일

| 파일 | 용도 |
|---|---|
| [샘플_신청데이터_JSON_1건.json](./샘플_신청데이터_JSON_1건.json) | 정상 케이스 — PDF 스키마 문서의 예시 JSON을 그대로 추출한 파일. `keyFunctions` 2건 / `aiImplementations` 1건으로 원본 예시 자체에 배열 개수 불일치가 이미 존재해, 정상 케이스인 동시에 배열 조인 경고 로그도 함께 검증됨 |
| [샘플_신청데이터_JSON_결함_필수객체누락.json](./샘플_신청데이터_JSON_결함_필수객체누락.json) | `agreements` 객체 자체가 없음 → `_JSON구조검증()`에서 등록이 중단되어야 하는 케이스 |
| [샘플_신청데이터_JSON_결함_미등록코드값.json](./샘플_신청데이터_JSON_결함_미등록코드값.json) | `serviceType=CLOUD`, `aiTechCategory=PRODUCT`, `submitMethodWrite=SIGN` — 매핑표에 없는 코드값 3종이 동시에 들어와도 등록이 계속 진행되고 원본값이 그대로 저장 + 경고 로그가 남는지 검증 |
| [샘플_신청데이터_JSON_결함_배열개수불일치.json](./샘플_신청데이터_JSON_결함_배열개수불일치.json) | `keyFunctions`(3건) vs `aiImplementations`(1건) 불일치 + `businessNo` 8자리(비정상 자릿수) — 배열 조인 경고 로그와 사업자번호 정규화 폴백(원본 유지) 동시 검증 |

각 파일 상단의 `_결함케이스설명` 키는 테스트 목적 설명용이며 등록 로직에서는 사용하지 않고 무시됩니다.

테스트 실행 방법은 추후 별도로 정합니다 (`testJsonImport()`가 `JsonAdapter.js`에 이미 있으나,
실행 절차는 다음 논의에서 확정).

---

## 6. 함수 레퍼런스

| 함수 | 파일 | 역할 |
|---|---|---|
| `JSON파싱등록()` | JsonAdapter.js | 메뉴 진입점. Drive 파일 ID/URL을 입력받아 JSON 파싱 후 등록 실행 |
| `JSON파싱실행_내부(json, 파일명)` | JsonAdapter.js | 오케스트레이터 — `_엑셀파싱처리`와 동일한 순서로 등록 함수 호출 |
| `_JSON구조검증(json)` | JsonAdapter.js | 필수 객체/배열 존재 여부 확인, 없으면 등록 중단 |
| `_접수번호결정(json)` | JsonAdapter.js | `keyId`를 접수번호로 채택 |
| `_신청일결정(json)` | JsonAdapter.js | 신청일 후보 필드 우선순위 조회, 없으면 파싱 시각 대체 |
| `_JSON을건객체로변환(json, 접수번호)` | JsonAdapter.js | `_파싱_세로형()`과 동일한 키 구조의 "건" 객체 생성 |
| `_코드값라벨매핑(카테고리, 원본코드, 로그용필드명)` | JsonAdapter.js | 코드값 → 한글 라벨. 미등록 시 원본 유지 + 경고 로그 |
| `_Boolean동의매핑(값)` / `_연락처병합(applicant)` / `_사업자번호정규화(원본)` | JsonAdapter.js | 단순 자동 변환 |
| `_보유인증단순화(json)` | JsonAdapter.js | `certFiles` 존재 여부 기반 단순화 |
| `_JSON에서제품모델목록추출(json)` | JsonAdapter.js | `productModels[]` → 제품모델 등록용 배열 |
| `_JSON에서기능목록추출(json)` | JsonAdapter.js | `keyFunctions[]` + `aiImplementations[]` 인덱스 조인 |
| `testJsonImport()` | JsonAdapter.js | dev 전용 테스트 하네스 (실제 시트에 기록됨, dev 프로젝트에서만 실행) |
| `_Sheets에등록` / `제품모델등록` / `AI기능상세등록` / `_접수대장기능수갱신` / `_드라이브ID추출` | Code.js (기존) | 수정 없이 재사용 |