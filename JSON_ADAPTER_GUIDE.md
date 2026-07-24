# JSON 신청서 어댑터 가이드

KOSA의 `인공지능_제품서비스_확인_신청서_JSON_Data_Schema_v1.0.pdf`에 정의된 JSON 신청서를
기존 XLSX 인수 파이프라인과 **동일한 정책**으로 시트에 등록하기 위한 문서입니다.

- **정의(데이터)**: [JsonDefinition.js](./JsonDefinition.js) — JSON 필드 경로·타입·시트컬럼·enum
  매핑·formatter를 선언만 함. 처리 로직은 전혀 없음.
- **처리 로직**: [JsonAdapter.js](./JsonAdapter.js) — 위 정의를 해석하는 범용 엔진, 등록 전
  사전 검증(보호된 범위 확인), 시트 등록 오케스트레이션, 메뉴 진입점.
- **대상 시트/등록 함수**: `Code.js`의 `_Sheets에등록` / `제품모델등록` / `AI기능상세등록` /
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
             ├─ _JSON구조검증(json)                         [JsonAdapter.js ⟶ JSON필수구조정의 참조]
             ├─ _JSON을건객체로변환(json)                   [JsonAdapter.js ⟶ JSON필드정의 참조]
             │     └─ _필드값추출() ⟶ _경로값() / _원시경로값() / _코드값라벨매핑() / JSON포매터.*
             ├─ _JSON에서제품모델목록추출(json)             ⟶ JSON제품모델필드정의 참조
             ├─ _JSON에서기능목록추출(json)                 ⟶ JSON기능상세필드정의 참조 (배열 인덱스 조인)
             ├─ _쓰기가능여부검사(ss, 시트별추가행수)        ⚠ 등록 전 사전 검증(A안) — 여기서 막히면 어떤 시트에도 안 씀
             ├─ _Sheets에등록(건, 파일명)                   ⟶ Code.js 재사용 (접수대장+일정관리+파싱로그)
             ├─ 제품모델등록(접수번호, 모델목록)             ⟶ Code.js 재사용
             └─ AI기능상세등록(접수번호, 기능목록)
                    → _접수대장기능수갱신()                  ⟶ Code.js 재사용
```

**파일별 책임**

| 파일 | 담당 |
|---|---|
| `Code.js` | 기존 진입점(`onOpen` 메뉴 한 줄)만 담당, 그 외 변경 없음 |
| `JsonDefinition.js` | JSON 경로(path)·필수여부(required)·타입(type)·시트컬럼(sheetColumn)·enum 코드값 매핑·formatter — **정의만**, 로직 없음 |
| `JsonAdapter.js` | 정의를 해석하는 범용 엔진(`_필드값추출` 등), 사전 검증, 시트 등록 호출, 메뉴 진입점, dev 테스트 |

Apps Script는 프로젝트 내 모든 `.js`/`.gs` 파일이 하나의 전역 스코프를 공유하므로, 세 파일이
import 없이 서로 참조합니다.

**배포 시 유의**: `.claspignore`는 기본적으로 화이트리스트 방식입니다. `Code.js`/`JsonAdapter.js`/
`JsonDefinition.js`/`appsscript.json` 네 파일만 push 대상으로 허용돼 있습니다 — 신규 파일을 또
추가할 경우 `.claspignore`도 함께 확인하세요.

---

## 2. 필드 정의 엔진 — JsonDefinition.js의 정의를 어떻게 해석하는가

`JsonAdapter.js`의 `_필드값추출(obj, 정의, index)`이 `JsonDefinition.js`의 각 정의 항목을 아래
규칙으로 해석합니다. 새 필드를 추가하거나 변환 규칙을 바꿀 때는 **이 표에 맞는 `type`을 골라
`JsonDefinition.js`에 항목만 추가**하면 되고, `JsonAdapter.js`는 건드릴 필요가 없습니다.

| type | 동작 | 사용 예 |
|---|---|---|
| (기본, 미지정) | `path`(dot 경로, 배열이면 fallback)로 값을 찾아 문자열로 반환 | 기업명, 개요 등 |
| `enum` | 값을 찾은 뒤 `코드값매핑표[enumCategory]`에서 라벨 조회. 없으면 원본값 유지 + 경고 로그(미등록 코드값 정책, 3번 참고) | 제공형태, 제품분류, 명세서작성방식 |
| `boolean` | `true`→`'동의'`, 그 외→`'미동의'` | 열람이용동의여부 등 3종 |
| `arrayLength` | `path`가 가리키는 배열의 길이. 배열이 아니면 `fallback` 값 | 제품수 (`productModels.length`) |
| `arrayPresence` | `path`가 가리키는 배열이 비어있지 않으면 `presentValue`, 비어있으면 `absentValue` | 보유인증 (`certification.certFiles`) |
| `joinField` | `path`가 가리키는 배열의 각 원소에서 `field` 값을 뽑아 `delimiter`로 합침 | AI기능명_원본 (`keyFunctions[].funcNm`) |
| `date` | 값이 있으면 그대로, 없고 `fallback:'now'`면 파싱 시각(`yyyy-MM-dd`) | 신청일 |
| `index1` | 값이 있으면 그대로, 없으면 배열 인덱스+1 | 기능번호 |
| `fixedValue` | `path` 없이 항상 고정값 (type 무관, 최우선 적용) | 첨부파일 관련 4개 필드(관리자 결정으로 제외), 비고 |
| `formatter` | 위 처리 후 `JSON포매터[formatter명]`을 추가 적용 | 사업자번호 (`사업자번호정규화`) |

`path`가 배열(`['applicant.managerMobile', 'applicant.managerTel']`)이면 앞에서부터 값이 있는
첫 후보를 채택합니다 — 연락처의 "휴대전화 우선, 없으면 일반전화" 규칙이 별도 함수 없이 이 방식
하나로 표현됩니다.

---

## 3. 관리자 결정 사항 (확정 반영됨)

| 항목 | 결정 내용 | 구현 위치 |
|---|---|---|
| **접수번호(receiptNo)** | JSON의 `keyId`(Base36)를 접수번호로 그대로 채택. 엑셀 경로는 기존 `AI2026...` 형식을 유지하고, JSON 경로는 Base36 원형을 그대로 사용 — 소스별로 서로 다른 형식이 공존함 | `JsonDefinition.js`의 `JSON필드정의` 중 `접수번호: { path: 'keyId' }` |
| **신청일(submittedAt)** | 현재 JSON에 날짜 필드가 없어 TTA 파싱 시각으로 대체. ⚠️ 실제 KOSA 신청일보다 항상 같거나 늦어지므로 법정 처리기한(15일) 판단 시 유의할 것 — 후보 필드 배열에 한 줄만 추가하면 향후 KOSA가 필드를 추가했을 때 즉시 반영 가능 | `JsonDefinition.js`의 `신청일: { path: [...], type: 'date', fallback: 'now' }` |
| **첨부파일(fileId ↔ Drive ID)** | JSON 경로에서는 첨부파일 처리를 제외(구조도파일명 등은 빈 값 유지). 기존 보고서 생성 코드(`_증적명세서Docs생성`)가 빈 값을 이미 "(미등록)" 문구로 정상 처리하므로 추가 예외처리 불필요 | `JsonDefinition.js`의 4개 필드 `fixedValue: ''` |
| **보유인증(AuthType)** | KOSA 스키마 자체가 `AuthType` 필드를 "매핑 필요" 상태로 미확정 상태로 남겨둠. `certification.certFiles` 존재 여부로만 단순화(`'보유(상세 확인 필요)'` / `'해당없음'`). 정책 확정 시 `JsonDefinition.js`의 이 필드 정의 한 항목만 교체하면 됨 | `JsonDefinition.js`의 `보유인증: { type: 'arrayPresence', ... }` |

---

## 4. 등록 전 사전 검증 (이슈 2, A안 — rollback 방식은 채택하지 않음)

실제 dev 테스트에서 접수대장·기능상세는 등록되고 일정관리·제품모델은 "보호된 셀" 오류로
실패해 시트 간 데이터가 어긋나는 문제가 발생했습니다. Rollback(사후 삭제) 대신, **등록을
시작하기 전에 4개 시트 모두 "이번에 쓸 행 전체 범위"가 편집 가능한지 먼저 확인**하고, 하나라도
막히면 아무 시트에도 쓰지 않는 방식을 채택했습니다.

- 구현: `JsonAdapter.js`의 `_쓰기가능여부검사(ss, 시트별추가행수)`
- Apps Script 기본 제공 API(`Sheet.getProtections()`, `Protection.canEdit()`,
  `Protection.getUnprotectedRanges()`)만으로 구현 — 별도 REST 호출 불필요
- 검사 대상: 접수대장/일정관리/인공지능제품모델/인공지능기능상세 4개 시트의 "다음에 추가될
  행 ~ 추가될 행 수만큼"(제품모델·기능상세는 여러 행이 될 수 있음)
- 시트 전체 보호(`ProtectionType.SHEET`)는 `getUnprotectedRanges()`로 예외 처리된 범위가 우리가
  쓸 범위를 완전히 덮는지까지 확인 — README에 명시된 "일정관리는 특이사항 컬럼만 예외로 열어둠"
  같은 컬럼 단위 보호 설정을 정확히 반영
- 하나라도 막히면 `Error`를 던져 `JSON파싱실행_내부`가 어떤 `_Sheets에등록` 호출도 하기 전에
  중단됨 — 오류 메시지에 **어느 시트의 어느 컬럼 때문인지** 명시

---

## 미등록 코드값 정책

- 매핑 테이블에 없는 코드값이 들어오면 자동화는 중단하지 않는다.
- 원본 코드값을 시트에 그대로 저장한다.
- Logger.log에 `[경고] 미등록 코드값` 형태로 기록한다.
- 운영자는 로그를 확인하여 매핑 테이블에 신규 코드를 추가한다.

적용 대상: `serviceInfo.serviceType`(제공형태), `serviceInfo.aiTechCategory`(제품분류),
`attachments.submitMethodWrite`(명세서작성방식) 3개 필드(`type: 'enum'`). 로그 형식 예시:

```
[경고] 미등록 코드값 — 필드: 제공형태, 값: "CLOUD" (매핑표에 없어 원본값을 그대로 저장합니다. JsonDefinition.js의 코드값매핑표.제공형태에 추가 필요)
```

배열 개수 불일치(`keyFunctions` vs `aiImplementations`) 역시 같은 철학으로, 등록을 막지 않고
경고 로그만 남긴 뒤 가능한 만큼(배열 인덱스 기준)만 조인해 계속 진행합니다.

이 정책은 **구조 검증 실패(`_JSON구조검증`)와는 다른 층위**입니다 — `applicant`/`serviceInfo`/
`agreements` 객체나 필수 배열 자체가 아예 없는 경우는 "건 객체를 만들 수조차 없는" 치명적
결함으로 간주해 등록을 중단시킵니다. 반면 코드값 미등록은 "값의 표현 형식을 모를 뿐 등록 자체는
가능한" 경우라 자동화를 계속 진행합니다.

---

## 5. TODO 목록 (JsonDefinition.js 보완 대상)

- [ ] `코드값매핑표.제공형태`에 `CLOUD`, `ETC` 등 KOSA 실사용 코드값 전체 목록 확보 후 추가
- [ ] `코드값매핑표.제품분류`에 `PRODUCT`, `ETC` 등 추가
- [ ] `코드값매핑표.명세서작성방식`에 `WRITE`/`FILE` 외 다른 코드값이 실제로 쓰이는지 확인
- [ ] `JSON필드정의`의 `신청일` 항목 `path` 배열에 KOSA가 향후 추가할 신청일 필드명 반영
- [ ] `JSON필드정의`의 `보유인증` 항목 — KOSA의 `AuthType` 필드 위치·허용값이 확정되면 `type`을
      `arrayPresence`에서 실제 값 매핑 방식으로 교체
- [ ] 첨부파일(KOSA fileId ↔ Google Drive ID) 처리 파이프라인 별도 설계 — 현재는 JSON 경로에서
      완전 제외 상태

---

## 6. 샘플 파일

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

## 7. 함수/정의 레퍼런스

**JsonDefinition.js (정의만)**

| 이름 | 역할 |
|---|---|
| `JSON필수구조정의` | 최상위 필수 객체/배열 목록 (`_JSON구조검증`이 순회) |
| `코드값매핑표` | enum 코드값 → 한글 라벨 |
| `JSON포매터` | 범용 type으로 표현 안 되는 개별 포매터 (현재는 `사업자번호정규화`만) |
| `JSON필드정의` | "건" 객체(접수대장/일정관리 공용) 필드 정의 |
| `JSON제품모델필드정의` | 제품모델 반복행 필드 정의 |
| `JSON기능상세필드정의` | 기능상세 반복행 필드 정의 (`f.*`/`impl.*` 컨텍스트 참조) |

**JsonAdapter.js (처리 로직)**

| 함수 | 역할 |
|---|---|
| `JSON파싱등록()` | 메뉴 진입점. Drive 파일 ID/URL을 입력받아 JSON 파싱 후 등록 실행 |
| `JSON파싱실행_내부(json, 파일명)` | 오케스트레이터 — 검증→변환→**사전 검증**→등록 순서로 실행 |
| `_JSON구조검증(json)` | `JSON필수구조정의` 순회, 없으면 등록 중단 |
| `_쓰기가능여부검사(ss, 시트별추가행수)` | 등록 전 4개 시트 보호 범위 확인 (A안) |
| `_코드값라벨매핑(카테고리, 원본코드, 로그용필드명)` | `코드값매핑표` 조회, 미등록 시 원본 유지 + 경고 로그 |
| `_경로값(obj, path)` / `_원시경로값(obj, path)` | dot 경로 조회 (fallback 배열 지원 / 원시값 조회) |
| `_필드값추출(obj, 정의, index)` | 정의 하나를 해석해 실제 값을 계산하는 범용 엔진 (2번 표 참고) |
| `_JSON을건객체로변환(json)` | `JSON필드정의` 순회 → "건" 객체 생성 |
| `_JSON에서제품모델목록추출(json)` | `JSON제품모델필드정의` 순회 → 제품모델 등록용 배열 |
| `_JSON에서기능목록추출(json)` | `JSON기능상세필드정의` 순회, `keyFunctions[]`/`aiImplementations[]` 인덱스 조인 |
| `testJsonImport()` | dev 전용 테스트 하네스 (실제 시트에 기록됨, dev 프로젝트에서만 실행) |

**Code.js (기존, 수정 없이 재사용)**: `_Sheets에등록` / `제품모델등록` / `AI기능상세등록` /
`_접수대장기능수갱신` / `_드라이브ID추출`
