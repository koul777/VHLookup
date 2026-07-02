# VHLookup Local

공공기관, 교육청, 학교, HR, 총무, 예산 담당자가 반복하는 엑셀 수합, 대조, 검증, 요약 업무를 로컬 PC에서 처리하는 Windows용 업무 자동화 도구입니다.

비개발자 사용자는 Python이나 개발 환경을 몰라도 됩니다. 배포본을 받은 경우 아래 exe 파일만 더블클릭해서 사용합니다.

```text
dist\VHLookupLocal_pivot.exe
```

GitHub 소스 저장소에는 빌드 산출물인 `dist` 폴더와 exe 파일을 포함하지 않습니다. 소스에서 직접 실행 파일을 만들려면 `build_exe.bat`을 실행합니다.

## 핵심 원칙

- 로컬 PC에서만 실행
- 원본 엑셀/CSV 파일 미수정
- 별도 서버, API, 클라우드 업로드 없음
- 결과 파일은 `outputs` 폴더에 새 엑셀 파일로 저장
- 컬럼명이 달라도 자동 추천 후 사용자가 드롭박스로 수정 가능
- 개인정보 의심 컬럼은 값이 아니라 유형과 건수 중심으로 점검

## 주요 기능

### 1. 원본 문제 표시 보고서

제출받은 엑셀/CSV 파일을 실제로 합치거나 정리하기 전에, 파일 상태만 확인하는 보고서를 만듭니다.

- 원본 파일은 수정하지 않음
- 여러 파일을 한 번에 넣어서 파일별 구조 차이 확인
- 실제 헤더가 몇 번째 줄인지 자동 탐지
- 파일마다 컬럼명이 어떻게 다른지 비교
- 필수값 누락, 숫자 오류, 날짜 오류, 중복 의심 확인
- 성명, 연락처, 주민번호 패턴 같은 개인정보 의심 위치 확인
- 결과 앞쪽의 `원본확인_*` 시트에 문제 위치를 색과 메모로 표시
- `먼저확인` 시트에서 전체 문제 요약과 바로 볼 시트 안내

이 기능은 "파일을 고치는 기능"이 아니라 "어떤 파일이 이상한지 먼저 표시하는 기능"입니다. 제출자료를 수합하기 전에 양식이 틀어진 파일, 개인정보가 섞인 파일, 값이 빠진 파일을 먼저 찾을 때 사용합니다.

### 2. 분류별 시트 나누기

한 파일을 특정 열 기준으로 여러 시트로 나눕니다.

- 부서, 기관명, 학교명, 상태, Department 같은 열 기준
- 기준 열은 드롭박스로 선택
- 원본 전체 시트와 분류별 시트 생성

### 3. 엑셀/CSV 파일 여러 개 합치기

여러 파일을 한 결과 파일로 합칩니다.

- `행 합치기`: 같은 양식의 여러 파일을 아래로 이어 붙임
- `열 합치기`: 공통 키를 찾아 다른 파일의 열을 오른쪽에 붙임
- 컬럼명이 달라도 자동 매칭
- 자동 매칭이 틀리면 `컬럼 매칭 수정`에서 드롭박스로 직접 지정
- 미리보기에서 앞 10행 확인 후 실행

### 4. 전/후 파일 검증

수정 전 파일과 수정 후 파일을 비교합니다.

- 같은 의미의 컬럼을 자동 매칭
- 비교 기준열 자동 추천
- 자동 추천이 틀리면 `컬럼 매칭 수정`에서 비교 기준열과 전/후 컬럼 직접 지정
- 값이 바뀐 셀, 추가 행, 누락 행, 추가/삭제 컬럼 확인
- 결과 파일의 `후파일_메모` 시트에서 변경 셀 메모 확인

### 5. 피벗 요약표 만들기

부서별, 기관별, 월별, 상태별 요약표를 만듭니다.

- 행 기준 선택
- 열 기준 선택 또는 선택 안 함
- 값 열 선택
- 집계 방식 선택: `건수`, `합계`, `평균`, `최대`, `최소`
- `건수`는 값 열 없이 행 개수 집계
- `합계`, `평균`, `최대`, `최소`는 숫자 값 열 기준 집계
- 결과 시트: `먼저확인`, `피벗요약`, `상위목록`, `기준설명`, `원본`, 필요 시 `오류행만`

## 사용 순서

1. `dist\VHLookupLocal_pivot.exe`를 더블클릭합니다.
2. 실행할 작업 버튼을 선택합니다.
3. 파일을 올립니다.
4. 미리보기에서 예상 결과를 확인합니다.
5. 자동 추천이 틀리면 드롭박스에서 컬럼 매칭이나 기준 열을 수정합니다.
6. `실행`을 누릅니다.
7. 결과는 `outputs` 폴더에 저장되고 완료 후 폴더가 열립니다.

## 샘플 데이터

샘플은 프로그램 안에서 버튼으로 만들지 않고 파일로 제공합니다.

```text
samples\public_admin
```

샘플 CSV는 Windows Excel에서 더블클릭으로 열어도 한글이 깨지지 않도록 UTF-8 BOM 형식으로 저장했습니다.

### 샘플별 체험 방법

좁은 화면에서도 보기 쉽도록 샘플별로 나누어 적었습니다.

**제출자료 파일 구조 점검**

- 누를 기능: `1. 원본 문제 표시 보고서`
- 선택할 파일:
  - `samples\public_admin\submission_errors\bad_school_a.csv`
  - `samples\public_admin\submission_errors\bad_school_b.csv`
- 화면 선택값: 파일 2개 선택
- 결과에서 먼저 볼 시트: `원본확인_bad_school_a`, `원본확인_bad_school_b`
- 그 다음 참고 시트: `먼저확인`, `파일점검`, `열비교`, `확인필요`, `개인정보점검`

**부서별 시트 나누기**

- 누를 기능: `2. 분류별 시트 나누기`
- 선택할 파일: `samples\public_admin\pivot_summary\budget_execution.csv`
- 화면 선택값: 분류 기준열 `부서`
- 결과에서 먼저 볼 시트: `총무과`, `예산과`, `복지과`, `민원과` 같은 분류별 시트
- 그 다음 참고 시트: `전체`, `먼저확인`

**제출자료 행 합치기**

- 누를 기능: `3. 엑셀/CSV 파일 여러 개 합치기`
- 선택할 파일:
  - `samples\public_admin\school_submissions\gangbuk_school.csv`
  - `samples\public_admin\school_submissions\gangnam_school.csv`
- 화면 선택값: 합치기 방향 `행 합치기`
- 결과에서 먼저 볼 시트: `결과`
- 그 다음 참고 시트: `먼저확인`, `자동추천근거`, `개인정보점검`

**제목/안내문 있는 파일 행 합치기**

- 누를 기능: `3. 엑셀/CSV 파일 여러 개 합치기`
- 선택할 파일:
  - `samples\public_admin\messy_headers\department_status_a.csv`
  - `samples\public_admin\messy_headers\department_status_b.csv`
- 화면 선택값: 합치기 방향 `행 합치기`
- 결과에서 먼저 볼 시트: `결과`
- 그 다음 참고 시트: `먼저확인`, `자동추천근거`

**직원 명단 열 합치기**

- 누를 기능: `3. 엑셀/CSV 파일 여러 개 합치기`
- 선택할 파일:
  - `samples\public_admin\hr_training_completion.csv`
  - `samples\public_admin\hr_employee_master.csv`
- 화면 선택값: 합치기 방향 `열 합치기`, 자동 매칭 확인
- 결과에서 먼저 볼 시트: `결과`
- 그 다음 참고 시트: `먼저확인`, `자동추천근거`, `확인필요`

**수당/예산 기준표 열 합치기**

- 누를 기능: `3. 엑셀/CSV 파일 여러 개 합치기`
- 선택할 파일:
  - `samples\public_admin\allowance_budget\payment_requests.csv`
  - `samples\public_admin\allowance_budget\rate_reference.csv`
- 화면 선택값: 합치기 방향 `열 합치기`, 자동 매칭 확인
- 결과에서 먼저 볼 시트: `결과`
- 그 다음 참고 시트: `먼저확인`, `자동추천근거`, `확인필요`

**전/후 파일 변경 검증**

- 누를 기능: `4. 전/후 파일 검증`
- 선택할 파일:
  - 수정 전: `samples\public_admin\before_after_validation\payment_before.csv`
  - 수정 후: `samples\public_admin\before_after_validation\payment_after.csv`
- 화면 선택값: 비교 기준열 자동 추천 확인, 필요 시 `컬럼 매칭 수정`
- 결과에서 먼저 볼 시트: `후파일_메모`
- 그 다음 참고 시트: `먼저확인`, `차이목록`, `행비교`, `컬럼비교`, `차이행만`

**부서/월별 예산 피벗**

- 누를 기능: `5. 피벗 요약표 만들기`
- 선택할 파일: `samples\public_admin\pivot_summary\budget_execution.csv`
- 화면 선택값: 행 기준 `부서`, 열 기준 `월`, 값 열 `금액`, 집계 방식 `합계`
- 결과에서 먼저 볼 시트: `피벗요약`
- 그 다음 참고 시트: `먼저확인`, `상위목록`, `기준설명`, `원본`

**상태별 처리 건수 피벗**

- 누를 기능: `5. 피벗 요약표 만들기`
- 선택할 파일: `samples\public_admin\pivot_summary\budget_execution.csv`
- 화면 선택값: 행 기준 `상태`, 열 기준 `(선택 안 함)`, 값 열 `(행 개수)`, 집계 방식 `건수`
- 결과에서 먼저 볼 시트: `피벗요약`
- 그 다음 참고 시트: `먼저확인`, `상위목록`, `기준설명`, `원본`

월별 가로표를 세로형 자료로 바꾸는 기능은 현재 exe 첫 화면의 1~5번 버튼에는 넣지 않았고, CLI의 `horizontal` 명령으로만 제공합니다.

샘플 종류와 확인 포인트는 [docs/sample_catalog.md](docs/sample_catalog.md)에 더 자세히 정리되어 있습니다.

## 결과 파일

결과 파일은 기본적으로 아래 폴더에 생성됩니다.

```text
dist\outputs
```

개발 환경에서 직접 실행하면 프로젝트 루트의 `outputs` 폴더에 저장될 수 있습니다.

자주 보는 시트:

- `먼저확인`: 결과를 열었을 때 먼저 봐야 할 요약
- `결과`: 수합, 대조 결과
- `확인필요`: 오류, 경고, 수동 확인 항목
- `자동추천근거`: 어떤 컬럼을 왜 맞췄는지
- `개인정보점검`: 개인정보 의심 컬럼 유형과 건수
- `오류행만`: 문제가 있는 행만 따로 모은 시트
- `피벗요약`: 피벗 요약표 결과
- `후파일_메모`: 전/후 파일 검증에서 변경 셀 메모가 붙은 후 파일

## exe 다시 만들기

배포용 exe를 다시 만들 때는 아래 파일을 실행합니다.

```text
build_exe.bat
```

빌드가 끝나면 아래 파일이 생성됩니다.

```text
dist\VHLookupLocal_pivot.exe
```

## 테스트

개발자가 기능 검증을 할 때 사용합니다.

```powershell
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD='1'; python -m pytest
```

현재 주요 테스트는 다음 업무 흐름을 검증합니다.

- 분류별 시트 나누기
- 행/열 합치기
- 컬럼 자동 매칭
- 전/후 파일 검증
- 피벗 요약표
- 공공기관 행정 샘플 처리

## CLI 사용

데스크톱 앱 없이 명령어로도 일부 기능을 실행할 수 있습니다.

```powershell
python -m vhlookup_cli.main inspect --path "C:\path\submissions" --out inspect.xlsx
python -m vhlookup_cli.main consolidate --folder "C:\path\submissions" --out result.xlsx
python -m vhlookup_cli.main lookup --reference master.xlsx --target submitted.xlsx --out lookup_result.xlsx
python -m vhlookup_cli.main reconcile --reference expected.xlsx --target received.xlsx --out missing_result.xlsx
python -m vhlookup_cli.main horizontal --file monthly.xlsx --out monthly_long.xlsx
```

비개발자용 기본 사용은 CLI가 아니라 `dist\VHLookupLocal_pivot.exe` 실행입니다.

## 문서

- [실행안내.md](실행안내.md): 비개발자용 실행 순서
- [docs/sample_catalog.md](docs/sample_catalog.md): 샘플 데이터 카탈로그
- [docs/security_review_brief.md](docs/security_review_brief.md): 보안 검토 설명
- [docs/threads_content_pack.md](docs/threads_content_pack.md): Threads 홍보 문구 초안

## 구조

- `src/vhlookup_core`: 파일 로더, 헤더 탐지, 컬럼 매칭, 수합, 대조, 피벗, 리포트 작성
- `src/vhlookup_app`: Windows 데스크톱 GUI
- `src/vhlookup_cli`: 명령어 실행 도구
- `samples`: 샘플 데이터
- `tests`: 자동 테스트
