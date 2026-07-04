# VHLookup Local

공공기관, 교육청, 학교, HR, 총무, 예산 담당자가 반복하는 엑셀 수합, 대조, 검증, 요약 업무를 로컬 PC에서 처리하는 Windows용 업무 자동화 도구입니다.

비개발자 사용자는 Python이나 개발 환경을 몰라도 됩니다. 실행 파일은 소스 저장소의 파일 목록이 아니라 GitHub Release에서 내려받습니다.

- [최신 Windows 실행파일 다운로드](https://github.com/koul777/VHLookup/releases/latest/download/VHLookupLocal_pivot.exe)
- 직접 다운로드가 안 되면 [GitHub Releases](https://github.com/koul777/VHLookup/releases/latest)에서 `VHLookupLocal_pivot.exe`를 내려받습니다.

내려받은 `VHLookupLocal_pivot.exe`를 더블클릭해서 사용합니다.

GitHub 소스 저장소에는 빌드 산출물인 `dist` 폴더와 exe 파일을 포함하지 않습니다. 소스 ZIP을 받거나 저장소를 clone한 경우 `dist\VHLookupLocal_pivot.exe`는 없는 것이 정상입니다.

배포 담당자가 아직 Release에 exe 파일을 올리지 않았다면 위 다운로드 링크에서도 파일을 받을 수 없습니다. 이 경우 배포 담당자가 `build_exe.bat`으로 만든 `dist\VHLookupLocal_pivot.exe`를 GitHub Release 첨부 파일로 올려야 합니다.

소스에서 직접 실행 파일을 만들려면 `build_exe.bat`을 실행합니다. 빌드가 끝나면 아래 위치에 exe가 생성됩니다.

```text
dist\VHLookupLocal_pivot.exe
```

## 핵심 원칙

- 로컬 PC에서만 실행
- 원본 엑셀/CSV 파일 미수정
- 별도 서버, API, 클라우드 업로드 없음
- 결과 파일은 `outputs` 폴더에 새 엑셀 파일로 저장
- 컬럼명이 달라도 자동 추천 후 사용자가 드롭박스로 수정 가능
- 개인정보 의심 컬럼은 값이 아니라 유형과 건수 중심으로 점검

## 주요 기능

### 1. 개인정보 마스킹

엑셀/CSV 파일에서 개인정보로 보이는 값을 가린 새 결과 파일을 만듭니다.

- 주민등록번호: 앞 6자리만 남기고 뒤 7자리 마스킹
- 이름/담당자/예금주: 가운데 글자 마스킹
- 연락처: 중간 번호 마스킹
- 이메일: 아이디 일부 마스킹
- 계좌번호: 가운데 숫자 마스킹
- 주소: 상세 주소 일부 마스킹
- 마스킹된 셀은 색으로 표시하고, 넓은 메모에 마스킹 유형 표시
- `마스킹내역` 시트에는 원본 값 없이 행 번호, 컬럼명, 마스킹 유형만 기록

### 2. 분류별 시트 나누기

한 파일을 특정 열 기준으로 여러 시트로 나눕니다.

- 부서, 기관명, 학교명, 상태, Department 같은 열 기준
- 기준 열은 드롭박스로 선택
- 원본 전체 시트와 분류별 시트 생성

### 3. 엑셀/CSV 파일 여러 개 합치기

여러 파일을 한 결과 파일로 합칩니다.

- `행 합치기`: 같은 양식의 여러 파일을 아래로 이어 붙임
- `열 합치기`: 공통 키를 찾아 다른 파일의 열을 오른쪽에 붙임
- 컬럼명이 달라도 동의어, 이름 유사도, 실제 값 겹침으로 자동 매칭
- `사번`과 `직원번호`, `관리번호`와 `접수ID`처럼 이름이 달라도 값이 충분히 겹치면 키 후보로 추천
- `00123`과 `123`처럼 앞자리 0 표시가 다른 숫자형 키도 같은 대상으로 비교
- 월별 지급자료처럼 단일 열만으로 중복되는 경우 `사번 + 지급월` 같은 복합 키 자동 추천
- 자동 매칭이 틀리면 `컬럼 매칭 수정`에서 드롭박스로 직접 지정
- 미리보기에서 앞 10행 확인 후 실행

### 4. 전/후 파일 검증

수정 전 파일과 수정 후 파일을 비교합니다.

- 같은 의미의 컬럼을 자동 매칭
- 비교 기준열 자동 추천
- 컬럼명이 바뀌어도 실제 값 겹침이 높으면 같은 컬럼으로 비교
- `001`과 `1`처럼 앞자리 0 표시가 다른 기준값도 같은 행으로 비교
- 자동 추천이 틀리면 `컬럼 매칭 수정`에서 비교 기준열과 전/후 컬럼 직접 지정
- 값이 바뀐 셀, 추가 행, 누락 행, 추가/삭제 컬럼 확인
- 결과 파일의 `후파일_메모` 시트에서 변경 셀 메모 확인
- 후 파일에서 빠진 행은 `빠진행` 시트에서 전 파일의 원본 행 전체 내용 확인

### 5. 피벗 요약표 만들기

부서별, 기관별, 월별, 상태별 요약표를 만듭니다.

- 행 기준 선택
- 열 기준 선택 또는 선택 안 함
- 값 열 선택
- 집계 방식 선택: `건수`, `합계`, `평균`, `최대`, `최소`
- `건수`는 값 열 없이 행 개수 집계
- `합계`, `평균`, `최대`, `최소`는 숫자 값 열 기준 집계
- 결과 숫자는 천 단위 콤마와 최대 소수 2자리로 보기 좋게 표시
- 결과 시트: `먼저확인`, `피벗요약`, `상위목록`, `기준설명`, `원본`, 필요 시 `오류행만`

## 사용 순서

1. 다운로드한 `VHLookupLocal_pivot.exe`를 더블클릭합니다. 직접 빌드한 경우에는 `dist\VHLookupLocal_pivot.exe`를 실행합니다.
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

**개인정보 마스킹**

- 누를 기능: `1. 개인정보 마스킹`
- 선택할 파일: `samples\public_admin\privacy_masking\citizen_service_requests.csv`
- 결과에서 먼저 볼 시트: `마스킹결과`
- 그 다음 참고 시트: `먼저확인`, `마스킹내역`
- 확인 포인트: 이름, 주민등록번호, 연락처, 이메일, 계좌번호, 주소가 가려졌는지 확인

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
- 그 다음 참고 시트: `먼저확인`, `차이목록`, `행비교`, `컬럼비교`, `차이행만`, `빠진행`

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
- `마스킹결과`: 개인정보가 가려진 결과표
- `마스킹내역`: 원본 값 없이 마스킹 위치와 유형만 기록한 시트
- `결과`: 수합, 대조 결과
- `확인필요`: 오류, 경고, 수동 확인 항목
- `자동추천근거`: 어떤 컬럼을 왜 맞췄는지
- `개인정보점검`: 개인정보 의심 컬럼 유형과 건수
- `오류행만`: 문제가 있는 행만 따로 모은 시트
- `피벗요약`: 피벗 요약표 결과
- `후파일_메모`: 전/후 파일 검증에서 변경 셀 메모가 붙은 후 파일
- `빠진행`: 전/후 파일 검증에서 후 파일에 없어진 행의 전 파일 원본 내용

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
- 다른 이름의 컬럼과 앞자리 0이 다른 키 자동 매칭
- 전/후 파일 검증
- 전/후 파일 검증의 추가 행, 빠진 행, 이름이 바뀐 컬럼 비교
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
