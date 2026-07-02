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

### 1. 파일 문제 찾기(선택)

여러 엑셀/CSV 파일의 구조를 먼저 점검합니다.

- 파일별 열 차이 확인
- 헤더 위치 자동 탐지
- 빈값, 숫자 오류, 날짜 오류 확인
- 개인정보 의심 컬럼 확인
- 필수 단계가 아니라 파일 구조가 의심될 때 쓰는 선택 기능

### 2. 엑셀 파일 정리하기

지저분한 엑셀 파일을 실무용 정리본으로 만듭니다.

- 빈 행/빈 열 제거
- 앞뒤 공백 정리
- 중복 행, 빈값 현황 확인
- 필터와 틀고정 적용
- 원본 파일명, 시트명, 원본 행 번호 보존

### 3. 분류별 시트 나누기

한 파일을 특정 열 기준으로 여러 시트로 나눕니다.

- 부서, 기관명, 학교명, 상태, Department 같은 열 기준
- 기준 열은 드롭박스로 선택
- 원본 전체 시트와 분류별 시트 생성

### 4. 엑셀/CSV 파일 여러 개 합치기

여러 파일을 한 결과 파일로 합칩니다.

- `행 합치기`: 같은 양식의 여러 파일을 아래로 이어 붙임
- `열 합치기`: 공통 키를 찾아 다른 파일의 열을 오른쪽에 붙임
- 컬럼명이 달라도 자동 매칭
- 자동 매칭이 틀리면 `컬럼 매칭 수정`에서 드롭박스로 직접 지정
- 미리보기에서 앞 10행 확인 후 실행

### 5. 전/후 파일 검증

수정 전 파일과 수정 후 파일을 비교합니다.

- 같은 의미의 컬럼을 자동 매칭
- 비교 기준열 자동 추천
- 자동 추천이 틀리면 `컬럼 매칭 수정`에서 비교 기준열과 전/후 컬럼 직접 지정
- 값이 바뀐 셀, 추가 행, 누락 행, 추가/삭제 컬럼 확인
- 결과 파일의 `후파일_메모` 시트에서 변경 셀 메모 확인

### 6. 피벗 요약표 만들기

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

주요 샘플:

- `school_submissions/`: 학교/부서 제출자료 수합
- `submission_errors/`: 필수값, 숫자, 날짜, 중복 오류 검증
- `hr_employee_master.csv`: 직원 기본명단
- `hr_training_completion.csv`: 교육 이수명단
- `allowance_budget/`: 수당/예산 기준표 대조
- `submission_reconciliation/`: 제출 대상 누락 확인
- `messy_headers/`: 제목/안내문이 붙은 부서별 현황 수합
- `monthly_budget_wide.csv`: 월별 가로표 샘플
- `pivot_summary/budget_execution.csv`: 피벗 요약표 샘플
- `before_after_validation/`: 전/후 파일 검증 샘플

피벗 샘플 추천 선택:

```text
파일: samples\public_admin\pivot_summary\budget_execution.csv
행 기준: 부서
열 기준: 월
값 열: 금액
집계 방식: 합계
```

샘플 종류와 확인 포인트는 [docs/sample_catalog.md](docs/sample_catalog.md)에 더 자세히 정리되어 있습니다.

## 결과 파일

결과 파일은 기본적으로 아래 폴더에 생성됩니다.

```text
dist\outputs
```

개발 환경에서 직접 실행하면 프로젝트 루트의 `outputs` 폴더에 저장될 수 있습니다.

자주 보는 시트:

- `먼저확인`: 결과를 열었을 때 먼저 봐야 할 요약
- `결과`: 수합, 대조, 정리 결과
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

- 파일 정리
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
