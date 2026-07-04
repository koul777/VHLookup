# 샘플 데이터 카탈로그

`samples/public_admin` 아래 샘플은 공공기관 행정직이 바로 눌러볼 수 있도록 실행 화면의 1~5번 기능 순서로 나뉘어 있습니다.

기본 폴더:

- `01_privacy_masking`: 1번 개인정보 마스킹
- `02_split_sheets`: 2번 분류별 시트 나누기
- `03_merge_files`: 3번 행/열 합치기
- `04_before_after_validation`: 4번 전/후 파일 검증
- `05_pivot_summary`: 5번 피벗 요약표 만들기
- `90_extra_cli_samples`: 첫 화면 밖의 추가 CLI 샘플

전체 데모 실행:

```text
run_demo.bat 더블클릭
```

또는 명령어로 실행:

```powershell
python scripts\run_public_admin_demo.py
```

결과는 `demo_output` 폴더에 생성됩니다. `run_demo.bat`으로 실행하면 생성 후 폴더가 자동으로 열립니다.

먼저 열어볼 파일:

- `demo_output/00_샘플_둘러보기.xlsx`

이 파일의 `업무템플릿` 시트에서 CLI에 넣을 템플릿 ID와 표준 컬럼을 확인할 수 있습니다.
각 결과 파일의 `확인사항` 시트에는 자동 매칭 근거, 한쪽 파일에만 있는 행/열, 개인정보 의심 요약이 함께 표시됩니다.

CLI 방식으로 샘플을 실행해 보려면 `run_cli_examples.bat`을 더블클릭합니다. 결과는 `cli_output` 폴더에 생성됩니다.
`cli_output/00_inspect_cli.xlsx`는 실제 수합 전 헤더 행과 컬럼 매칭을 미리 확인하는 사전점검 예시입니다.

## 01. 제출자료 수합

입력:

- `samples/public_admin/03_merge_files/row_merge_school_submissions/gangbuk_school.csv`
- `samples/public_admin/03_merge_files/row_merge_school_submissions/gangnam_school.csv`

결과:

- `demo_output/01_제출자료_수합결과.xlsx`

확인 포인트:

- 제목 행과 안내문을 건너뛰고 실제 헤더를 찾습니다.
- `제출기관`, `담당자명`, `작성일`, `신청액` 같은 다른 표현을 표준 컬럼에 맞춥니다.
- 결과 첫 시트에는 실제 업무 컬럼만 남깁니다.
- 자동 매칭 근거와 확인할 점은 `확인사항` 시트에 남깁니다.

## 02. 오류 검증 제출자료

입력:

- `samples/public_admin/03_merge_files/row_merge_submission_errors/bad_school_a.csv`
- `samples/public_admin/03_merge_files/row_merge_submission_errors/bad_school_b.csv`

결과:

- `demo_output/02_오류검증_제출자료.xlsx`

확인 포인트:

- 필수값 누락
- 숫자 오류
- 날짜 오류
- 중복 제출 의심
- `확인사항` 시트의 조치 안내

## 03. 교육이수 명단 대조

입력:

- `samples/public_admin/03_merge_files/column_merge_hr_training/hr_employee_master.csv`
- `samples/public_admin/03_merge_files/column_merge_hr_training/hr_training_completion.csv`

결과:

- `demo_output/03_교육이수_명단대조.xlsx`

확인 포인트:

- `사번`과 `직원번호`를 자동으로 연결합니다.
- 기준명단의 `성명`, `부서`, `직급`, `소속`을 교육이수 명단에 붙입니다.
- `00123`과 `123`처럼 형식이 다른 키는 자동으로 몰래 붙이지 않고 `형식 불일치`로 분리합니다.
- 한쪽 파일에만 있는 `00999` 같은 기준값도 결과에 남기고 빈 칸은 노란색 메모로 표시합니다.

## 04. 교육 누락자 확인

입력:

- `samples/public_admin/03_merge_files/column_merge_hr_training/hr_employee_master.csv`
- `samples/public_admin/03_merge_files/column_merge_hr_training/hr_training_completion.csv`

결과:

- `demo_output/04_교육누락자_대조결과.xlsx`

확인 포인트:

- 기준명단에만 있는 사람
- 교육이수 명단에만 있는 사람
- 양쪽 모두 있는 사람

## 05. 수당/예산 기준표 대조

입력:

- `samples/public_admin/03_merge_files/column_merge_allowance_budget/rate_reference.csv`
- `samples/public_admin/03_merge_files/column_merge_allowance_budget/payment_requests.csv`

결과:

- `demo_output/05_수당예산_대조결과.xlsx`

확인 포인트:

- 복합 키 `사번 + 지급월`
- `직원번호`와 `사번` 자동 연결
- 단가, 지급기준, 예산과목 자동 붙이기
- 기준표에 없는 신청자 분리

## 06. 제출 대상 누락 확인

입력:

- `samples/public_admin/90_extra_cli_samples/submission_reconciliation/expected_submitters.csv`
- `samples/public_admin/90_extra_cli_samples/submission_reconciliation/received_submitters.csv`

결과:

- `demo_output/06_제출대상_누락확인.xlsx`

확인 포인트:

- 제출 대상인데 제출하지 않은 기관
- 대상 명단에 없는데 제출한 기관

## 07. 부서별 현황 수합

입력:

- `samples/public_admin/03_merge_files/row_merge_messy_headers/department_status_a.csv`
- `samples/public_admin/03_merge_files/row_merge_messy_headers/department_status_b.csv`

결과:

- `demo_output/07_부서별현황_수합결과.xlsx`

확인 포인트:

- 제목, 작성일, 안내문이 앞에 있어도 실제 헤더를 찾습니다.
- `담당 부서`와 `소속`, `현원`과 `인원수`를 같은 표준 컬럼으로 맞춥니다.

## 08. 월별 가로표 세로 변환

입력:

- `samples/public_admin/90_extra_cli_samples/horizontal_table/monthly_budget_wide.csv`

결과:

- `demo_output/08_월별가로표_세로변환.xlsx`

확인 포인트:

- `1월`, `2월`, `3월` 컬럼을 자동 감지합니다.
- 월별 가로표를 `기관명`, `항목`, `열 기준`, `값` 형태로 변환합니다.

## 09. 전/후 파일 검증

입력:

- `samples/public_admin/04_before_after_validation/payment_before.csv`
- `samples/public_admin/04_before_after_validation/payment_after.csv`

확인 포인트:

- 같은 사번의 변경된 금액을 찾습니다.
- 전 파일에만 있거나 후 파일에만 있는 행을 분리합니다.
- 결과 엑셀의 `후파일_메모` 시트에서 변경된 셀의 메모를 확인합니다.
- 행 추가/누락과 컬럼 변경은 `확인사항` 시트에서 확인합니다.

## 10. 피벗 요약표 만들기

입력:

- `samples/public_admin/05_pivot_summary/budget_execution.csv`

추천 선택:

- 행 기준: `부서`
- 열 기준: `월`
- 값 열: `금액`
- 집계 방식: `합계`

확인 포인트:

- 부서별/월별 예산 집행 합계를 `피벗요약` 시트로 만듭니다.
- 기준 설명과 상위 항목은 `확인사항` 시트에서 확인합니다.
- `건수` 집계를 선택하면 값 열 없이 행 개수를 요약할 수 있습니다.
