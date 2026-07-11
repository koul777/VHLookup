# Public Admin Demo Samples

공공기관 행정직 사용 흐름을 바로 시연하기 위한 로컬 샘플입니다.

CSV 파일은 Windows Excel에서 더블클릭으로 열어도 한글이 깨지지 않도록 UTF-8 BOM 형식으로 저장했습니다.

## 폴더 구조

실행 화면의 1~7번 주요 기능 순서에 맞춰 샘플 폴더를 나눴습니다.

```text
samples/public_admin
├─ 01_privacy_masking
├─ 02_split_sheets
├─ 03_merge_files
│  ├─ row_merge_school_submissions
│  ├─ row_merge_messy_headers
│  ├─ row_merge_submission_errors
│  ├─ column_merge_hr_training
│  ├─ column_merge_allowance_budget
│  └─ large_column_merge_10k
├─ 04_before_after_validation
│  └─ large_before_after_diagonal_10k
├─ 05_horizontal_table
│  ├─ monthly_budget_wide.csv
│  └─ large_monthly_budget_wide_10k.csv
└─ 06_pivot_summary
   └─ budget_execution.csv
└─ 07_organization_order
   └─ department_tasks.csv
```

## 1. 개인정보 마스킹

`01_privacy_masking/citizen_service_requests.csv`를 사용합니다.

확인할 수 있는 흐름:
- 성명 `***` 마스킹
- 주민등록번호/외국인등록번호 `******-*******` 마스킹
- 연락처 `***` 마스킹
- 이메일 `***` 마스킹
- 계좌번호 `***` 마스킹
- 주소 `***` 마스킹
- 성별, 나이, 생년월일 `***` 마스킹
- `생년월일 1900/01/01`, `93년8월11일생`, `생년 930811` 같은 문맥형 생년월일 마스킹

## 2. 분류별 시트 나누기

`02_split_sheets/budget_execution.csv`를 사용합니다.

추천 선택:
- 분류 기준열: `부서`

확인할 수 있는 흐름:
- 부서별 시트 자동 생성
- 전체 원본 시트 보존

## 3. 엑셀/CSV 파일 여러 개 합치기

행 합치기 샘플:
- `03_merge_files/row_merge_school_submissions`
- `03_merge_files/row_merge_messy_headers`
- `03_merge_files/row_merge_submission_errors`

열 합치기 샘플:
- `03_merge_files/column_merge_hr_training`
- `03_merge_files/column_merge_allowance_budget`
- `03_merge_files/large_column_merge_10k`

확인할 수 있는 흐름:
- 제목/안내문이 있는 파일의 헤더 자동 탐지
- 합치기 방향 자동 선택
- `사번`과 `직원번호`처럼 다른 열 이름 자동 매칭
- `00125`와 `125`처럼 앞자리 0이 다른 값 비교
- 한쪽 파일에만 있는 행이나 열은 빈 셀에 노란색과 메모 표시
- 결과 첫 시트에는 실제 업무 컬럼만 표시
- 파일 합치기 결과에서는 개인정보 의심 컬럼을 따로 색칠하지 않음
- 자동 매칭 근거와 확인할 점은 `확인사항` 시트에 표시
- 복합 키 `사번 + 지급월` 기준 열 합치기
- 대용량 실행 중 진행률 창에 파일 읽기, 키 찾기, 결과 저장 단계 표시

대용량 성능 확인 샘플:
- `03_merge_files/large_column_merge_10k/large_employee_master_10k.xlsx`
- `03_merge_files/large_column_merge_10k/large_training_results_10k.xlsx`

두 파일은 각각 10,000행 x 71열이며, 열 합치기 결과는 10,000행 x 141열입니다. 실제 개인정보가 아닌 합성 식별자와 테스트 값만 들어 있습니다.

## 4. 전/후 파일 검증

`04_before_after_validation/payment_before.csv`와 `04_before_after_validation/payment_after.csv`를 사용합니다.

확인할 수 있는 흐름:
- 같은 사번의 금액 변경 감지
- 전 파일에 있던 행이 후 파일에서 빠진 경우 표시
- 후 파일에 새로 생긴 행 표시
- 결과 엑셀의 `후파일_메모` 시트에서 변경 셀 메모 확인
- 새로 생긴 행은 파란색 전체 행, 사라진 행은 빨간색 전체 행으로 표시
- 행 추가/누락과 컬럼 변경은 `확인사항` 시트에서 확인

대용량 대각선 변경 샘플:
- `04_before_after_validation/large_before_after_diagonal_10k/before_10k_50cols.csv`
- `04_before_after_validation/large_before_after_diagonal_10k/after_diagonal_changes_10k_50cols.csv`

두 파일은 각각 10,000행 x 50열입니다. `검증ID` 키는 양쪽 모두 같고, 후 파일의 `항목01`~`항목49` 대각선 49개 셀만 다른 글자로 바뀌어 있습니다.

## 5. 월별표 목록형 변환

`05_horizontal_table/monthly_budget_wide.csv`를 사용합니다.

확인할 수 있는 흐름:
- `1월`, `2월`, `3월`, `4월` 같은 값 열 자동 감지
- 월별 값 열을 `열 기준`, `값` 컬럼으로 풀기
- 기관명, 사업구분, 담당부서, 비고는 행 기준 컬럼으로 유지
- 파일 탑재 미리보기와 실행 중 진행률 창 표시

대용량 성능 확인 샘플:
- `05_horizontal_table/large_monthly_budget_wide_10k.csv`

대용량 샘플은 10,000행 x 17열이며, 목록형 변환 결과는 120,000행입니다. 실제 개인정보가 아닌 합성 기관명과 테스트 금액만 들어 있습니다.

## 6. 피벗 요약표 만들기

`06_pivot_summary/budget_execution.csv`를 사용합니다.

추천 선택:
- 행 기준: `부서`
- 열 기준: `월`
- 값 열: `금액`
- 집계 방식: `합계`

확인할 수 있는 흐름:
- 부서별/월별 예산 집행 합계
- 상태별 건수 요약
- `피벗요약` 시트에 요약표 생성
- 기준 설명과 확인할 점은 `확인사항` 시트에 표시

## 7. 사용자 지정 순서 정렬
`07_organization_order/department_tasks.csv`를 사용합니다.

추천 선택:
- 컬럼: `부서`
- 값 순서 예시: `기획조정실`, `총무과`, `인사과`, `예산과`, `복지정책과`, `민원봉사과`, `홍보담당관`, `감사담당관`

확인 포인트:
- 파일을 올리면 컬럼 목록이 표시됩니다.
- `부서`를 고르면 CSV 안의 부서명이 값 순서 목록에 자동으로 표시됩니다.
- 값 목록을 위/아래로 움직인 순서대로 전체 행이 정렬됩니다.
