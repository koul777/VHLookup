# 샘플 데이터 카탈로그

`samples/public_admin` 아래 샘플은 실행 화면의 1~7번 기능 순서로 나뉘어 있습니다.

기본 폴더:

- `01_privacy_masking`: 1번 개인정보 마스킹
- `02_split_sheets`: 2번 분류별 시트 나누기
- `03_merge_files`: 3번 행/열 합치기
- `04_before_after_validation`: 4번 전/후 파일 검증
- `05_horizontal_table`: 5번 월별표 목록형 변환
- `06_pivot_summary`: 6번 피벗 요약표 만들기
- `07_organization_order`: 7번 사용자 지정 순서 정렬

전체 데모 실행:

```text
run_demo.bat 더블클릭
```

또는 명령어로 실행:

```powershell
python scripts\run_public_admin_demo.py
```

결과는 `demo_output` 폴더에 생성됩니다. 먼저 `demo_output/00_샘플_둘러보기.xlsx`를 열어 기능 순서와 확인 포인트를 봅니다.

## 01. 개인정보 마스킹

입력:

- `samples/public_admin/01_privacy_masking/citizen_service_requests.csv`

결과:

- `demo_output/01_개인정보_마스킹결과.xlsx`

확인 포인트:

- 이름, 주민등록번호/외국인등록번호, 연락처, 이메일, 계좌번호, 주소가 가려집니다.
- 성별, 나이, 생년월일과 문장 안의 생년월일 표현도 가려집니다.
- `마스킹내역` 시트에는 원본 값 없이 위치와 유형만 남습니다.

## 02. 분류별 시트 나누기

입력:

- `samples/public_admin/02_split_sheets/budget_execution.csv`

결과:

- `demo_output/02_분류별_시트나누기.xlsx`

추천 선택:

- 분류 기준열: `부서`

확인 포인트:

- 부서별 시트가 자동 생성됩니다.
- `전체` 시트에는 원본 전체 데이터가 보존됩니다.

## 03. 엑셀/CSV 파일 여러 개 합치기

입력 예시:

- `samples/public_admin/03_merge_files/row_merge_school_submissions/gangbuk_school.csv`
- `samples/public_admin/03_merge_files/row_merge_school_submissions/gangnam_school.csv`

결과:

- `demo_output/03_제출자료_수합결과.xlsx`

추가 샘플:

- `samples/public_admin/03_merge_files/row_merge_messy_headers`
- `samples/public_admin/03_merge_files/row_merge_submission_errors`
- `samples/public_admin/03_merge_files/column_merge_hr_training`
- `samples/public_admin/03_merge_files/column_merge_allowance_budget`
- `samples/public_admin/03_merge_files/large_column_merge_10k`

확인 포인트:

- 제목 행과 안내문을 건너뛰고 실제 헤더를 찾습니다.
- `제출기관`, `담당자명`, `작성일`, `신청액` 같은 다른 표현을 표준 컬럼에 맞춥니다.
- 행 합치기와 열 합치기를 자동 추천합니다.
- 대용량 열 합치기 샘플은 입력 10,000행 x 71열 파일 2개이며, 결과는 10,000행 x 141열입니다.
- 대용량 실행 중 진행률 창에서 파일 읽기, 키 찾기, 결과 저장 단계가 표시됩니다.

## 04. 전/후 파일 검증

입력:

- `samples/public_admin/04_before_after_validation/payment_before.csv`
- `samples/public_admin/04_before_after_validation/payment_after.csv`

결과:

- `demo_output/04_전후파일_검증결과.xlsx`

대용량 샘플:

- `samples/public_admin/04_before_after_validation/large_before_after_diagonal_10k/before_10k_50cols.csv`
- `samples/public_admin/04_before_after_validation/large_before_after_diagonal_10k/after_diagonal_changes_10k_50cols.csv`

확인 포인트:

- 같은 사번의 변경된 금액을 찾습니다.
- 전 파일에만 있거나 후 파일에만 있는 행을 분리합니다.
- 변경된 셀은 `후파일_메모` 시트에서 노란색과 메모로 표시됩니다.
- 대용량 샘플은 각 파일 10,000행 x 50열이며, 후 파일의 `항목01`~`항목49` 대각선 49개 셀만 바뀌어 있습니다.

## 05. 월별표 목록형 변환

입력:

- `samples/public_admin/05_horizontal_table/monthly_budget_wide.csv`

결과:

- `demo_output/05_월별표_목록형변환.xlsx`

대용량 샘플:

- `samples/public_admin/05_horizontal_table/large_monthly_budget_wide_10k.csv`

확인 포인트:

- `1월`, `2월`, `3월`, `4월` 같은 값 열을 자동 감지합니다.
- 기관명, 사업구분, 담당부서, 비고는 행 기준 컬럼으로 유지합니다.
- 월별 값 열은 `열 기준`, `값` 컬럼으로 풀립니다.
- 대용량 샘플은 입력 10,000행 x 17열이며, 목록형 변환 결과는 120,000행입니다.
- 파일 탑재 미리보기와 실행 중 진행률 창에서 파일 읽기, 값 열 감지, 목록형 변환, 저장 단계가 표시됩니다.

## 06. 피벗 요약표 만들기

입력:

- `samples/public_admin/06_pivot_summary/budget_execution.csv`

결과:

- `demo_output/06_피벗요약표_결과.xlsx`

추천 선택:

- 행 기준: `부서`
- 열 기준: `월`
- 값 열: `금액`
- 집계 방식: `합계`

확인 포인트:

- 부서별/월별 예산 집행 합계를 `피벗요약` 시트로 만듭니다.
- `건수` 집계를 선택하면 값 열 없이 행 개수를 요약할 수 있습니다.
- 기준 설명과 상위 항목은 `확인사항` 시트에서 확인합니다.

## 07. 사용자 지정 순서 정렬
입력:

- `samples/public_admin/07_organization_order/department_tasks.csv`

추천 선택:

- 컬럼: `부서`
- 값 순서 예시: `기획조정실`, `총무과`, `인사과`, `예산과`, `복지정책과`, `민원봉사과`, `홍보담당관`, `감사담당관`

확인 포인트:

- 파일에서 컬럼을 고르면 해당 컬럼의 고유값이 목록에 자동으로 표시됩니다.
- 목록의 값을 위/아래로 옮긴 순서대로 전체 행이 정렬됩니다.
- 부서뿐 아니라 상태, 구분, 등급 같은 다른 컬럼에도 같은 방식으로 사용할 수 있습니다.
