# Public Admin Demo Samples

공공기관 행정직 사용 흐름을 바로 시연하기 위한 로컬 샘플입니다.

CSV 파일은 Windows Excel에서 더블클릭으로 열어도 한글이 깨지지 않도록 UTF-8 BOM 형식으로 저장했습니다.

## 폴더 구조

실행 화면의 1~5번 기능 순서에 맞춰 샘플 폴더를 나눴습니다.

```text
samples/public_admin
├─ 01_privacy_masking
├─ 02_split_sheets
├─ 03_merge_files
│  ├─ row_merge_school_submissions
│  ├─ row_merge_messy_headers
│  ├─ row_merge_submission_errors
│  ├─ column_merge_hr_training
│  └─ column_merge_allowance_budget
├─ 04_before_after_validation
├─ 05_pivot_summary
└─ 90_extra_cli_samples
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

## 4. 전/후 파일 검증

`04_before_after_validation/payment_before.csv`와 `04_before_after_validation/payment_after.csv`를 사용합니다.

확인할 수 있는 흐름:
- 같은 사번의 금액 변경 감지
- 전 파일에 있던 행이 후 파일에서 빠진 경우 표시
- 후 파일에 새로 생긴 행 표시
- 결과 엑셀의 `후파일_메모` 시트에서 변경 셀 메모 확인
- 새로 생긴 행은 파란색 전체 행, 사라진 행은 빨간색 전체 행으로 표시
- 행 추가/누락과 컬럼 변경은 `확인사항` 시트에서 확인

## 5. 피벗 요약표 만들기

`05_pivot_summary/budget_execution.csv`를 사용합니다.

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

## 추가 CLI 샘플

첫 화면 1~5번 메뉴 밖의 추가 샘플은 `90_extra_cli_samples`에 따로 둡니다.

- `90_extra_cli_samples/submission_reconciliation`: 제출 대상자 누락 확인
- `90_extra_cli_samples/horizontal_table`: 월별 가로표 세로 변환
