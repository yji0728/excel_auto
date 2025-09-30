# Excel Auto - 엑셀 자동화 워크플로우 GUI 애플리케이션

Python과 tkinter를 사용하여 제작된 엑셀 자동화 워크플로우 GUI 애플리케이션입니다. 다양한 시나리오와 사용자 정의 워크플로우를 통해 엑셀 파일을 효율적으로 관리하고 처리할 수 있습니다.

## 주요 기능

### 📊 엑셀 파일 조작
- 엑셀 파일 열기, 생성, 저장
- 셀 읽기/쓰기
- 범위 데이터 읽기/쓰기
- 시트 관리 (추가, 삭제, 선택)
- 서식 적용 (폰트, 색상, 정렬)

### 🔄 워크플로우 엔진
- **사용자 정의 워크플로우**: 여러 작업을 순차적으로 조합하여 자동화
- **워크플로우 저장/불러오기**: JSON 형식으로 워크플로우 저장 및 재사용
- **단계별 실행**: 각 단계의 실행 결과를 실시간으로 확인
- **진행 상황 추적**: 워크플로우 실행 중 진행률과 로그 제공

### 🎯 사전 정의된 시나리오
- **데이터 병합**: 여러 엑셀 파일을 하나로 통합
- **데이터 정렬**: 특정 열을 기준으로 데이터 정렬
- **데이터 필터링**: 조건에 맞는 데이터만 추출
- **합계 계산**: 열의 합계 및 평균 계산
- **서식 적용**: 헤더 및 데이터에 자동 서식 적용
- **시트 복사**: 시트를 다른 파일로 복사

### 💻 GUI 인터페이스
- 직관적인 사용자 인터페이스
- 워크플로우 목록 관리
- 워크플로우 편집기
- 실행 로그 및 결과 표시
- 최근 파일 목록

## 설치 방법

### 1. 저장소 클론
```bash
git clone https://github.com/yji0728/excel_auto.git
cd excel_auto
```

### 2. 필수 패키지 설치
```bash
pip install -r requirements.txt
```

필요한 패키지:
- `openpyxl`: 엑셀 파일 읽기/쓰기
- `pandas`: 데이터 처리 및 분석
- `xlrd`: 구형 엑셀 파일 지원
- `pillow`: 이미지 처리 (선택사항)

### 3. 애플리케이션 실행
```bash
python main.py
```

## 사용 방법

### 기본 사용법

#### 1. 엑셀 파일 열기
- `파일` > `엑셀 열기` 메뉴를 선택하거나
- 최근 파일 목록에서 선택

#### 2. 워크플로우 생성
```
워크플로우 > 새 워크플로우
```
- 워크플로우 이름과 설명 입력
- 작업 단계 추가
  - 작업 선택 (open_file, write_cell, save_file 등)
  - JSON 형식으로 매개변수 입력

#### 3. 워크플로우 실행
- 워크플로우 목록에서 선택
- `실행` 버튼 클릭
- 실행 결과는 `출력/로그` 탭에서 확인

### 사전 정의된 시나리오 사용

`시나리오` 탭에서 자주 사용하는 작업을 빠르게 실행:

1. **데이터 병합**: 여러 파일 선택 → 자동 병합 → 결과 저장
2. **데이터 정렬**: 파일 열기 → 정렬할 열 입력 → 자동 정렬
3. **서식 적용**: 파일 열기 → 헤더에 자동으로 서식 적용

## 워크플로우 작업 목록

### 파일 작업
- `open_file`: 엑셀 파일 열기
  ```json
  {"file_path": "data.xlsx", "sheet_name": "Sheet1"}
  ```
- `save_file`: 파일 저장
  ```json
  {"file_path": "output.xlsx"}
  ```

### 셀 작업
- `read_cell`: 셀 값 읽기
  ```json
  {"row": 1, "col": 1}
  ```
- `write_cell`: 셀에 값 쓰기
  ```json
  {"row": 1, "col": 1, "value": "제목"}
  ```

### 범위 작업
- `read_range`: 범위 데이터 읽기
  ```json
  {"start_row": 1, "start_col": 1, "end_row": 10, "end_col": 5}
  ```
- `write_range`: 범위에 데이터 쓰기
  ```json
  {"start_row": 1, "start_col": 1, "data": [[1, 2], [3, 4]]}
  ```

### 서식 작업
- `format_cell`: 셀 서식 적용
  ```json
  {"row": 1, "col": 1, "bold": true, "font_size": 12, "bg_color": "FFFF00", "align": "center"}
  ```

### 데이터 처리
- `sum_column`: 열 합계 계산
  ```json
  {"col": 2, "start_row": 2, "end_row": 10}
  ```
- `average_column`: 열 평균 계산
  ```json
  {"col": 2, "start_row": 2, "end_row": 10}
  ```
- `sort_data`: 데이터 정렬
  ```json
  {"column": "날짜", "ascending": true}
  ```
- `filter_data`: 데이터 필터링
  ```json
  {"column": "금액", "operator": ">", "value": 1000}
  ```

### 고급 작업
- `merge_files`: 여러 파일 병합
  ```json
  {"file_paths": ["file1.xlsx", "file2.xlsx", "file3.xlsx"]}
  ```
- `copy_data`: 데이터 복사
  ```json
  {"src_start_row": 1, "src_start_col": 1, "src_end_row": 5, "src_end_col": 3, "dest_row": 10, "dest_col": 1}
  ```
- `add_sheet`: 시트 추가
  ```json
  {"sheet_name": "새시트"}
  ```
- `delete_sheet`: 시트 삭제
  ```json
  {"sheet_name": "Sheet1"}
  ```

## 워크플로우 예제

### 예제 1: 데이터 병합 및 정렬
```json
{
  "name": "데이터_병합_정렬",
  "description": "여러 파일을 병합하고 정렬",
  "steps": [
    {
      "operation": "merge_files",
      "params": {"file_paths": ["data1.xlsx", "data2.xlsx"]}
    },
    {
      "operation": "sort_data",
      "params": {"column": "날짜", "ascending": false}
    },
    {
      "operation": "save_file",
      "params": {"file_path": "merged_sorted.xlsx"}
    }
  ]
}
```

### 예제 2: 보고서 생성
```json
{
  "name": "월간_보고서",
  "description": "월간 보고서 자동 생성",
  "steps": [
    {
      "operation": "open_file",
      "params": {"file_path": "monthly_data.xlsx"}
    },
    {
      "operation": "sum_column",
      "params": {"col": 3, "start_row": 2, "end_row": 32}
    },
    {
      "operation": "write_cell",
      "params": {"row": 33, "col": 3, "value": "합계:"}
    },
    {
      "operation": "format_cell",
      "params": {"row": 1, "col": 1, "bold": true, "bg_color": "4472C4"}
    },
    {
      "operation": "save_file",
      "params": {"file_path": "monthly_report.xlsx"}
    }
  ]
}
```

## 프로젝트 구조

```
excel_auto/
├── main.py                 # GUI 애플리케이션 메인 파일
├── excel_operations.py     # 엑셀 파일 조작 모듈
├── workflow_engine.py      # 워크플로우 엔진
├── config.py              # 설정 관리
├── requirements.txt       # 필수 패키지 목록
├── .gitignore            # Git 무시 파일 목록
├── workflows/            # 워크플로우 저장 디렉토리
│   ├── merge_example.json
│   ├── report_example.json
│   └── sort_filter_example.json
├── examples/             # 예제 파일 디렉토리
└── README.md            # 프로젝트 문서
```

## 기술 스택

- **Python 3.8+**: 주 프로그래밍 언어
- **tkinter**: GUI 프레임워크
- **openpyxl**: 엑셀 파일 처리
- **pandas**: 데이터 분석 및 처리
- **JSON**: 워크플로우 저장 형식

## 개발 로드맵

### 완료된 기능 ✅
- [x] 기본 엑셀 파일 읽기/쓰기
- [x] 워크플로우 엔진 구현
- [x] GUI 인터페이스 개발
- [x] 사전 정의된 시나리오
- [x] 워크플로우 저장/불러오기
- [x] 실행 로그 및 진행 상황 표시

### 향후 계획 🚀
- [ ] 워크플로우 단계 재정렬 기능
- [ ] 조건부 실행 (if/else)
- [ ] 반복 실행 (loop)
- [ ] 변수 및 표현식 지원
- [ ] 차트 및 그래프 생성
- [ ] 데이터 검증 및 오류 처리
- [ ] 템플릿 라이브러리
- [ ] 다국어 지원 확장
- [ ] 클라우드 저장소 연동
- [ ] 웹 버전 개발

## GitHub에서 검색한 유용한 기능

이 프로젝트는 다음 GitHub 프로젝트들에서 영감을 받았습니다:

1. **openpyxl/openpyxl**: 엑셀 파일 조작
   - 셀 서식, 스타일, 차트 기능

2. **python-excel**: 다양한 엑셀 처리 라이브러리
   - pandas, xlrd, xlwt 등의 통합

3. **workflow-engines**: 워크플로우 패턴
   - 단계별 실행, 조건 분기, 반복 처리

4. **tkinter-examples**: GUI 디자인 패턴
   - 탭, 트리뷰, 다이얼로그 등

## 기여 방법

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.

## 문의 및 지원

문제가 발생하거나 제안사항이 있으시면 GitHub Issues를 통해 문의해 주세요.

## 스크린샷

(애플리케이션 실행 후 스크린샷 추가 예정)

---

**Excel Auto** - 엑셀 작업을 더 쉽고 빠르게! 🚀