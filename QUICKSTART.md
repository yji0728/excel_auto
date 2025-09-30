# Excel Auto - 빠른 시작 가이드

## 설치

### 1. 저장소 클론
```bash
git clone https://github.com/yji0728/excel_auto.git
cd excel_auto
```

### 2. 의존성 설치
```bash
pip install -r requirements.txt
```

### 3. (선택사항) tkinter 설치
GUI를 사용하려면 tkinter가 필요합니다:

**Ubuntu/Debian:**
```bash
sudo apt-get install python3-tk
```

**macOS:**
```bash
brew install python-tk
```

**Windows:**
Python 설치 시 tkinter가 자동으로 포함됩니다.

## 실행 방법

### GUI 실행
```bash
python main.py
```

또는 런처 사용:
```bash
python run.py
```

### 테스트 실행
```bash
python run.py --test
```

### 사용 예제 실행
```bash
python run.py --example
```

## GUI 사용법

### 1. 파일 열기
1. 메뉴: `파일` > `엑셀 열기`
2. 파일 선택
3. 파일이 열리면 상태 바에 표시됨

### 2. 워크플로우 생성
1. 메뉴: `워크플로우` > `새 워크플로우`
2. 이름과 설명 입력
3. `워크플로우 편집기` 탭에서 단계 추가
4. 각 단계의 작업과 매개변수 설정

### 3. 워크플로우 실행
1. 왼쪽 목록에서 워크플로우 선택
2. `실행` 버튼 클릭
3. `출력/로그` 탭에서 결과 확인

### 4. 시나리오 사용
1. `시나리오` 탭 선택
2. 원하는 시나리오 버튼 클릭
3. 필요한 정보 입력 (파일, 열 이름 등)
4. 자동으로 워크플로우가 생성되고 실행됨

## 프로그래밍 방식 사용

### 기본 예제
```python
from excel_operations import ExcelOperations

# 새 파일 생성
excel_ops = ExcelOperations()
excel_ops.create_new_file()

# 데이터 쓰기
excel_ops.write_cell(1, 1, "이름")
excel_ops.write_cell(1, 2, "나이")
excel_ops.write_range(2, 1, [
    ["홍길동", 30],
    ["김철수", 25],
])

# 파일 저장
excel_ops.save_file("output.xlsx")
excel_ops.close()
```

### 워크플로우 예제
```python
from workflow_engine import WorkflowEngine

# 워크플로우 생성
engine = WorkflowEngine()
engine.create_workflow("내_작업", "데이터 처리")

# 단계 추가
engine.add_step("내_작업", "open_file", {"file_path": "input.xlsx"})
engine.add_step("내_작업", "sum_column", {"col": 2, "start_row": 2, "end_row": 10})
engine.add_step("내_작업", "save_file", {"file_path": "output.xlsx"})

# 실행
results = engine.execute_workflow("내_작업")
```

## 자주 사용하는 작업

### 데이터 병합
```python
engine.create_workflow("merge", "파일 병합")
engine.add_step("merge", "merge_files", {
    "file_paths": ["file1.xlsx", "file2.xlsx", "file3.xlsx"]
})
engine.add_step("merge", "save_file", {"file_path": "merged.xlsx"})
engine.execute_workflow("merge")
```

### 데이터 정렬
```python
engine.create_workflow("sort", "데이터 정렬")
engine.add_step("sort", "open_file", {"file_path": "data.xlsx"})
engine.add_step("sort", "sort_data", {"column": "날짜", "ascending": False})
engine.add_step("sort", "save_file", {"file_path": "sorted.xlsx"})
engine.execute_workflow("sort")
```

### 서식 적용
```python
engine.create_workflow("format", "헤더 서식")
engine.add_step("format", "open_file", {"file_path": "data.xlsx"})

for col in range(1, 6):
    engine.add_step("format", "format_cell", {
        "row": 1, "col": col,
        "bold": True, "bg_color": "4472C4", "align": "center"
    })

engine.add_step("format", "save_file", {"file_path": "formatted.xlsx"})
engine.execute_workflow("format")
```

## 문제 해결

### tkinter를 찾을 수 없음
```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# macOS
brew install python-tk
```

### openpyxl을 찾을 수 없음
```bash
pip install openpyxl pandas
```

### 파일 열기 오류
- 파일 경로가 올바른지 확인
- 파일이 다른 프로그램에서 열려있지 않은지 확인
- 파일 권한 확인

### GUI가 실행되지 않음
- 테스트 모드로 실행: `python run.py --test`
- 예제 실행: `python run.py --example`
- 모듈만 사용하고 GUI 없이 프로그래밍 방식으로 사용

## 추가 리소스

- 전체 문서: README.md
- 테스트 코드: test_excel_auto.py
- 사용 예제: usage_example.py
- 워크플로우 예제: workflows/ 디렉토리

## 지원

문제가 발생하면 GitHub Issues에 보고해 주세요:
https://github.com/yji0728/excel_auto/issues
