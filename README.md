# 업무 데이터 편집기 (work-data-editor)

엑셀(.xlsx) 시트와 Access(.accdb) 테이블을 열어 **찾고 · 고치고 · 저장하는** 파이썬 프로그램입니다.
내 PC 안에서만 동작하며(인터넷 · 서버 불필요), 수업에서 "작은 자동화 프로그램이 커지면 이런 모습이 된다"를 보여 주는 예제입니다.

![엑셀에서 X99 13건을 한 번에 바꾸고 행을 추가한 화면](screenshots/1_excel_replace.png)

이 문서는 처음 하는 분도 따라 할 수 있도록 **VS Code 에서 폴더 열기 → 가상 환경 만들기 → 라이브러리 설치 → 실행 → 노트북** 순서로 적었습니다.

## 목차

1. [준비물](#1-준비물)
2. [내려받기](#2-내려받기)
3. [VS Code 에서 폴더 열기](#3-vs-code-에서-폴더-열기)
4. [가상 환경 만들기](#4-가상-환경-만들기)
5. [라이브러리 설치](#5-라이브러리-설치)
6. [pip freeze 와 requirements.txt](#6-pip-freeze-와-requirementstxt)
7. [프로그램 실행](#7-프로그램-실행)
8. [노트북으로 실행하기](#8-노트북으로-실행하기)
9. [프로그램 사용법](#9-프로그램-사용법)
10. [막혔을 때](#10-막혔을-때)
11. [파일 구성 · 확인한 환경](#11-파일-구성--확인한-환경)

---

## 1. 준비물

| 준비물 | 설명 |
| --- | --- |
| Python 3.10 이상 | [python.org](https://www.python.org/downloads/) 에서 설치. 설치 첫 화면에서 **Add python.exe to PATH** 를 체크 |
| VS Code | [code.visualstudio.com](https://code.visualstudio.com/) |
| VS Code 확장 2개 | 왼쪽 확장(네모 4개) 아이콘에서 **Python**, **Jupyter** 를 검색해 설치 (둘 다 Microsoft 제작) |
| (선택) Access 드라이버 | Access 파일을 열 때만 필요. Access 가 설치된 PC 라면 보통 이미 있습니다 |

설치 확인: VS Code 메뉴 **터미널 → 새 터미널** 을 열고 아래를 입력합니다. 버전 숫자가 나오면 됩니다.

```
python --version
```

## 2. 내려받기

**방법 A — ZIP (쉬움)**: 이 페이지 위쪽 초록색 **Code** 버튼 → **Download ZIP** → 압축을 풀어 원하는 곳에 둡니다.

**방법 B — git**

```
git clone https://github.com/lsy3709/work-data-editor.git
```

## 3. VS Code 에서 폴더 열기

1. VS Code 메뉴 **파일 → 폴더 열기** 에서 `work-data-editor` 폴더를 고릅니다. (ZIP 으로 받았다면 `data_editor.py` 가 바로 보이는 폴더)
2. "이 폴더의 작성자를 신뢰합니까?" 가 나오면 **예** 를 누릅니다.
3. 메뉴 **터미널 → 새 터미널** 을 엽니다. 터미널의 경로 끝이 `work-data-editor` 인지 확인합니다.

파일 하나만 여는 것이 아니라 **폴더를 여는 것**이 중요합니다. 그래야 터미널 위치와 가상 환경이 이 폴더 기준으로 잡힙니다.

## 4. 가상 환경 만들기

가상 환경은 **이 폴더 전용 파이썬 상자**입니다. 여기에 설치한 라이브러리는 다른 작업에 영향을 주지 않습니다.

### 4-1. 만들기 (한 번만)

```
python -m venv .venv
```

폴더 안에 `.venv` 폴더가 생깁니다. 지워도 다시 만들면 됩니다.

### 4-2. 켜기 (터미널을 새로 열 때마다)

```
.venv\Scripts\Activate.ps1
```

줄 맨 앞에 `(.venv)` 가 붙으면 켜진 것입니다. 끌 때는 `deactivate` 를 입력합니다.

"이 시스템에서 스크립트를 실행할 수 없으므로…" 라는 빨간 글씨가 나오면, 터미널 오른쪽 위 `+` 옆 화살표에서 **Command Prompt** 를 골라 새 터미널을 열고 아래처럼 켭니다.

```
.venv\Scripts\activate.bat
```

### 4-3. VS Code 에 알려 주기

`Ctrl + Shift + P` → **Python: Select Interpreter** 입력 → 목록에서 **`.venv`** 가 들어간 항목을 고릅니다.
이렇게 해 두면 새 터미널을 열 때 가상 환경이 자동으로 켜지고, 오른쪽 위 ▶ 실행도 이 환경으로 돌아갑니다.

> VS Code 에 맡기는 방법도 있습니다: `Ctrl + Shift + P` → **Python: Create Environment** → **Venv** → 파이썬 버전 선택 → `requirements.txt` 체크. 4 · 5번을 한 번에 해 줍니다.

## 5. 라이브러리 설치

가상 환경이 켜진 터미널(`(.venv)` 표시)에서 설치합니다.

### 방법 A — 하나씩

```
pip install openpyxl
pip install pyodbc
pip install pandas
pip install ipykernel
```

| 라이브러리 | 쓰는 곳 |
| --- | --- |
| openpyxl | 엑셀 읽기 · 쓰기 (필수) |
| pyodbc | Access 연결 (Access 를 안 쓰면 없어도 됨) |
| pandas | 노트북에서 표 보기 |
| ipykernel | VS Code 에서 노트북 실행 |

### 방법 B — 한꺼번에 (권장)

필요한 라이브러리 이름을 적어 둔 `requirements.txt` 를 pip 에 넘깁니다.

```
pip install -r requirements.txt
```

설치된 것 확인:

```
pip list
```

## 6. pip freeze 와 requirements.txt

### requirements.txt 란

"이 프로젝트에 필요한 라이브러리 목록"을 적은 글자 파일입니다. 한 줄에 하나씩 씁니다.

```
openpyxl
pyodbc
pandas
ipykernel
```

이 파일이 있으면 다른 PC · 다른 사람도 `pip install -r requirements.txt` 한 줄로 같은 환경을 만듭니다.

### pip freeze — 지금 설치된 것을 버전까지 적어 내기

```
pip freeze
```

`openpyxl==3.1.5` 처럼 **이름==버전** 이 줄줄이 나옵니다. 내가 직접 설치한 4개 말고도, 그 라이브러리들이 필요로 해서 같이 깔린 것까지 모두 나옵니다.

파일로 저장하려면:

```
pip freeze > requirements-lock.txt
```

### 두 파일의 차이

| 파일 | 내용 | 언제 쓰나 |
| --- | --- | --- |
| `requirements.txt` | 직접 필요한 것만, 버전 없음 | 평소. 설치할 때의 최신 버전이 깔림 |
| `requirements-lock.txt` | `pip freeze` 결과, 버전 고정 | "내 PC 에서 되던 그 버전 그대로" 다시 만들 때 |

버전까지 똑같이 설치:

```
pip install -r requirements-lock.txt
```

이 저장소의 `requirements-lock.txt` 는 **Python 3.13** 에서 만든 것입니다. 다른 파이썬 버전에서는 일부 버전이 설치되지 않을 수 있으니 `requirements.txt` 를 쓰세요.

### 알아 둘 것

- **가상 환경을 켠 상태에서** freeze 하세요. 끈 상태에서 하면 PC 전체에 깔린 것이 모두 나옵니다.
- 라이브러리를 새로 설치했으면 `requirements.txt` 에 이름을 한 줄 추가합니다.
- `requirements.txt` 에는 **한글을 쓰지 마세요**(주석 포함). 한국어 Windows 의 pip 이 파일을 읽지 못해 `UnicodeDecodeError` 가 납니다. 이 프로젝트를 만들 때 실제로 겪은 문제입니다.
- `.venv` 폴더는 다른 사람에게 주지 않습니다. `requirements.txt` 만 주면 각자 만듭니다. (그래서 `.gitignore` 에 `.venv/` 가 들어 있습니다)

## 7. 프로그램 실행

### 7-1. 연습용 데이터 준비 (처음 한 번)

```
python reset_samples.py
```

`samples_original` 의 원본을 `samples` 폴더로 복사합니다. 연습하다가 처음 상태로 돌리고 싶을 때도 이 명령을 씁니다.

### 7-2. 실행

```
python data_editor.py
```

창이 뜨면 **[파일 열기]** 로 `samples` 폴더의 파일을 엽니다. 파일을 바로 지정할 수도 있습니다.

```
python data_editor.py samples\보금자리론_신청_2026_10.xlsx
python data_editor.py samples\업무DB.accdb
```

VS Code 에서 `data_editor.py` 를 열고 오른쪽 위 ▶ 를 눌러도 됩니다.

### 7-3. 제대로 설치됐는지 확인 (선택)

```
python test_data_editor.py
python gui_check.py
```

- 첫 줄: 화면 없이 핵심 기능 11가지를 검사합니다. 마지막에 `OK` 가 나오면 정상입니다. (pyodbc 가 없으면 `OK (skipped=3)`)
- 둘째 줄: 창을 실제로 띄워 열기 · 찾기 · 바꾸기 · 저장을 자동으로 해 봅니다. 창이 잠깐 떴다 닫히고 `RESULT PASS` 가 나오면 정상입니다.

## 8. 노트북으로 실행하기

노트북(`.ipynb`)은 코드를 **한 칸씩 실행하면서 결과를 바로 아래에서 보는** 방식입니다. 화면 프로그램이 속에서 하는 일을 따라가 보기에 좋습니다.

### 8-1. 준비

1. VS Code 확장 **Jupyter** 가 설치되어 있어야 합니다. ([1. 준비물](#1-준비물))
2. 가상 환경에 `ipykernel` 과 `pandas` 가 있어야 합니다. 5번에서 `pip install -r requirements.txt` 를 했다면 이미 있습니다.

### 8-2. 열고 커널 고르기

1. 왼쪽 탐색기에서 `notebook_tutorial.ipynb` 를 클릭합니다.
2. 노트북 오른쪽 위 **커널 선택**(Select Kernel) → **Python 환경** → **`.venv`** 가 들어간 항목을 고릅니다.
   커널은 "이 노트북을 실행할 파이썬"입니다. 4번에서 만든 가상 환경을 고르는 것입니다.
3. 목록에 `.venv` 가 없으면 VS Code 를 닫았다 다시 열거나, `Ctrl + Shift + P` → **Developer: Reload Window** 를 실행합니다.

### 8-3. 실행

- 칸(셀)을 클릭하고 `Shift + Enter` → 그 칸을 실행하고 다음 칸으로 이동
- 위쪽 **모두 실행**(Run All) → 처음부터 끝까지 실행
- 변수가 꼬였으면 위쪽 **다시 시작**(Restart) 후 처음부터

**1번 칸**의 결과에서 `실행 위치` 가 이 폴더의 `.venv` 안이면 환경이 맞게 잡힌 것입니다. 위에서부터 순서대로 실행하세요. 2번 칸은 연습용 데이터를 처음 상태로 되돌립니다.

### 8-4. 노트북에서 배우는 것

파일 열기 → 찾기 → 일괄 바꾸기 → 칸 고치기(잘못된 값 거부) → 행 추가 · 삭제 · 되돌리기 → 변경 내역 → 저장(백업) → pandas 로 확인 → Access. 마지막에 연습 문제 3개가 있습니다.

GitHub 에서 [notebook_tutorial.ipynb](notebook_tutorial.ipynb) 를 누르면 실행 결과가 들어 있는 상태로 미리 볼 수 있습니다.

## 9. 프로그램 사용법

| 기능 | 방법 |
| --- | --- |
| 파일 열기 | [파일 열기] 또는 `Ctrl+O`. 엑셀은 시트, Access 는 테이블을 오른쪽 위에서 고름 |
| 찾기 | 열을 고르고(또는 전체 열) 글자를 넣고 Enter. [전체 보기]로 해제 |
| 정렬 | 열 이름 클릭 (다시 누르면 반대로) |
| 칸 고치기 | 칸을 더블클릭 → 입력 → Enter (Esc 취소). 숫자 · 날짜 열에 맞지 않는 값은 거부 |
| 행 추가 · 삭제 | [행 추가], 행을 고르고 [행 삭제] 또는 Delete |
| 일괄 바꾸기 | 한 열에서 ① 칸 전체가 같을 때 ② 일부 글자만 ③ 모두 같은 값으로 채우기 |
| 되돌리기 | [되돌리기] 또는 `Ctrl+Z` (저장 전까지) |
| 변경 내역 | 저장하면 반영될 수정 · 추가 · 삭제 목록 |
| 저장 | [저장] 또는 `Ctrl+S`. 원본 옆 `백업` 폴더에 복사본을 먼저 만든 뒤 저장 |
| 내보내기 | 화면에 보이는 행만 새 엑셀 파일로 |

노란 줄은 고친 행, 초록 줄은 새로 넣은 행입니다.

![Access 테이블에서 반려 건을 재심사로 바꾼 화면](screenshots/3_access_edit.png)

**저장 방식**

- **엑셀**: 고른 시트의 데이터만 다시 씁니다. 다른 시트와 머리글 서식은 그대로입니다. 엑셀에서 그 파일이 열려 있으면 저장할 수 없습니다.
- **Access**: 바뀐 곳만 반영합니다. 어느 행인지 알려 주는 **기준 열**(값이 행마다 다른 열, 예: 신청번호)이 필요하며 자동으로 골라 줍니다. 하나라도 실패하면 전부 취소합니다.

**따라 해 보기 (5분)**

1. `samples\보금자리론_신청_2026_10.xlsx` 열기 → 240건
2. 찾기: 열 `지점코드`, 글자 `X99` → 13건
3. [일괄 바꾸기]: `지점코드` 열, `X99` → `ON1` → 13줄이 노랗게
4. 금액 칸을 더블클릭해 글자를 넣어 보기 → 거부됨
5. [변경 내역] → [저장] → `samples\백업` 폴더에 복사본이 생긴 것 확인
6. `samples\업무DB.accdb` 를 열어 같은 방법으로 고치고 저장 → Access 에서 열어 확인

## 10. 막혔을 때

| 증상 | 이렇게 |
| --- | --- |
| `python` 을 찾을 수 없다고 나옴 | 파이썬 설치 때 **Add python.exe to PATH** 를 체크하지 않은 경우. 다시 설치하거나 `py` 로 바꿔 입력 (`py -m venv .venv`) |
| `Activate.ps1` … 스크립트를 실행할 수 없음 | Command Prompt 터미널에서 `.venv\Scripts\activate.bat` ([4-2](#4-2-켜기-터미널을-새로-열-때마다)) |
| `ModuleNotFoundError: No module named 'openpyxl'` | 가상 환경이 꺼져 있거나 설치를 안 한 경우. `(.venv)` 표시 확인 후 `pip install -r requirements.txt` |
| `pip install -r` 에서 `UnicodeDecodeError` | `requirements.txt` 에 한글이 들어 있음. 영문만 남기기 |
| 노트북에서 커널 목록에 `.venv` 가 없음 | `ipykernel` 설치 확인 후 **Developer: Reload Window** |
| 노트북에서 `No module named 'data_editor'` | 파일만 열지 말고 **폴더를 열어야** 합니다 ([3번](#3-vs-code-에서-폴더-열기)) |
| Access 에 연결하지 못했습니다 | pyodbc 설치 여부, Access 드라이버 설치 여부, 드라이버와 Python 의 비트 수(64/32)가 같은지 확인 |
| 엑셀 저장이 안 됨 | 엑셀에서 그 파일을 닫고 다시 저장 |
| 연습 데이터를 망쳤다 | `python reset_samples.py` |

## 11. 파일 구성 · 확인한 환경

| 파일 | 내용 |
| --- | --- |
| `data_editor.py` | 프로그램 전체 (한 파일). 위에서부터 값 다루기 → 메모리에서 고치기 → 엑셀 → Access → 화면 |
| `notebook_tutorial.ipynb` | 한 칸씩 따라 하는 노트북 |
| `requirements.txt` | 필요한 라이브러리 목록 |
| `requirements-lock.txt` | `pip freeze` 로 만든 버전 고정 목록 (Python 3.13) |
| `reset_samples.py` | 연습용 데이터를 처음 상태로 |
| `test_data_editor.py` | 화면 없이 도는 핵심 기능 테스트 11개 |
| `gui_check.py` | 창을 띄워 자동으로 확인하는 점검 24항목 |
| `samples_original/` | 연습용 가상 데이터 원본 (엑셀 4개, Access 1개) |
| `screenshots/` | 실행 화면 |

**확인한 환경 (2026-11-17, Windows 11)**

| 항목 | Python 3.13.2 | Python 3.10.11 |
| --- | --- | --- |
| 새 가상 환경 + `pip install -r requirements.txt` | 통과 | 통과 |
| `test_data_editor.py` 11개 | 통과 | 통과 |
| `gui_check.py` 24항목 | 통과 | 통과 |
| `pip freeze > requirements-lock.txt` → 다른 가상 환경에 설치 | 통과 | 통과 |
| 노트북 12칸 실행 | 통과 | 통과 |

- 터미널은 Windows PowerShell 5.1 기준으로 확인했습니다. Python 3.11 · 3.12 는 직접 돌려 보지 않았습니다.
- VS Code 화면에서 버튼을 누르는 과정(인터프리터 · 커널 선택)은 글로 적은 순서이며, 자동 점검에는 포함되지 않았습니다.

**알아 둘 것**

- 연습용 데이터는 모두 **가상 데이터**입니다. 실제 기관 · 인물 · 수치와 무관합니다.
- 엑셀은 첫 줄을 열 이름으로 봅니다. 위에 제목 줄이나 병합 칸이 있는 양식에는 맞지 않습니다.
- 차트 · 피벗이 있는 엑셀은 저장 때 일부가 빠질 수 있습니다. 저장 전에 만들어지는 백업을 확인하세요.
- 실제 업무 파일에 쓰기 전에 반드시 복사본으로 먼저 해 보세요.
