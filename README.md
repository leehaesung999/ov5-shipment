# OV5 지정출고 자동매칭 (Streamlit 웹 UI)

매일 OV5 위치의 Lock 재고를 거래처(농협/대리점/FS/MS/급식)에 자동 매칭하는 도구.
브라우저로 열리고 단계별로 결과가 나타납니다.

## 빠른 시작

### A. 개발 환경 (이미 Python 있음)
```bash
pip install -r requirements.txt
실행.bat   # 더블클릭
```
또는:
```bash
streamlit run app_streamlit.py
```

### B. 배포 환경 (Python 없는 PC) — portable Python 동봉
1. 받은 portable Python zip (`입고_파레트구분기.zip`)의 `python` 폴더를
   이 프로젝트 폴더에 그대로 복사
2. `pip install` 대신, 동봉 python에 의존성 설치:
   ```cmd
   python\python.exe -m pip install streamlit pandas openpyxl
   ```
3. `실행.bat` 더블클릭 → 자동으로 portable Python 사용

`실행.bat`이 동봉 python 폴더가 있으면 자동으로 사용, 없으면 시스템 Python 사용.

### C. 배포용 zip 만들기 (다른 사람에게 줄 때)

```cmd
build.bat   # 더블클릭
```

자동으로:
1. 동봉 python에 streamlit/pandas/openpyxl 설치
2. 모든 필요 파일(`core/`, `config/`, `app_streamlit.py`, `실행.bat`, `.streamlit/`, `python/`) 을
   `dist/OV5_지정출고_배포_YYYYMMDD.zip` 으로 묶음
3. 임시·캐시 파일(`__pycache__`, `tmp_uploads`, `output`, `.bak`)은 자동 제외

배포받은 사람:
- zip 풀고 → `실행.bat` 더블클릭 → 끝

## 사용 흐름

```
실행.bat 더블클릭
   ↓
브라우저 자동 오픈 (http://localhost:8501)
   ↓
① 재고 파일 업로드 (로케이션별 재고조회)
   → 즉시 OV Lock 재고 현황판 표시 (잔존율·판정 미리보기)
   → 각 행의 [지정 카테고리] 셀에서 농협/대리점/FS/MS/급식 직접 선택 가능
   ↓
② 주문 파일 업로드 (출고진행현황)
   → 즉시 매칭 자동 실행
   ↓
③ 매칭 결과 확인 (OK/NG/매칭 색상 구분)
   → [매칭된 항목만 보기/저장] 토글
   ↓
④ Excel 다운로드 (브라우저로 다운로드 + 자동 로트 누적)
```

## 매칭 우선순위

```
재고 1건당
   ↓
화면 편집한 카테고리 또는 등록된 lot이 있는가?
   ├─ YES → 그 카테고리 주문에만 매칭 (나머지 차단)
   └─ NO  → ① 농협 → ② FS/MS (등록 재고만) → ③ GT 대리점 (OK + FS/MS 미등록)
```

같은 카테고리 안에서 거래처명에 "급식" 포함된 곳은 **항상 마지막** (잔여 처리용).

## 사이드바 — 기준정보 관리

| 항목 | 설명 |
|---|---|
| 잔존율 기준 | 기본 0.70. 대리점 매칭 OK/NG 판정 기준 |
| OV 로케이션 | OV1/4/5/6 중 다중 선택 |
| 기준일자 | 잔존율 계산 기준일 (기본 오늘) |
| 기준정보 xlsx | 품목별 유통기한(월) 캐시. `Item code` + `유통기한(월)` 컬럼 |
| FS/MS 품목 xlsx | `Item code` + `구분(FS/MS)` + `유통기한`(옵션) |
| 로트 지정 xlsx | `Item code` + `유통기한`(옵션) + `카테고리(농협/대리점/FS/MS/급식)` |

## 자동 로트 누적

- 다운로드 클릭 시 매칭 결과를 `config/lot_assignments_auto.xlsx`에 자동 누적
- 다음 날 매칭 실행 시 **오늘 재고에 같은 (품목+유통기한)이 없으면 자동 제거**
- 사용자 수동 등록한 `lot_assignments.xlsx`는 영구 유지

## 폴더 구조

```
designated-shipment/
├── 실행.bat                      ← 더블클릭 실행
├── app_streamlit.py              ← Streamlit 앱 (메인)
├── app.py                        ← (구) Tkinter 앱 — 보존
├── requirements.txt
├── README.md
├── .streamlit/config.toml
├── core/                         ← 핵심 매칭 엔진
│   ├── stock_loader.py
│   ├── order_loader.py
│   ├── master_loader.py
│   ├── shelf_life.py
│   ├── classifier.py
│   ├── matcher.py
│   └── writer.py
├── config/
│   ├── settings.json
│   ├── master_info.xlsx          ← 기준정보 캐시
│   ├── fs_ms_items.xlsx          ← 업로드 시 생성
│   ├── lot_assignments.xlsx      ← 수동 등록 (영구)
│   └── lot_assignments_auto.xlsx ← 자동 누적
├── tmp_uploads/                  ← 업로드 임시 파일
└── output/                       ← 결과 xlsx 저장
```

## 종료
- 브라우저는 닫아도 OK
- `실행.bat` 콘솔 창은 **닫지 마세요** — 서버가 멈춤
- 종료할 때 콘솔 창 닫기 또는 Ctrl+C
