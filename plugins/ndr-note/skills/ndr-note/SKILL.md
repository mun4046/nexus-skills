---
name: ndr-note
description: >
  기업 탐방(NDR/IR) 노트를 Nexus Asset Management 디자인의 Word/PDF 보고서로 자동 작성하는 스킬.
  사용자가 "탐방노트", "탐방 보고서", "NDR 노트", "기업탐방", "탐방 정리" 등을 언급하거나
  특정 종목명과 함께 "탐방노트 만들어줘/작성해줘/정리해줘"라고 하면 이 스킬을 사용할 것.
  데이터는 NotebookLM의 종목 전용 노트북(특히 *ndr.docx 녹취록)에서 가져온다.
  이것은 매매콜(투자의견/목표가)이 아니라 탐방 정리 노트임 — 매매전략 콜 필드는 넣지 않는다.
---

# 기업탐방(NDR)노트 작성 스킬

## 개요
NDR/IR 탐방 내용을 **Nexus Asset Management 양식(골드/브라운, 4파트)**의 Word 보고서로 작성하고 PDF로 변환한다.
데이터 소스는 **NotebookLM 종목 노트북**(NDR 녹취록 + 공시). 콘텐츠를 `config.json`으로 구성한 뒤
`scripts/build_ndr_note.py`로 docx를 생성하고, `scripts/convert_to_pdf.py`로 PDF를 만든다.

이 스킬은 **포터블**이다 — 폴더를 다른 PC의 `~/.claude/skills/ndr-note/`에 복사하면 그대로 동작한다(설치/이식은 `README.md` 참고). 절대경로 가정 없음: 출력물은 사용자 바탕화면에 저장된다.

**산출 구성(4파트):**
1. **표지(1p)** — 헤더밴드(`Nexus 로고 | 기업탐방 Note | 날짜`) + 좌측 데이터컬럼(Stock Data·상대수익률·주가추이·Reported by) + 우측 논지블록(종목명·섹터·헤드라인·▌탐방 개요·▌핵심 요약·▌탐방 포인트 ①②③) + 하단 Forecasts & Valuations 표
2. **심층 분석(2~3p)** — 큰 중앙 제목 + ▌소제목 섹션(`**...**`로 핵심문구 볼드+밑줄) + 로드맵/제품 표
3. **Q&A** — "Q&A" 골드 탭. NDR 녹취록에서 **실제 문답을 추출해 수치 중심으로 구체적으로**(보통 12~16개)
4. **Compliance Notice** — 골드 배경 전체면, 면책 항목

디자인 근거 양식은 `reference/` 폴더에 동봉(`(템플릿)매매전략 보고서.docx` = 1페이지 디자인·색, `선익시스템 탐방 보고서.pdf` = 내용 구성). 생성 자체는 이 파일들이 없어도 동작(디자인은 제너레이터에 내장).

> ⚠️ **이건 탐방노트지 매매콜이 아니다.** 투자의견·목표주가·매수/손절밴드 같은 매매전략 콜 필드는 넣지 않는다.

---

## 워크플로우

### 0단계: 입력 확인
- 사용자가 **종목명** 제공(필요시 종목코드).

### 1단계: NotebookLM 노트북 연결
```
mcp__notebooklm__refresh_auth        # 인증 만료 시. 실패하면 터미널에서 `nlm login`
mcp__notebooklm__notebook_list       # 종목명 포함 노트북 찾기 → notebook_id
mcp__notebooklm__notebook_get        # 소스 목록 확인 (NDR docx / 공시 PDF 식별)
```
동일 종목 노트북이 여러 개면 **소스에 NDR 녹취록(`*ndr.docx`)·공시가 든 것**을 택한다.

### 2단계: 데이터 추출
- `mcp__notebooklm__source_get_content` 로 **NDR 녹취록 전문**과 "붙여넣은 텍스트" 정리본 확보 → 핵심 요약·탐방 포인트·Q&A 원천.
- `mcp__notebooklm__notebook_query` 로 공시 수치 확보(여러 해 묶어 한 번에 질의):
  - 다년(예: 2023~2025) 매출/영업이익/순이익, 자산·부채·자본총계·부채비율, EPS·BPS·ROE
  - 유상증자 규모·신주수·자금용도, 발행주식총수, 최대주주·특수관계인 지분
  - 사업부문별 매출 비중, 최근 분기 실적

### 3단계: 시세 데이터
- 현재가·시가총액: WebFetch `https://www.google.com/finance/quote/<코드>:KOSDAQ`
- 52주 최고/최저: WebSearch(없으면 `–`). PER=현재가/EPS, PBR=현재가/BPS 계산.
- (finance.naver.com·m.stock.naver.com 은 WebFetch 차단될 수 있음)

### 4단계: config.json 작성
`scripts/config_amotech.json` 을 템플릿으로 복사해 값만 교체. 필드:
- `meta`(firm/note_type/date), `reported_by`(name/role/email)
- `cover`: name, code, sector, headline, `stock_data`[[라벨,값]…], `rel_return`, `tamban_meta`(탐방형식/일시/참석자/작성), `summary`(핵심문구는 `**...**`), `points`[[제목,설명]×3]
- `forecasts`: years[], rows[[항목,…년도값]], footnote (E열은 컨센서스 없으면 자체추정+주석)
- `deepdive`: title, sections[{head, body(`**...**`)}], roadmap{title,header,widths,rows}, source
- `qa`: [[질문, 답변]…]  ← NDR 실제 문답, 수치 포함
- `compliance`: [면책…]
- `output`: **파일명만 적으면 바탕화면에 저장**(절대경로도 가능, 생략 시 종목명 기반 자동)

### 5단계: docx 생성
```
python scripts/build_ndr_note.py <config.json>
```
필요 패키지: `pip install -r scripts/requirements.txt` (python-docx)

### 6단계: PDF 변환 (OS 독립)
```
python scripts/convert_to_pdf.py "<생성된 .docx 경로>"
```
내부적으로 LibreOffice(soffice, 전 OS) → 없으면 Windows+MS Word(COM) 순으로 시도.

### 7단계: 검증 & 확인
- Read 로 PDF를 열어 1페이지 디자인·표 분할 여부 확인.
- 표가 페이지 경계서 갈라지면 generator의 `keep_table_together`가 자동 처리. 커버가 꽉 차면 Forecasts 표가 통째로 2페이지로 이동할 수 있음(정상). 1페이지에 넣고 싶으면 핵심요약/포인트를 줄인다.
- 사용자에게 결과를 보고하고 수정 받기.

---

## 작성 규칙
- **사실/수치만**: 회사 가이던스·추정치는 "자체 추정/가이던스"로 명시하고 주석. 종목코드 등 불확실하면 확인 요청.
- **Q&A는 구체적으로**: NDR에서 실제 오간 질문을 그대로 살려 수치·고객사·일정 포함(추상적 요약 금지).
- **핵심문구 강조**: summary/deepdive body에서 투자 포인트가 되는 구절을 `**...**`로 감싸면 볼드+밑줄로 렌더.
- **매매콜 금지**: 투자의견/목표가/매수·손절밴드 넣지 않음.
- 임시 빌드 스크립트를 새로 만들지 말고 이 스킬의 제너레이터를 사용. 산출물은 바탕화면에 `.docx` + `.pdf` 동시 저장.

## 폴더 구성
```
ndr-note/
├── SKILL.md
├── README.md                  # 다른 PC 이식·설치 안내
├── assets/
│   └── nexus_logo.png         # 표지 헤더밴드 로고(없으면 회사명 텍스트로 폴백)
├── reference/                 # 디자인 근거 양식(생성에 필수 아님)
│   ├── (템플릿)매매전략 보고서.docx
│   └── 선익시스템 탐방 보고서.pdf
└── scripts/
    ├── build_ndr_note.py      # config.json → docx 제너레이터(포터블 출력)
    ├── convert_to_pdf.py      # docx → pdf (LibreOffice/Word 자동)
    ├── config_amotech.json    # config 스키마 예시(아모텍 실데이터)
    └── requirements.txt
```

## 관련
- 메모리: `reference_tamban_report_format.md`(양식 2종·워크플로우), `project_shadowing.md`(NotebookLM 노트북 운용).
