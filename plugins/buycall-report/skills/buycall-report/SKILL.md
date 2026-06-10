---
name: buycall-report
description: >-
  한국 주식 매수콜(매매전략) 보고서를 buy-side 운용역 관점으로 생성한다. 매매내역(종목·종목코드·최초매수일·매수가·전략)을
  입력받아, 최초 매수일 기준 as-of 시점에 작성된 것처럼 1페이지 sell-side 원페이저 + 상세 3~4페이지 Word(.docx)
  보고서를 만든다. pykrx·네이버금융·DART 실데이터를 사용하고, 개조식 어투로 거래별 전략 내러티브를 강조한다.
  트리거: "매수콜 보고서", "매매전략 보고서", "buy call", "종목 보고서 작성", 매매내역 엑셀 첨부.
---

# 매수콜(매매전략) 보고서 생성 스킬

buy-side 운용사 관점의 한국 주식 매수콜/매매전략 보고서를 **증권사 sell-side 리포트 1페이지 양식**으로
재현하고, 2~3페이지의 상세 분석을 덧붙인 Word 보고서를 만든다.

## 언제 쓰나
- 사용자가 매매내역(과거 매수 종목들)을 주고 "그 시점에 작성한 매수 콜 보고서"를 요청할 때.
- 단일 종목 매수 논거를 정형화된 보고서로 정리할 때.

## 입력
종목당 최소: **종목명, 종목코드, 최초매수일, 매수가, (가능하면) 전략 유형**.
매도일·매도가가 있어도 **본문에 노출하지 않는다**(작성일 이후 미래 정보).

## 워크플로우
1. **작성일(as-of) 결정** — 최초 매수일 기준: 추세추종·테마·단기 트레이딩은 **돌파 확인 1영업일 전**, 가치투자는 **2영업일 전**. (휴장일 제외) 그 시점에 알 수 있는 정보만 사용한다.
2. **전략 유형 분류** — 매매내역의 '전략' 또는 가격·거래량 패턴으로 분류: 추세추종(돌파)/테마 모멘텀/단기 트레이딩(데이트레이드)/가치투자(턴어라운드) 등. 전략에 따라 목표·손절·기간·내러티브가 달라진다.
3. **데이터 수집** (`references/data_sources.md` 참조)
   - 가격·52주·거래대금·KOSPI 대비 상대수익률 → `scripts/gather_data.py` 의 `pykrx_snapshot(code, asof)`.
   - 시총·발행주식수·외국인·PER/PBR/BPS·연간실적·동일업종PER·뉴스 → 브라우저(Playwright)로 네이버 금융 메인.
   - 재무상태표·현금흐름(부록) → DART MCP `get_complete_financial_statements`.
   - Peer 멀티플(동종 3종목) → 네이버에서 PER/PBR 수집, ROE ≈ PBR/PER×100.
   - as-of 정합성: 가격은 pykrx 작성일 종가로, PER/PBR 은 (작성일 종가 ÷ EPS·BPS)로 재계산.
4. **내러티브 작성** — 개조식(~함/~임/~했음). 재무지표 나열 지양, **전략 유형에 맞는 논리**로:
   - 추세추종/테마: 거래량 동반 돌파·수급·테마 촉매 중심. 밸류 부담은 손절로 통제한다고 명시.
   - 단기 트레이딩(적자주 등): 펀더멘털보다 테마·수급 베팅임을 솔직히, 작은 비중·짧은 손절 강조.
   - 가치투자: 실적 턴어라운드·저PBR·배당·자산가치 중심, 분할 매수.
5. **보고서 생성** — `D` dict(스키마: `references/format_spec.md`)를 채워 `scripts/gen_report.py` 의 `build(D)` 호출. 또는 JSON 배열로 만들어 `python gen_report.py stocks.json [outdir]`.
6. **검증(선택)** — Word COM 으로 PDF 변환 후 페이지 수·레이아웃 확인(`references/data_sources.md`). 3~4페이지가 정상.

## 핵심 규칙 (반드시 준수)
- 파일명 `종목명_매매전략_YYMMDD.docx`(일자=작성일).
- 개조식 어투, **면책 조항 없음**.
- 작성일 이후의 정보(매도 결과 등) 본문 미노출.
- 수치는 실데이터. Peer 의 실적저점/적자로 인한 PER 왜곡은 각주로 명시(`peer_note`).
- 1페이지 맞춤: 핵심결론 ~3줄, Forecasts ~6행.

## 실행 환경
- `pip install -r requirements.txt` (python-docx, pykrx, pymupdf).
- Windows 한글: `PYTHONUTF8=1 python -X utf8 ...`.
- 로고/브랜드: `assets/` 로고 교체, `scripts/gen_report.py` 상단 컬러 상수 수정으로 다른 회사에 적용.
- 환경변수: `BUYCALL_LOGO`(로고 경로), `BUYCALL_OUT`(출력 폴더).

## 참고 파일
- `scripts/gen_report.py` — 보고서 빌더(`build(D)`).
- `scripts/gather_data.py` — pykrx/네이버 수집 헬퍼.
- `references/format_spec.md` — 레이아웃·규칙·`D` dict 스키마.
- `references/data_sources.md` — 데이터 소스별 수집·검증 방법.
- `assets/` — 로고(`nexus_logo.png`), 빈 템플릿(`template_buycall.docx`).
