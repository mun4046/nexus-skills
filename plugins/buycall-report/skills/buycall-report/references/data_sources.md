# 데이터 소스 가이드

작성일(as-of) 시점 기준의 **실데이터**로 채운다. 세 소스를 조합한다.

## 1) pykrx — 가격·52주·상대수익률·거래대금 (작성일 기준)
- `scripts/gather_data.py` 의 `pykrx_snapshot(code, asof)` 사용.
- **주의**: pykrx 의 `get_market_cap`·`get_market_fundamental`·`get_index_ohlcv`·외국인 함수는 환경에 따라 **KRX OTP 차단으로 빈 값**이 올 수 있음. `get_market_ohlcv`(가격)·`get_market_ticker_name` 만 안정적.
  - 시총 = (작성일 종가 × 발행주식수)로 직접 계산(발행주식수는 네이버/DART).
  - KOSPI 대비 상대수익률은 지수 함수 대신 **069500(KODEX200) OHLCV**를 프록시로 사용.
- Windows 한글 깨짐 방지: `PYTHONUTF8=1 python -X utf8 ...` (pykrx 의 cp949 캐시 버그 회피).

## 2) 네이버 금융 — 시총·발행주식수·외국인·PER/PBR/BPS·연간실적·동일업종PER·뉴스
- **권장(가장 안정적)**: 브라우저(Playwright)로 `https://finance.naver.com/item/main.naver?code={코드}` 접속 후
  - `#tab_con1`(투자정보 박스) innerText → 시가총액·상장주식수·외국인소진율·52주 최고/최저·PER/EPS·PBR/BPS·배당수익률·동일업종 PER.
  - `.section.cop_analysis` innerText → 연간/분기 실적표(매출·영업이익·순이익·영업이익률·ROE·부채비율·EPS·PER·BPS·PBR·배당).
  - 뉴스 헤드라인 → 테마/촉매 파악(특히 모멘텀·테마 트레이드 내러티브에 활용).
- Peer 멀티플도 같은 페이지에서 종목코드만 바꿔 수집. 같은 origin 이므로 한 번의 `browser_evaluate` 안에서 `fetch('/item/main.naver?code=...')` 로 여러 Peer 를 동시에 받아 `<em id="_per">`, `<em id="_pbr">`, `<em id="_eps">` 를 파싱하면 효율적. ROE ≈ PBR/PER×100.
- **보조(브라우저 불가 시)**: `gather_data.py` 의 `naver_summary(code)` → requests 로 per/pbr/eps 만.
- 네이버 요약치는 '현재' 기준이라 작성일과 며칠 차이가 날 수 있음 → 가격은 pykrx 작성일 종가로 덮어쓰고, PER/PBR 은 (작성일 종가 ÷ EPS·BPS)로 재계산하면 as-of 정합성 확보.

## 3) DART(전자공시) — 재무상태표·현금흐름 (부록)
- MCP 의 `get_corpcode(corp_name)` → 상장사 `corp_code`(stock_code 일치하는 것 선택).
- `get_complete_financial_statements(corp_code, bsns_year, reprt_code='11011'(사업보고서), fs_div='CFS')` → 자산총계·부채총계·자본총계·영업활동현금흐름·유형자산취득(=capex)·차입금·현금 등.
  - FCF ≈ 영업활동현금흐름 − 유형자산 취득. 순현금 = 현금 − 차입금.
- 부록 BS 는 네이버 BPS×발행주식수(자본)·부채비율(부채)로 근사 가능(DART 없이도).

## 검증(선택, Windows+MS Word)
- Word COM 으로 docx→PDF 변환 후 페이지 수 확인, pymupdf 로 PNG 렌더해 레이아웃 육안 점검.
  - `Documents.Open(src,$false,$true)`(읽기전용) → `ExportAsFixedFormat(pdf,17)` 가 `SaveAs([ref],17)` 보다 안정적.
  - 저장 PermissionError = docx 가 Word 에 열려 있음 → 해당 WINWORD 종료 후 재시도.
