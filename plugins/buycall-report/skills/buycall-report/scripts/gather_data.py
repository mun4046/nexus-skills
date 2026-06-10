# -*- coding: utf-8 -*-
"""
매수콜 보고서용 데이터 수집 헬퍼.

핵심:
  - pykrx 의 시총/PER/PBR/지수/외국인 함수는 환경에 따라 KRX OTP 차단으로 비어 올 수 있음.
    OHLCV·종목명만 안정적 → as-of 종가/52주/거래대금/상대수익률은 OHLCV로 계산하고,
    KOSPI 대비 상대수익률은 069500(KODEX200)을 프록시로 사용한다.
  - 시총·발행주식수·외국인·PER/PBR/BPS·연간실적·동일업종PER 은 네이버 금융에서 가져온다.
  - 재무상태표·현금흐름(부록)은 DART(전자공시) get_complete_financial_statements 로 가져온다.

Windows 한글 깨짐 방지: PYTHONUTF8=1 python -X utf8 gather_data.py 319660 20260604
"""
import sys, io, re, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from datetime import datetime, timedelta

def pykrx_snapshot(code, asof, market_proxy="069500"):
    """asof(YYYYMMDD) 기준 종가/52주/60일 거래대금/KOSPI 대비 상대수익률."""
    from pykrx import stock
    d=datetime.strptime(asof,"%Y%m%d")
    s1=(d-timedelta(days=400)).strftime("%Y%m%d")
    o=stock.get_market_ohlcv(s1,asof,code)
    px=o['종가']; mkt=stock.get_market_ohlcv(s1,asof,market_proxy)['종가']
    ret=lambda s,n: None if len(s)<=n else s.iloc[-1]/s.iloc[-1-n]-1
    rel={}
    for lbl,n in [('1M',21),('3M',63),('6M',126),('12M',248)]:
        rs,rk=ret(px,n),ret(mkt,n)
        rel[lbl]= None if rs is None or rk is None else round((rs-rk)*100,1)
    return dict(
        close=int(px.iloc[-1]),
        high52=int(o['고가'].tail(248).max()),
        low52=int(o['저가'].tail(248).min()),
        val60_eok=round((o['종가']*o['거래량']).tail(60).mean()/1e8),
        rel=rel,
    )

_NUM=lambda s: re.sub(r'[^0-9.\-]','',s) if s else s
def naver_summary(code):
    """네이버 금융 메인에서 요약 지표를 정규식으로 추출(서버렌더 HTML)."""
    url=f"https://finance.naver.com/item/main.naver?code={code}"
    req=urllib.request.Request(url, headers={'User-Agent':'Mozilla/5.0'})
    html=urllib.request.urlopen(req, timeout=15).read().decode('euc-kr','replace')
    def by_id(i):
        m=re.search(rf'id="{i}"[^>]*>([^<]+)<', html); return m.group(1).strip() if m else None
    # 투자정보 박스(#tab_con1) 텍스트만 추출 → 시총/주식수/외국인 등을 여기서 파싱
    # 시총/발행주식수/외국인/52주/연간실적/뉴스 등 풍부한 요약은 브라우저(Playwright)로
    # 네이버 메인(#tab_con1, .section.cop_analysis)을 읽는 것이 가장 안정적이다(SKILL.md 참조).
    # requests 경로에서는 id 기반 per/pbr/eps 만 신뢰성 있게 추출한다.
    return dict(per=by_id('_per'), pbr=by_id('_pbr'), eps=by_id('_eps'))

if __name__=='__main__':
    code=sys.argv[1] if len(sys.argv)>1 else '319660'
    asof=sys.argv[2] if len(sys.argv)>2 else datetime.today().strftime('%Y%m%d')
    print(f"[pykrx as-of {asof}] {code}")
    try:
        for k,v in pykrx_snapshot(code,asof).items(): print(f"  {k}: {v}")
    except Exception as e:
        print("  pykrx err:", e)
    print(f"[naver summary] {code}")
    try:
        for k,v in naver_summary(code).items(): print(f"  {k}: {v}")
    except Exception as e:
        print("  naver err:", e)
