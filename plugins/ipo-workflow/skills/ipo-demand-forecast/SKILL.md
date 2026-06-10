---
name: ipo-demand-forecast
description: IPO 수요예측 참여 후 결재서류(엑셀 점검 리스트 시트)를 자동 생성하는 스킬. 수요예측 참여내역 스크린샷 이미지와 총괄집계표 엑셀을 입력받아 기존 통합 엑셀 파일에 새 시트를 추가한다. 수요예측, 점검 리스트, 결재서류, IPO, 공모주, 엑셀 시트 생성, 청약 등의 키워드가 나오면 이 스킬을 사용할 것. 사용자가 수요예측 관련 이미지를 첨부하거나 총괄집계표를 언급하면 반드시 트리거할 것.
---

# IPO 수요예측 점검 리스트 자동 생성

넥서스자산운용의 IPO 수요예측 참여 결재서류를 자동으로 생성한다.

## 워크플로우

### Step 1: 사용자로부터 입력 수집

사용자에게 다음 3가지를 요청한다:

1. **수요예측 참여내역 스크린샷** (이미지 1장 이상)
   - 증권사 홈페이지의 "수요예측 신청내역 확인서" 캡처
2. **총괄집계표 엑셀 파일** (1~2개)
   - `집합투자재산_총괄집계표.xlsx` → IPO 펀드 데이터
   - `벤처기업투자신탁_총괄집계표.xlsx` → 코벤 펀드 데이터 (해당 시)
3. **통합 엑셀 파일 경로** (시트를 추가할 대상)
   - 기본값: `C:/Users/76167/Desktop/01 펀드 수요예측_청약_통합_2026.xlsx`

### Step 2: 스크린샷 이미지에서 데이터 추출 (OCR)

이미지를 Read 도구로 읽은 뒤, 아래 필드를 정확히 추출한다:

| 필드명 | 변수명 | 예시 |
|--------|--------|------|
| 종목명 | `stockName` | (주)인벤테라 |
| 참여가격 | `unitPrice` | 16,600 |
| 수량 (총 참여수량) | `totalQuantity` | 361,000 |
| 총참여금액 | `totalAmount` | 5,992,600,000 |
| 청약기간 | `subscriptionPeriod` | 2026.03.23 ~ 2026.03.24 |
| 납입일 | `paymentDate` | 2026.03.26 |
| 참여일자 | `participationDate` | 2026.03.12 |
| 의무보유/확약 | `lockUp` | 1개월 확약 / 미 확약 / 15일 확약 |
| 상장예정일 | `listingDate` | 2026.04.02 |
| 참여계좌번호 | `accountNumber` | 209-01-560276 |
| 참여기관상호 | `institutionName` | 넥서스자산운용 주식회사 |
| 참여기관구분 | `institutionType` | 집합투자업자(집합투자자산) (사모) |
| IPO담당자 | `ipoManager` | 윤서정 |
| IPO담당자 전화번호 | `ipoManagerPhone` | 02-768-7590 |
| 주관사 (IPO담당자 e-mail 도메인에서 유추) | `leadManager` | NH투자증권 |

**주관사 판별 규칙:**
- `nhsec.com` → NH투자증권
- `miraeasset.com` → 미래에셋증권
- `samsungpop.com` → 삼성증권
- `kbsec.com` / `kbfg.com` → KB증권
- `shinhaninvest.com` → 신한투자증권
- `iprovest.com` → 교보증권
- `daishin.com` → 대신증권
- `hantoosi.com` / `truefriend.com` → 한국투자증권
- `shinyoung.com` → 신영증권
- `ibks.com` → IBK투자증권
- `kiwoom.com` → 키움증권
- `sks.co.kr` → SK증권
- 판별이 안 되면 사용자에게 물어볼 것

**참여내역에 IPO/코벤 구분이 없는 경우:**
- 스크린샷이 2장이면: 각각 IPO용, 코벤용으로 사용자에게 확인
- 총괄집계표가 2개(집합투자재산 + 벤처기업투자신탁)면 코벤 참여가 있는 것
- 총괄집계표가 1개(집합투자재산만)면 IPO만

### Step 3: 총괄집계표에서 펀드 데이터 읽기

`scripts/read_summary_tables.js` 스크립트를 실행하여 총괄집계표에서 펀드 목록과 자산총액을 읽는다.

```bash
node <skill-path>/scripts/read_summary_tables.js "<집합투자재산_총괄집계표_경로>" "<벤처기업투자신탁_총괄집계표_경로>"
```

두 번째 인자는 선택적이다 (코벤 참여가 없으면 생략).

출력은 JSON:
```json
{
  "ipoFunds": [
    {"name": "넥서스공모주일반사모투자신탁제1호", "code": "100100", "totalAssets": 3806373241},
    {"name": "넥서스공모주일반사모투자신탁제2호", "code": "100200", "totalAssets": 20111335575}
  ],
  "kobenFunds": [
    {"name": "넥서스코스닥벤처일반사모투자신탁제1호", "code": "200100", "totalAssets": 621893050}
  ],
  "ipoAccount": "1025-69211-01",
  "kobenAccount": "1025-69211-02"
}
```

### Step 4: 펀드별 수량 배분 계산

**배분 로직:**
1. 각 펀드의 자산총액 비율 계산: `ratio = fund.totalAssets / sumOfAllFundAssets`
2. 수량 배분: `quantity = Math.floor(totalQuantity * ratio)`
3. 나머지 조정: 총합이 totalQuantity와 일치하도록 가장 큰 펀드에 나머지 추가
4. 금액 계산: `amount = quantity * unitPrice`

**자산총액이 0인 펀드 처리:**
- 수량 0, 금액 0으로 설정 (행은 생성하되 비워둠)
- 예: 넥서스멀티일반사모1호는 총자산 0이면 수량/금액 0

### Step 5: 엑셀 시트 생성

`scripts/create_demand_sheet_v2.js` 스크립트를 실행한다.

```bash
node <skill-path>/scripts/create_demand_sheet_v2.js '<JSON_DATA>'
```

JSON_DATA 구조:
```json
{
  "targetFile": "C:/Users/76167/Desktop/01 펀드 수요예측_청약_통합_2026.xlsx",
  "sheetName": "인벤테라_NH",
  "stockName": "인벤테라",
  "writeDate": "2026년 03월 13일",
  "baseDate": "2026년 03월 12일",
  "leadManager": "NH투자증권",
  "isVenture": false,
  "shareType": "신주 100%",
  "unitPrice": 16600,
  "lockUp": "1개월",
  "ipoFunds": [...],
  "kobenFunds": [...],
  "ipoTotalQuantity": 328000,
  "kobenTotalQuantity": 7000
}
```

### Step 6: 결과 확인

생성된 시트 내용을 사용자에게 요약해서 보여준다:
- 종목명, 주관사, 참여가격
- 펀드별 배분 수량/금액 테이블
- 총 참여수량/금액

사용자가 확인 후 수정 요청하면 스크립트를 재실행한다.

## 중요 규칙

- 숫자 데이터는 반드시 쉼표 제거 후 숫자로 변환 (예: "16,600" → 16600)
- 날짜는 한국어 형식 유지 (YYYY년 MM월 DD일)
- 시트 이름에 특수문자 금지 (엑셀 제한)
- 기존 시트는 절대 삭제/수정하지 않음 — 새 시트 추가만
- 의무보유 확약 표현 통일: "미확약", "15일", "1개월", "3개월", "6개월"
- 벤처기업 해당여부는 의무보유에 "확약"이 있으면 보통 "해당", 없으면 "미해당" — 확실하지 않으면 사용자에게 확인

## 참고: 시트 이름 규칙

종목명에서 괄호와 코드를 제거하고 축약:
- `(주)인벤테라(0007J0)` → `인벤테라`
- `엔에이치기업인수목적33호 (주)(0130H0)` → `엔에이치기업인수목적33호` 또는 더 짧게

주관사 약칭:
- NH투자증권 → `NH`
- 미래에셋증권 → `미래`
- 삼성증권 → `삼성`
- KB증권 → `KB`
- 신한투자증권 → `신한`
- 교보증권 → `교보`
- 대신증권 → `대신`
- 한국투자증권 → `한투`
- 신영증권 → `신영`
- IBK투자증권 → `IBK`
- 키움증권 → `키움`
- SK증권 → `sk`

최종 시트 이름 예: `인벤테라_NH`, `교보20호스팩`, `신한제17호스팩`
