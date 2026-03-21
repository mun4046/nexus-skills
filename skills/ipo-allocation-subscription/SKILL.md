---
name: ipo-allocation-subscription
description: IPO 배정 결과 수령 후 결재서류(배분 및 청약 내역 엑셀 시트)를 자동 생성하는 스킬. 배정 결과 스크린샷과 총괄집계표를 입력받아 기존 통합 엑셀 파일에 청약 시트를 추가한다. 배정, 청약, 배분, 수수료, 청약 내역, 배정수량 등의 키워드가 나오면 이 스킬을 사용할 것. 이 스킬은 수요예측 참여 이후 배정 결과를 받은 단계에서 사용한다.
---

# IPO 배분 및 청약 내역 자동 생성

넥서스자산운용의 IPO 배정 결과에 따른 청약 결재서류를 자동 생성한다.
수요예측 참여(ipo-demand-forecast) 이후, 배정 결과를 받은 후 사용하는 스킬이다.

## 워크플로우

### Step 1: 사용자로부터 입력 수집

1. **배정 결과 스크린샷** (이미지)
   - 증권사 홈페이지의 배정 결과 화면 캡처
   - 포함 정보: 종목명, 배정수량, 확정공모가, 청약수수료, 총납입금액 등
2. **총괄집계표 엑셀 파일** (수요예측 때 사용한 것과 동일하거나 업데이트된 버전)
3. **통합 엑셀 파일 경로** (기본값: `C:/Users/76167/Desktop/01 펀드 수요예측_청약_통합_2026.xlsx`)

### Step 2: 스크린샷에서 데이터 추출 (OCR)

배정 결과 스크린샷에서 아래 필드를 추출한다:

| 필드명 | 변수명 | 예시 |
|--------|--------|------|
| 종목명 | `stockName` | 한패스 주식회사 |
| 배정수량 (집합투자분) | `ipoAllocatedQty` | 421 주 |
| 배정수량 (코벤분) | `kobenAllocatedQty` | 0 주 (없으면 0) |
| 확정공모가액 | `unitPrice` | 19,000 원 |
| 의무보유/확약 | `lockUp` | 15일확약 |
| 청약수수료율 | `feeRate` | 0.01 (1%) |
| 청약수수료 | `totalFee` | 79,990 원 (검증용) |
| 총납입금액 | `totalPayment` | 8,078,990 원 (검증용) |

**배정수량이 IPO/코벤 구분 없이 하나만 있는 경우:**
- 기존 수요예측 시트에서 IPO/코벤 참여 여부를 확인하여 판단
- 또는 사용자에게 확인

### Step 3: 총괄집계표에서 펀드 데이터 읽기

`ipo-demand-forecast` 스킬의 `read_summary_tables.js`를 공유 사용한다:
```bash
node <ipo-demand-forecast-skill-path>/scripts/read_summary_tables.js "<집합투자재산_총괄집계표_경로>" "<벤처기업투자신탁_총괄집계표_경로>"
```

### Step 4: 펀드별 배분 계산

**배분 로직 (청약 시트 전용):**
1. 비율 계산: `ratio = fund.totalAssets / sumOfAllFundAssets`
2. 수량 배분: `quantity = Math.round(ipoAllocatedQty * ratio)`
3. 나머지 조정 (총합 일치)
4. 금액: `amount = quantity * unitPrice`
5. 청약수수료: `fee = amount * feeRate` (기본 1%, 즉 0.01)
6. 금액+수수료: `totalWithFee = amount + fee`

### Step 5: 엑셀 시트 생성

`scripts/create_subscription_sheet_v2.js` 스크립트를 실행한다.

```bash
node <skill-path>/scripts/create_subscription_sheet_v2.js '<JSON_DATA>'
```

JSON_DATA 구조:
```json
{
  "targetFile": "C:/Users/76167/Desktop/01 펀드 수요예측_청약_통합_2026.xlsx",
  "sheetName": "교보20호스팩_청약",
  "stockName": "교보20호스팩",
  "subscriptionDate": "2026년 3월 23일",
  "baseDate": "2026년 3월 15일",
  "leadManager": "교보증권",
  "isVenture": false,
  "shareType": "신주 100%",
  "unitPrice": 2000,
  "feeRate": 0.01,
  "ipoFunds": [...],
  "kobenFunds": [...],
  "ipoAllocatedQty": 1136,
  "kobenAllocatedQty": 214,
  "custodianBank": "넥서스-KB증권-우리은행",
  "lockUp": "15일확약",
  "templateSheet": null
}
```

## 엑셀 시트 구조 (시트 B: 배분 및 청약 내역)

청약 시트는 수요예측 시트와 구조가 다르다:
- 데이터가 **B열**부터 시작 (A열이 아님)
- 타이틀이 `B2:N2` 범위로 병합

### 헤더 블록
- B2: "{종목명} 배분 및 청약 내역" (merged B2:N2)
- B4: "청약일 : {청약일}"
- B5: "작성기준일 : {기준일}"
- B6: "작성자 : 투자운용본부 운용지원팀 최영"
- B7: "주관사 : {주관사}"
- B8: "벤처기업 해당여부 : {해당/미해당}"
- B9: "{신주 비율}"

### IPO 배분 테이블
- B11 또는 B12: "<펀드별 배분 내역>"
- 헤더 행 1: B="펀드정보" (merged), G="최종배정" (merged)
- 헤더 행 2: No(B), 기관구분(C), 펀드명(D), 펀드코드(E), 총자산 3개월 평균(F), 비율%(G), 수량(H), 단가(I), 금액(J), 청약수수료(K), 금액+수수료(L), 확약(M), 수탁은행(N)
- 데이터 행: 펀드별 1행씩
- 합계 행: F, G(=1), H, I, J, K, L 합계
- "집합투자분 배정수량" 행: 총 배정수량 입력값

### 코벤 배분 테이블 (해당 시)
- 동일 구조로 코벤 펀드 배분

### 최하단 합계
- "총 배정수량" 행: IPO + 코벤 총합계

## 수탁은행 정보

현재 모든 펀드의 수탁은행은 동일: `넥서스-KB증권-우리은행`
변경 시 사용자가 알려줄 것.

## 시트 이름 규칙

수요예측 시트 이름에 `_청약` 접미사:
- 수요예측: `교보20호스팩` → 청약: `교보20호스팩_청약`
- 수요예측: `인벤테라_NH` → 청약: `인벤테라_NH_청약`
