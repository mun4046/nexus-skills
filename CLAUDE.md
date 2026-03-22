# CLAUDE.md — IPO Demand Forecast Skills

This file provides AI assistants with everything needed to understand, navigate, and contribute to this repository.

---

## Project Overview

**넥서스자산운용 (Nexus Asset Management)** IPO automation skill system. Two Claude AI skills that automate the creation of IPO approval documents (Excel sheets) by reading screenshots and fund summary tables.

**Language**: Korean-language financial domain. Most documentation, variable names for domain concepts, and output data are in Korean.

---

## Repository Structure

```
ipo-demand-forecast-skills/
├── package.json                              # Node.js dependencies
├── skills/
│   ├── ipo-demand-forecast/                  # Skill 1: Demand forecast sheet generation
│   │   ├── SKILL.md                          # Full skill spec (Korean)
│   │   └── scripts/
│   │       ├── read_summary_tables.js        # Reads fund data from Excel summary tables
│   │       ├── create_demand_sheet_v2.js     # Creates demand forecast sheet (current)
│   │       └── create_demand_sheet.js        # Legacy version (deprecated)
│   └── ipo-allocation-subscription/          # Skill 2: Allocation & subscription sheet generation
│       ├── SKILL.md                          # Full skill spec (Korean)
│       └── scripts/
│           ├── create_subscription_sheet_v2.js  # Creates subscription sheet (current)
│           └── create_subscription_sheet.js     # Legacy version (deprecated)
```

**Always prefer `_v2.js` scripts** over their counterparts — v2 uses JSZip for direct XML manipulation enabling precise Excel formatting.

---

## Dependencies

Install with `npm install` (no build step required).

| Package | Purpose |
|---|---|
| `jszip` | Direct ZIP/XML manipulation for Excel files (primary in v2 scripts) |
| `exceljs` | Excel workbook manipulation |
| `xlsx` | Excel file parsing (used in legacy v1 scripts) |
| `xlsx-populate` | Template-based Excel generation |
| `archiver` | ZIP archive creation |
| `unzipper` | ZIP extraction |

---

## Running Scripts

All scripts are invoked directly with Node.js — no build step, no test runner.

### Read fund summary tables
```bash
node skills/ipo-demand-forecast/scripts/read_summary_tables.js \
  "<path/to/집합투자재산_총괄집계표.xlsx>" \
  ["<path/to/벤처기업투자신탁_총괄집계표.xlsx>"]
```
Output: JSON to stdout with `ipoFunds`, `kobenFunds`, `ipoAccount`, `kobenAccount`.

### Create demand forecast sheet
```bash
node skills/ipo-demand-forecast/scripts/create_demand_sheet_v2.js '<JSON_DATA>'
```

### Create subscription sheet
```bash
node skills/ipo-allocation-subscription/scripts/create_subscription_sheet_v2.js '<JSON_DATA>'
```

There are no automated tests. Manual verification is done by opening the generated Excel file.

---

## Skill 1: `ipo-demand-forecast`

**Trigger keywords**: 수요예측, 점검 리스트, 결재서류, IPO, 공모주, 엑셀 시트 생성, 청약

**Workflow:**
1. Collect from user: demand forecast screenshot(s), summary table Excel file(s), target workbook path
2. OCR the screenshot(s) to extract: `stockName`, `unitPrice`, `totalQuantity`, `totalAmount`, `subscriptionPeriod`, `paymentDate`, `participationDate`, `lockUp`, `listingDate`, `accountNumber`, `leadManager`
3. Run `read_summary_tables.js` to get fund list and asset totals
4. Allocate quantities by asset ratio using **floor** (`Math.floor`)
5. Run `create_demand_sheet_v2.js` with assembled JSON
6. Summarize results to user

**Key JSON input for create_demand_sheet_v2.js:**
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
  "ipoFunds": [{"name": "...", "code": "...", "totalAssets": 0}],
  "kobenFunds": [],
  "ipoTotalQuantity": 328000,
  "kobenTotalQuantity": 0
}
```

---

## Skill 2: `ipo-allocation-subscription`

**Trigger keywords**: 배정, 청약, 배분, 수수료, 청약 내역, 배정수량

**Used after** `ipo-demand-forecast` — when allocation results are received.

**Key difference from Skill 1**: Uses **round** (`Math.round`) instead of floor for quantity allocation.

**Workflow:**
1. Collect: allocation result screenshot, summary table Excel, target workbook path
2. OCR screenshot to extract: `stockName`, `ipoAllocatedQty`, `kobenAllocatedQty`, `unitPrice`, `lockUp`, `feeRate`, `totalFee`, `totalPayment`
3. Run `read_summary_tables.js` (shared from ipo-demand-forecast skill)
4. Allocate quantities by asset ratio using **round** (`Math.round`)
5. Calculate fees: `fee = amount * feeRate` (default 1% = 0.01)
6. Run `create_subscription_sheet_v2.js`

**Key JSON input for create_subscription_sheet_v2.js:**
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
  "ipoFunds": [],
  "kobenFunds": [],
  "ipoAllocatedQty": 1136,
  "kobenAllocatedQty": 0,
  "custodianBank": "넥서스-KB증권-우리은행",
  "lockUp": "15일확약",
  "templateSheet": null
}
```

---

## Shared Data: `read_summary_tables.js` Output

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

---

## Business Rules & Conventions

### Quantity Allocation
- **Demand forecast (Skill 1)**: `Math.floor(totalQuantity * ratio)` — conservative floor
- **Subscription (Skill 2)**: `Math.round(allocatedQty * ratio)` — rounded
- Remainder goes to the fund with largest assets to ensure totals match exactly
- Funds with `totalAssets = 0` get quantity = 0, amount = 0 (row still appears)

### Lead Manager Domain Mapping
| Email domain | Securities firm |
|---|---|
| `nhsec.com` | NH투자증권 |
| `miraeasset.com` | 미래에셋증권 |
| `samsungpop.com` | 삼성증권 |
| `kbsec.com` / `kbfg.com` | KB증권 |
| `shinhaninvest.com` | 신한투자증권 |
| `iprovest.com` | 교보증권 |
| `daishin.com` | 대신증권 |
| `hantoosi.com` / `truefriend.com` | 한국투자증권 |
| `shinyoung.com` | 신영증권 |
| `ibks.com` | IBK투자증권 |
| `kiwoom.com` | 키움증권 |
| `sks.co.kr` | SK증권 |

If domain cannot be identified, **ask the user**.

### Sheet Name Convention
Strip company type prefixes `(주)` and stock codes, shorten if needed, append lead manager abbreviation:
- `(주)인벤테라(0007J0)` → `인벤테라_NH`
- Subscription sheets append `_청약`: `인벤테라_NH_청약`

**Lead manager abbreviations**: NH→`NH`, 미래에셋→`미래`, 삼성→`삼성`, KB→`KB`, 신한→`신한`, 교보→`교보`, 대신→`대신`, 한국투자→`한투`, 신영→`신영`, IBK→`IBK`, 키움→`키움`, SK→`sk`

### Lock-up (의무보유 확약) Values
Normalize to: `"미확약"`, `"15일"`, `"1개월"`, `"3개월"`, `"6개월"`

### Venture Enterprise (벤처기업 해당여부)
- If lock-up includes "확약" → typically `"해당"`
- No lock-up → `"미해당"`
- If uncertain, **ask the user**

### Hardcoded Values
- Custodian bank: `넥서스-KB증권-우리은행`
- Writer: `투자운용본부 운용지원팀 최영`
- Default target file: `C:/Users/76167/Desktop/01 펀드 수요예측_청약_통합_2026.xlsx`

### Data Formatting Rules
- Strip commas before parsing numbers: `"16,600"` → `16600`
- Dates in Korean format: `YYYY년 MM월 DD일`
- **Never delete or modify existing sheets** — only add new ones
- Sheet names must not contain special characters (Excel limitation: `\ / * ? [ ] :`)

### Fund Code Mapping (hardcoded in read_summary_tables.js)
| Fund name (partial) | Code |
|---|---|
| 공모주일반사모1호 | 100100 |
| 공모주일반사모2호 | 100200 |
| 공모주일반사모3호 | 100300 |
| 공모주일반사모4호 | 100400 |
| 공모주일반사모5호 | 100500 |
| 멀티일반사모1호 | 100600 |
| 코스닥벤처일반사모1호 | 200100 |
| 코스닥벤처일반사모2호 | 200200 |

---

## Excel Generation Architecture

The v2 scripts use **JSZip** to directly read and write the XLSX file's internal XML — an XLSX file is a ZIP archive containing XML files.

**Why direct XML manipulation?**
Higher-level libraries (ExcelJS, xlsx) don't always support the precise formatting required (specific fonts, colors, border styles, merged cells in complex layouts). Direct XML gives full control.

**Key XML files modified inside the XLSX ZIP:**
- `xl/worksheets/sheetN.xml` — cell data and layout
- `xl/styles.xml` — font, fill, border, and cell format definitions
- `xl/workbook.xml` — sheet registry
- `[Content_Types].xml` — content type declarations
- `xl/_rels/workbook.xml.rels` — workbook relationships

**Style conventions used:**
- Font: 맑은 고딕 (Malgun Gothic)
- Light blue headers: column header rows
- Dark blue section headers: major section titles
- Yellow sum rows: totals
- Green highlights: key result values

---

## Development Notes

- **No linting or formatting tools configured** — maintain consistency with existing code style
- **No test suite** — verify by opening generated `.xlsx` files in Excel
- The `_v1` scripts (`create_demand_sheet.js`, `create_subscription_sheet.js`) are kept for reference but should not be modified or used for new work
- All scripts are standalone — no inter-script imports (except `read_summary_tables.js` being shared conceptually)
- `node_modules/`, `package-lock.json`, `skills.zip`, `.claude/`, `*.tmp` are gitignored

---

## Git Workflow

- Main branch: `master`
- Feature branches follow: `claude/<description>-<id>` pattern
- Remote: `http://local_proxy@127.0.0.1:43301/git/mun4046/ipo-demand-forecast-skills`
