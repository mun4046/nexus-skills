/**
 * 총괄집계표 엑셀 파일에서 펀드 데이터를 읽어 JSON으로 출력한다.
 *
 * Usage:
 *   node read_summary_tables.js "<집합투자재산_총괄집계표.xlsx>" ["<벤처기업투자신탁_총괄집계표.xlsx>"]
 *
 * Output: JSON to stdout
 */

const XLSX = require('xlsx');
const path = require('path');

function findXlsxModule() {
    // Try multiple possible locations for the xlsx module
    const possiblePaths = [
        path.resolve(__dirname, '../../node_modules/xlsx'),
        path.resolve(__dirname, '../../../node_modules/xlsx'),
        'xlsx'
    ];
    for (const p of possiblePaths) {
        try { return require(p); } catch (e) { /* continue */ }
    }
    throw new Error('xlsx module not found. Run: npm install xlsx');
}

function readIpoSummary(filePath) {
    const wb = XLSX.readFile(filePath);
    // First sheet is the data sheet (not 증빙예시)
    const dataSheetName = wb.SheetNames.find(n => !n.includes('증빙'));
    if (!dataSheetName) throw new Error('집합투자재산 데이터 시트를 찾을 수 없습니다.');

    const ws = wb.Sheets[dataSheetName];
    const range = XLSX.utils.decode_range(ws['!ref']);

    const funds = [];
    let accountNumber = '';
    let stockName = '';

    // Find account number and stock name from header area
    for (let r = 0; r <= Math.min(range.e.r, 20); r++) {
        for (let c = 0; c <= Math.min(range.e.c, 10); c++) {
            const cell = ws[XLSX.utils.encode_cell({r, c})];
            if (!cell) continue;
            const val = String(cell.v || '');
            if (val.includes('계좌번호')) {
                // Account number is in the next column(s)
                for (let cc = c + 1; cc <= Math.min(range.e.c, 10); cc++) {
                    const acCell = ws[XLSX.utils.encode_cell({r, c: cc})];
                    if (acCell && acCell.v) {
                        accountNumber = String(acCell.v).trim();
                        break;
                    }
                }
            }
            if (val.includes('참여종목명') || val.includes('종목명')) {
                for (let cc = c + 1; cc <= Math.min(range.e.c, 10); cc++) {
                    const snCell = ws[XLSX.utils.encode_cell({r, c: cc})];
                    if (snCell && snCell.v) {
                        stockName = String(snCell.v).trim();
                        break;
                    }
                }
            }
        }
    }

    // Find fund data rows - look for rows with "구분" header
    let headerRow = -1;
    for (let r = 15; r <= Math.min(range.e.r, 30); r++) {
        const cellA = ws[XLSX.utils.encode_cell({r, c: 0})];
        if (cellA && String(cellA.v || '').includes('구분')) {
            headerRow = r;
            break;
        }
    }

    if (headerRow === -1) {
        // Try alternate: look for sequential number 1
        for (let r = 18; r <= Math.min(range.e.r, 35); r++) {
            const cellA = ws[XLSX.utils.encode_cell({r, c: 0})];
            if (cellA && cellA.v === 1) {
                headerRow = r - 1;
                break;
            }
        }
    }

    // Read fund rows (start from headerRow + 1)
    for (let r = headerRow + 1; r <= range.e.r; r++) {
        const cellA = ws[XLSX.utils.encode_cell({r, c: 0})];
        if (!cellA || cellA.v === undefined) continue;

        // Check if this is a number (fund index) or "합계"
        if (String(cellA.v).includes('합계')) break;
        if (typeof cellA.v !== 'number') continue;

        const nameCell = ws[XLSX.utils.encode_cell({r, c: 1})];
        const fundName = nameCell ? String(nameCell.v || '').replace(/\r?\n/g, '') : '';
        if (!fundName) continue;

        // Total assets: look in columns E, F, G for the largest non-zero value
        // Column G is usually 자산총액(원)
        let totalAssets = 0;
        for (let c = 4; c <= 7; c++) {
            const assetCell = ws[XLSX.utils.encode_cell({r, c})];
            if (assetCell && typeof assetCell.v === 'number' && assetCell.v > totalAssets) {
                totalAssets = assetCell.v;
            }
        }

        // Extract fund code from name or use known mapping
        const code = extractFundCode(fundName);

        funds.push({
            name: fundName,
            code: code,
            totalAssets: totalAssets
        });
    }

    // ── FoF (피투자펀드) 차감 처리 ──
    // "주금납입능력" 행을 찾아 전체 조정 후 자산총액을 읽고,
    // 각 펀드별 원래 자산총액에서 역산한다.
    let adjustedTotalAssets = 0;
    for (let r = 0; r <= range.e.r; r++) {
        for (let c = 0; c <= Math.min(range.e.c, 10); c++) {
            const cell = ws[XLSX.utils.encode_cell({r, c})];
            if (!cell) continue;
            const val = String(cell.v || '');
            if (val.includes('주금납입능력')) {
                // Find the numeric value in the same row
                for (let cc = c + 1; cc <= Math.min(range.e.c, 10); cc++) {
                    const valCell = ws[XLSX.utils.encode_cell({r, c: cc})];
                    if (valCell && typeof valCell.v === 'number' && valCell.v > 0) {
                        adjustedTotalAssets = valCell.v;
                        break;
                    }
                }
                break;
            }
        }
        if (adjustedTotalAssets > 0) break;
    }

    if (adjustedTotalAssets > 0) {
        // 주금납입능력 = 전체 자산총액 - 피투자펀드 자산총액
        // 펀드별로 FoF 차감액을 구해야 함
        // 보유 집합투자증권 섹션에서 펀드별 차감액을 읽는다
        const fofByFund = {}; // fundName -> total FoF deduction

        // Find the FoF section: "개별 위탁재산이 보유한 집합투자증권" header
        // (must be specific to avoid matching disclaimer text at top of sheet)
        let fofSectionFound = false;
        for (let r = 0; r <= range.e.r && !fofSectionFound; r++) {
            for (let c = 0; c <= Math.min(range.e.c, 10); c++) {
                const cell = ws[XLSX.utils.encode_cell({r, c})];
                if (!cell) continue;
                const val = String(cell.v || '');
                if (!val.includes('개별 위탁재산') || !val.includes('보유한 집합투자증권')) continue;

                fofSectionFound = true;
                // Found FoF header — read entries below
                for (let rr = r + 1; rr <= Math.min(range.e.r, r + 20); rr++) {
                    const numCell = ws[XLSX.utils.encode_cell({r: rr, c: 0})];
                    if (!numCell || typeof numCell.v !== 'number') continue;

                    // Column B: parent fund name
                    const parentCell = ws[XLSX.utils.encode_cell({r: rr, c: 1})];
                    if (!parentCell) continue;
                    const parentName = String(parentCell.v || '').replace(/\r?\n/g, '');

                    // Match to our fund list — use fund number (제1호, 제2호) for disambiguation
                    const matchedFund = funds.find(f => {
                        const cleanParent = parentName.replace(/\s/g, '');
                        const cleanFund = f.name.replace(/\s/g, '');
                        // Must match the specific fund number (1호/2호/3호)
                        const numMatch = cleanFund.match(/(제?\d+호)/);
                        if (numMatch) {
                            return cleanParent.includes(numMatch[1]) &&
                                   (cleanParent.includes('공모주') === cleanFund.includes('공모주')) &&
                                   (cleanParent.includes('멀티') === cleanFund.includes('멀티')) &&
                                   (cleanParent.includes('코벤') === cleanFund.includes('코벤') || cleanParent.includes('코스닥') === cleanFund.includes('코스닥'));
                        }
                        return cleanParent === cleanFund;
                    });

                    if (!matchedFund) continue;

                    // FoF asset amount is in column H (index 7)
                    const amtCell = ws[XLSX.utils.encode_cell({r: rr, c: 7})];
                    const fofAmount = (amtCell && typeof amtCell.v === 'number') ? amtCell.v : 0;
                    if (fofAmount > 0) {
                        fofByFund[matchedFund.name] = (fofByFund[matchedFund.name] || 0) + fofAmount;
                    }
                }
                break; // Only process first FoF section
            }
        }

        // Apply deductions
        funds.forEach(f => {
            if (fofByFund[f.name]) {
                f.totalAssetsBeforeFoF = f.totalAssets;
                f.fofDeduction = fofByFund[f.name];
                f.totalAssets = f.totalAssets - fofByFund[f.name];
            }
        });

        // Verify: sum of adjusted assets should equal 주금납입능력
        const sumAdjusted = funds.reduce((s, f) => s + f.totalAssets, 0);
        if (Math.abs(sumAdjusted - adjustedTotalAssets) > 1) {
            console.error(`[경고] 조정 후 자산총액 합계(${sumAdjusted})와 주금납입능력(${adjustedTotalAssets}) 불일치`);
        }
    }

    return { funds, accountNumber, stockName };
}

function readKobenSummary(filePath) {
    const wb = XLSX.readFile(filePath);
    const dataSheetName = wb.SheetNames.find(n => !n.includes('증빙'));
    if (!dataSheetName) throw new Error('벤처기업투자신탁 데이터 시트를 찾을 수 없습니다.');

    const ws = wb.Sheets[dataSheetName];
    const range = XLSX.utils.decode_range(ws['!ref']);

    const funds = [];
    let accountNumber = '';

    // Find account number
    for (let r = 0; r <= Math.min(range.e.r, 20); r++) {
        for (let c = 0; c <= Math.min(range.e.c, 15); c++) {
            const cell = ws[XLSX.utils.encode_cell({r, c})];
            if (!cell) continue;
            const val = String(cell.v || '');
            if (val.includes('계좌번호')) {
                for (let cc = c + 1; cc <= Math.min(range.e.c, 15); cc++) {
                    const acCell = ws[XLSX.utils.encode_cell({r, c: cc})];
                    if (acCell && acCell.v) {
                        accountNumber = String(acCell.v).trim();
                        break;
                    }
                }
            }
        }
    }

    // Find fund data - look for "펀드명" header
    let headerRow = -1;
    for (let r = 15; r <= Math.min(range.e.r, 25); r++) {
        const cellA = ws[XLSX.utils.encode_cell({r, c: 0})];
        if (cellA && String(cellA.v || '').includes('펀드명')) {
            headerRow = r;
            break;
        }
    }

    if (headerRow === -1) return { funds, accountNumber };

    // Skip sub-header rows (벤처기업투자신탁 has 2 header rows)
    let dataStart = headerRow + 1;
    // Check if next row is also a header
    const nextCell = ws[XLSX.utils.encode_cell({r: dataStart, c: 0})];
    if (nextCell && String(nextCell.v || '').includes('벤처기업')) {
        dataStart = headerRow + 2;
    }

    for (let r = dataStart; r <= range.e.r; r++) {
        const cellA = ws[XLSX.utils.encode_cell({r, c: 0})];
        if (!cellA || cellA.v === undefined) continue;

        const nameVal = String(cellA.v || '').replace(/\r?\n/g, '');
        if (nameVal.includes('합') || nameVal.includes('작성 요령')) break;
        if (!nameVal || nameVal.trim() === '') continue;

        // 재산총액 is in column E
        let totalAssets = 0;
        const assetCell = ws[XLSX.utils.encode_cell({r, c: 4})];
        if (assetCell && typeof assetCell.v === 'number') {
            totalAssets = assetCell.v;
        }

        const code = extractFundCode(nameVal);

        funds.push({
            name: nameVal,
            code: code,
            totalAssets: totalAssets
        });
    }

    return { funds, accountNumber };
}

function extractFundCode(fundName) {
    // Known fund code mapping
    const codeMap = {
        '공모주일반사모투자신탁제1호': '100100',
        '공모주일반사모1호': '100100',
        '공모주일반사모투자신탁제2호': '100200',
        '공모주일반사모2호': '100200',
        '멀티일반사모투자신탁제1호': '400100',
        '멀티일반사모1호': '400100',
        '코스닥벤처일반사모투자신탁제1호': '200100',
        '코스닥벤처일반사모1호': '200100',
        '코벤일반사모투자신탁제1호': '200100',
        '코벤일반사모1호': '200100',
    };

    for (const [key, code] of Object.entries(codeMap)) {
        if (fundName.includes(key.replace('일반사모', '').replace('투자신탁', '')) || fundName.includes(key)) {
            return code;
        }
    }

    // Try to extract from known patterns
    if (fundName.includes('공모주') && fundName.includes('1호')) return '100100';
    if (fundName.includes('공모주') && fundName.includes('2호')) return '100200';
    if (fundName.includes('멀티') && fundName.includes('1호')) return '400100';
    if (fundName.includes('코스닥벤처') || fundName.includes('코벤')) return '200100';

    return '';
}

// Main
const args = process.argv.slice(2);
if (args.length < 1) {
    console.error('Usage: node read_summary_tables.js "<집합투자재산.xlsx>" ["<벤처기업투자신탁.xlsx>"]');
    process.exit(1);
}

const result = {
    ipoFunds: [],
    kobenFunds: [],
    ipoAccount: '',
    kobenAccount: '',
    stockName: ''
};

try {
    const ipoData = readIpoSummary(args[0]);
    result.ipoFunds = ipoData.funds;
    result.ipoAccount = ipoData.accountNumber;
    result.stockName = ipoData.stockName;
} catch (e) {
    console.error('집합투자재산 총괄집계표 읽기 오류:', e.message);
    process.exit(1);
}

if (args.length >= 2 && args[1]) {
    try {
        const kobenData = readKobenSummary(args[1]);
        result.kobenFunds = kobenData.funds;
        result.kobenAccount = kobenData.accountNumber;
    } catch (e) {
        console.error('벤처기업투자신탁 총괄집계표 읽기 오류:', e.message);
    }
}

console.log(JSON.stringify(result, null, 2));
