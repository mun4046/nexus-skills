/**
 * 수요예측 점검 리스트 시트를 기존 엑셀 파일에 추가한다.
 *
 * Usage:
 *   node create_demand_sheet.js '<JSON_DATA>'
 *
 * JSON_DATA fields:
 *   targetFile, sheetName, stockName, writeDate, baseDate,
 *   leadManager, isVenture, shareType, unitPrice, lockUp,
 *   ipoFunds[], kobenFunds[], ipoTotalQuantity, kobenTotalQuantity
 */

const XLSX = require('xlsx');
const path = require('path');

const data = JSON.parse(process.argv[2]);

const {
    targetFile, sheetName, stockName, writeDate, baseDate,
    leadManager, isVenture, shareType, unitPrice, lockUp,
    ipoFunds, kobenFunds, ipoTotalQuantity, kobenTotalQuantity
} = data;

// ── helpers ──

function allocateByRatio(funds, totalQty) {
    const totalAssets = funds.reduce((s, f) => s + f.totalAssets, 0);
    if (totalAssets === 0) return funds.map(f => ({ ...f, quantity: 0, amount: 0 }));

    // Use ROUND (matching Excel's ROUND function) for proportional allocation
    let allocated = funds.map(f => {
        const ratio = f.totalAssets / totalAssets;
        const qty = Math.round(totalQty * ratio);
        return { ...f, ratio, quantity: qty, amount: qty * unitPrice };
    });

    // Adjust if rounding caused total mismatch
    let sumQty = allocated.reduce((s, f) => s + f.quantity, 0);
    let diff = totalQty - sumQty;
    while (diff !== 0) {
        // Find fund with largest fractional error to adjust
        const errors = allocated.map((f, i) => ({
            i,
            err: (totalQty * f.ratio) - f.quantity
        }));
        if (diff > 0) {
            // Need to add — pick the one most under-allocated
            errors.sort((a, b) => b.err - a.err);
            allocated[errors[0].i].quantity += 1;
            diff -= 1;
        } else {
            // Need to subtract — pick the one most over-allocated
            errors.sort((a, b) => a.err - b.err);
            allocated[errors[0].i].quantity -= 1;
            diff += 1;
        }
    }
    // Recalculate amounts
    allocated.forEach(f => { f.amount = f.quantity * unitPrice; });

    return allocated;
}

function shortFundName(fullName) {
    // "넥서스공모주\r\n일반사모투자신탁제1호" → "넥서스공모주일반사모1호"
    let name = fullName.replace(/\r?\n/g, '');
    name = name.replace('투자신탁', '');
    name = name.replace('제1호', '1호').replace('제2호', '2호').replace('제3호', '3호');
    name = name.replace('코스닥벤처', '코벤');
    return name;
}

function normalizedLockUp(raw) {
    if (!raw) return '미확약';
    const s = raw.replace(/\s/g, '');
    if (s.includes('미확약') || s.includes('미')) return '미확약';
    if (s.includes('15일')) return '15일';
    if (s.includes('1개월')) return '1개월';
    if (s.includes('3개월')) return '3개월';
    if (s.includes('6개월')) return '6개월';
    return raw.trim();
}

// ── Build worksheet ──

const ws = {};
const merges = [];

function setCell(addr, value, type) {
    const t = type || (typeof value === 'number' ? 'n' : 's');
    ws[addr] = { v: value, t };
}

function addMerge(s, e) {
    merges.push(XLSX.utils.decode_range(`${s}:${e}`));
}

// Row 2: Title
setCell('A2', `${stockName} 수요예측 점검 리스트`);
addMerge('A2', 'K2');

// Row 4-9: Metadata
setCell('A4', `작성일 : ${writeDate}`);
setCell('A5', `작성기준일 : ${baseDate}`);
setCell('A6', '작성자 : 투자운용본부 운용지원팀 최영');
setCell('A7', `주관사 : ${leadManager}`);
setCell('A8', `벤처기업 해당여부 : ${isVenture ? '해당' : '미해당'}`);
setCell('A9', shareType || '신주 100%');

// Row 11: Section header
setCell('A11', '*기관투자자 참여수량');

// Rows 12-17: Participation summary
setCell('A12', 'IPO'); addMerge('A12', 'C12');
setCell('D12', ipoTotalQuantity || 0); addMerge('D12', 'E12');
setCell('A13', '코벤'); addMerge('A13', 'C13');
setCell('D13', kobenTotalQuantity || 0); addMerge('D13', 'E13');
setCell('A14', '하이일드'); addMerge('A14', 'C14');
addMerge('D14', 'E14');
setCell('A15', '하이일드'); addMerge('A15', 'C15');
setCell('D15', 'N/A', 's'); addMerge('D15', 'E15');
setCell('A16', '일임'); addMerge('A16', 'C16');
addMerge('D16', 'E16');
setCell('A17', '수요예측 참여 공모가액'); addMerge('A17', 'C17');
setCell('D17', unitPrice); addMerge('D17', 'E17');

// Row 20-21: Table headers
setCell('A20', '펀드정보'); addMerge('A20', 'E20');
setCell('F20', '수요예측 내역'); addMerge('F20', 'J20');
setCell('K20', '비고'); addMerge('K20', 'K21');

setCell('A21', '연번');
setCell('B21', '기관구분');
setCell('C21', '펀드명');
setCell('D21', '펀드코드');
setCell('E21', '총자산 3개월 평균\r\n(FoF제외)');
setCell('F21', '수량\r\n(주)');
setCell('G21', '단가\r\n(원)');
setCell('H21', '금액\r\n(원)');
setCell('I21', '확약여부');
setCell('J21', 'NAV한도내\r\n참여여부');

// ── IPO Fund Rows ──
// Cell addresses are 1-indexed: A22 = Excel row 22
const ipoAllocated = allocateByRatio(ipoFunds || [], ipoTotalQuantity || 0);
let currentExcelRow = 22; // First data row in Excel (1-indexed)

ipoAllocated.forEach((fund, idx) => {
    const row = currentExcelRow + idx;
    setCell(`A${row}`, idx + 1);
    setCell(`B${row}`, '집합투자');
    setCell(`C${row}`, shortFundName(fund.name));
    setCell(`D${row}`, fund.code ? parseInt(fund.code) || fund.code : '');
    setCell(`E${row}`, fund.totalAssets);
    setCell(`F${row}`, fund.quantity);
    setCell(`G${row}`, unitPrice);
    setCell(`H${row}`, fund.amount);
    setCell(`I${row}`, normalizedLockUp(lockUp));
    setCell(`J${row}`, fund.totalAssets > 0 && fund.amount <= fund.totalAssets ? '한도내참여' : (fund.totalAssets === 0 ? '' : '한도이상'));
});

// IPO sum row
const ipoSumExcelRow = currentExcelRow + ipoAllocated.length;
// Sum values
setCell(`E${ipoSumExcelRow}`, ipoAllocated.reduce((s, f) => s + f.totalAssets, 0));
setCell(`F${ipoSumExcelRow}`, ipoAllocated.reduce((s, f) => s + f.quantity, 0));
setCell(`H${ipoSumExcelRow}`, ipoAllocated.reduce((s, f) => s + f.amount, 0));

// ── 코벤 Fund Rows (if any) ──
if (kobenFunds && kobenFunds.length > 0 && kobenTotalQuantity > 0) {
    const kobenAllocated = allocateByRatio(kobenFunds, kobenTotalQuantity);
    const kobenStartExcelRow = ipoSumExcelRow + 2; // Skip a blank row

    kobenAllocated.forEach((fund, idx) => {
        const row = kobenStartExcelRow + idx;
        setCell(`A${row}`, idx + 1);
        setCell(`B${row}`, '코스닥벤처');
        setCell(`C${row}`, shortFundName(fund.name));
        setCell(`D${row}`, fund.code ? parseInt(fund.code) || fund.code : '');
        setCell(`E${row}`, fund.totalAssets);
        setCell(`F${row}`, fund.quantity);
        setCell(`G${row}`, unitPrice);
        setCell(`H${row}`, fund.amount);
        setCell(`I${row}`, normalizedLockUp(lockUp));
        setCell(`J${row}`, fund.totalAssets > 0 && fund.amount <= fund.totalAssets ? '한도내참여' : '');
    });

    // Koben sum row
    const kobenSumExcelRow = kobenStartExcelRow + kobenAllocated.length;
    setCell(`E${kobenSumExcelRow}`, kobenAllocated.reduce((s, f) => s + f.totalAssets, 0));
    setCell(`F${kobenSumExcelRow}`, kobenAllocated.reduce((s, f) => s + f.quantity, 0));
    setCell(`H${kobenSumExcelRow}`, kobenAllocated.reduce((s, f) => s + f.amount, 0));
}

// Set merges and range
ws['!merges'] = merges;

// Calculate the range
let maxRow = 0, maxCol = 0;
Object.keys(ws).forEach(key => {
    if (key.startsWith('!')) return;
    const { r, c } = XLSX.utils.decode_cell(key);
    if (r > maxRow) maxRow = r;
    if (c > maxCol) maxCol = c;
});
ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });

// Column widths (approximate, matching existing sheets)
ws['!cols'] = [
    { wch: 6 },   // A: 연번
    { wch: 10 },  // B: 기관구분
    { wch: 28 },  // C: 펀드명
    { wch: 10 },  // D: 펀드코드
    { wch: 18 },  // E: 총자산
    { wch: 12 },  // F: 수량
    { wch: 10 },  // G: 단가
    { wch: 16 },  // H: 금액
    { wch: 10 },  // I: 확약여부
    { wch: 14 },  // J: NAV한도
    { wch: 10 },  // K: 비고
];

// ── Save to existing workbook ──
let wb;
try {
    wb = XLSX.readFile(targetFile);
} catch (e) {
    console.error(`대상 파일을 열 수 없습니다: ${targetFile}`);
    console.error(e.message);
    process.exit(1);
}

// Check for duplicate sheet name
if (wb.SheetNames.includes(sheetName)) {
    console.error(`이미 동일한 시트명이 존재합니다: ${sheetName}`);
    console.error('다른 시트명을 사용하거나, 기존 시트를 삭제 후 재실행하세요.');
    process.exit(1);
}

// Add sheet before the last sheet (결재방)
const insertIdx = wb.SheetNames.length - 1; // Before 결재방
XLSX.utils.book_append_sheet(wb, ws, sheetName);

// Move to correct position (before 결재방)
const names = wb.SheetNames;
const newSheetIdx = names.indexOf(sheetName);
if (newSheetIdx !== insertIdx) {
    names.splice(newSheetIdx, 1);
    names.splice(insertIdx, 0, sheetName);
}

XLSX.writeFile(wb, targetFile);

console.log(`✓ 시트 "${sheetName}" 생성 완료`);
console.log(`  파일: ${targetFile}`);
console.log(`  종목: ${stockName}`);
console.log(`  주관사: ${leadManager}`);
console.log(`  참여가격: ${unitPrice.toLocaleString()}원`);
if (ipoTotalQuantity) {
    console.log(`  IPO 참여수량: ${ipoTotalQuantity.toLocaleString()}주`);
    (ipoAllocated || []).forEach(f => {
        console.log(`    - ${shortFundName(f.name)}: ${f.quantity.toLocaleString()}주 (${f.amount.toLocaleString()}원)`);
    });
}
if (kobenTotalQuantity) {
    console.log(`  코벤 참여수량: ${kobenTotalQuantity.toLocaleString()}주`);
}
