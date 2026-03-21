/**
 * 배분 및 청약 내역 시트를 기존 엑셀 파일에 추가한다.
 *
 * Usage:
 *   node create_subscription_sheet.js '<JSON_DATA>'
 */

const XLSX = require('xlsx');

const data = JSON.parse(process.argv[2]);

const {
    targetFile, sheetName, stockName, subscriptionDate, baseDate,
    leadManager, isVenture, shareType, unitPrice, feeRate,
    ipoFunds, kobenFunds, ipoAllocatedQty, kobenAllocatedQty,
    custodianBank
} = data;

const fee = feeRate || 0.01;
const bank = custodianBank || '넥서스-KB증권-우리은행';

// ── helpers ──

function allocateByRatio(funds, totalQty) {
    const totalAssets = funds.reduce((s, f) => s + f.totalAssets, 0);
    if (totalAssets === 0) return funds.map(f => ({ ...f, ratio: 0, quantity: 0, amount: 0, fee: 0, totalWithFee: 0 }));

    let allocated = funds.map(f => {
        const ratio = f.totalAssets / totalAssets;
        const qty = Math.round(totalQty * ratio);
        return { ...f, ratio, quantity: qty };
    });

    // Adjust rounding
    let sumQty = allocated.reduce((s, f) => s + f.quantity, 0);
    let diff = totalQty - sumQty;
    while (diff !== 0) {
        const errors = allocated.map((f, i) => ({
            i,
            err: (totalQty * f.ratio) - f.quantity
        }));
        if (diff > 0) {
            errors.sort((a, b) => b.err - a.err);
            allocated[errors[0].i].quantity += 1;
            diff -= 1;
        } else {
            errors.sort((a, b) => a.err - b.err);
            allocated[errors[0].i].quantity -= 1;
            diff += 1;
        }
    }

    // Calculate amounts
    allocated.forEach(f => {
        f.amount = f.quantity * unitPrice;
        f.fee = Math.round(f.amount * fee);
        f.totalWithFee = f.amount + f.fee;
    });

    return allocated;
}

function shortFundName(fullName) {
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

// ── Header block (column B) ──
setCell('B2', `${stockName} 배분 및 청약 내역`);
addMerge('B2', 'N2');

setCell('B4', `청약일 : ${subscriptionDate}`);
setCell('B5', `작성기준일 : ${baseDate}`);
setCell('B6', '작성자 : 투자운용본부 운용지원팀 최영');
setCell('B7', `주관사 : ${leadManager}`);
setCell('B8', `벤처기업 해당여부 : ${isVenture ? '해당' : '미해당'}`);
setCell('B9', shareType || '신주 100%');

// ── IPO Allocation Table ──
let row = 11; // Start row for section label (matches existing sheets)

setCell(`B${row}`, '<펀드별 배분 내역>');

// Table header row 1
row++;
setCell(`B${row}`, '펀드정보'); addMerge(`B${row}`, `F${row}`);
setCell(`G${row}`, '최종배정'); addMerge(`G${row}`, `L${row}`);

// Table header row 2
row++;
setCell(`B${row}`, 'No');
setCell(`C${row}`, '기관구분');
setCell(`D${row}`, '펀드명');
setCell(`E${row}`, '펀드코드');
setCell(`F${row}`, '총자산 3개월 평균\r\n(원)');
setCell(`G${row}`, '비율\r\n(%)');
setCell(`H${row}`, '수량\r\n(주) ');
setCell(`I${row}`, '단가\r\n(원)');
setCell(`J${row}`, '금액\r\n(원)');
setCell(`K${row}`, '청약수수료');
setCell(`L${row}`, '금액+수수료');
setCell(`M${row}`, '확약');
setCell(`N${row}`, '수탁은행');

// IPO fund data rows
const ipoAllocated = allocateByRatio(ipoFunds || [], ipoAllocatedQty || 0);
row++;
const ipoDataStartRow = row;

ipoAllocated.forEach((fund, idx) => {
    const r = ipoDataStartRow + idx;
    setCell(`B${r}`, idx + 1);
    setCell(`C${r}`, '집합투자');
    setCell(`D${r}`, shortFundName(fund.name));
    setCell(`E${r}`, fund.code ? parseInt(fund.code) || fund.code : '');
    setCell(`F${r}`, fund.totalAssets);
    setCell(`G${r}`, fund.ratio);
    setCell(`H${r}`, fund.quantity);
    setCell(`I${r}`, unitPrice);
    setCell(`J${r}`, fund.amount);
    setCell(`K${r}`, fund.fee);
    setCell(`L${r}`, fund.totalWithFee);
    if (fund.quantity > 0) {
        setCell(`M${r}`, normalizedLockUp(data.lockUp));
    }
    setCell(`N${r}`, bank);
    row = r + 1;
});

// IPO totals row
const ipoTotalRow = row;
setCell(`F${ipoTotalRow}`, ipoAllocated.reduce((s, f) => s + f.totalAssets, 0));
setCell(`G${ipoTotalRow}`, 1);
setCell(`H${ipoTotalRow}`, ipoAllocated.reduce((s, f) => s + f.quantity, 0));
setCell(`I${ipoTotalRow}`, unitPrice);
setCell(`J${ipoTotalRow}`, ipoAllocated.reduce((s, f) => s + f.amount, 0));
setCell(`K${ipoTotalRow}`, ipoAllocated.reduce((s, f) => s + f.fee, 0));
setCell(`L${ipoTotalRow}`, ipoAllocated.reduce((s, f) => s + f.totalWithFee, 0));
row++;

// "집합투자분 배정수량" summary row
setCell(`B${row}`, '집합투자분 배정수량');
setCell(`E${row}`, ipoAllocatedQty || 0);
const ipoTotalAssets = ipoAllocated.reduce((s, f) => s + f.totalAssets, 0);
const ipoTotalAmount = ipoAllocated.reduce((s, f) => s + f.amount, 0);
if (ipoTotalAssets > 0) {
    setCell(`G${row}`, ipoTotalAmount / ipoTotalAssets);
}
row++;

// ── 코벤 Allocation Table (if any) ──
if (kobenFunds && kobenFunds.length > 0 && kobenAllocatedQty > 0) {
    row++; // blank row

    // Table header row 1
    setCell(`B${row}`, '펀드정보'); addMerge(`B${row}`, `F${row}`);
    setCell(`G${row}`, '최종배정'); addMerge(`G${row}`, `L${row}`);
    row++;

    // Table header row 2
    setCell(`B${row}`, 'No');
    setCell(`C${row}`, '기관구분');
    setCell(`D${row}`, '펀드명');
    setCell(`E${row}`, '펀드코드');
    setCell(`F${row}`, '총자산 3개월 평균\r\n(원)');
    setCell(`G${row}`, '비율\r\n(%)');
    setCell(`H${row}`, '수량\r\n(주) ');
    setCell(`I${row}`, '단가\r\n(원)');
    setCell(`J${row}`, '금액\r\n(원)');
    setCell(`K${row}`, '청약수수료');
    setCell(`L${row}`, '금액+수수료');
    setCell(`M${row}`, '확약');
    setCell(`N${row}`, '수탁은행');
    row++;

    const kobenAllocated = allocateByRatio(kobenFunds, kobenAllocatedQty);
    const kobenDataStartRow = row;

    kobenAllocated.forEach((fund, idx) => {
        const r = kobenDataStartRow + idx;
        setCell(`B${r}`, idx + 1);
        setCell(`C${r}`, '집합투자');
        setCell(`D${r}`, shortFundName(fund.name));
        setCell(`E${r}`, fund.code ? parseInt(fund.code) || fund.code : '');
        setCell(`F${r}`, fund.totalAssets);
        setCell(`G${r}`, fund.ratio);
        setCell(`H${r}`, fund.quantity);
        setCell(`I${r}`, unitPrice);
        setCell(`J${r}`, fund.amount);
        setCell(`K${r}`, fund.fee);
        setCell(`L${r}`, fund.totalWithFee);
        if (fund.quantity > 0) {
            setCell(`M${r}`, normalizedLockUp(data.lockUp));
        }
        setCell(`N${r}`, bank);
        row = r + 1;
    });

    // Koben totals (if multiple funds, otherwise same as data row)
    // "코스닥벤처분 배정수량"
    setCell(`B${row}`, '코스닥벤처분 배정수량');
    setCell(`E${row}`, kobenAllocatedQty);
    const kobenTotalAssets = kobenAllocated.reduce((s, f) => s + f.totalAssets, 0);
    const kobenTotalAmount = kobenAllocated.reduce((s, f) => s + f.amount, 0);
    if (kobenTotalAssets > 0) {
        setCell(`G${row}`, kobenTotalAmount / kobenTotalAssets);
    }
    row++;

    // Grand total row
    row++; // blank row
    setCell(`B${row}`, '총 배정수량');
    const grandQty = (ipoAllocatedQty || 0) + kobenAllocatedQty;
    const grandAmount = ipoTotalAmount + kobenAllocated.reduce((s, f) => s + f.amount, 0);
    const grandFee = ipoAllocated.reduce((s, f) => s + f.fee, 0) + kobenAllocated.reduce((s, f) => s + f.fee, 0);
    const grandTotal = ipoAllocated.reduce((s, f) => s + f.totalWithFee, 0) + kobenAllocated.reduce((s, f) => s + f.totalWithFee, 0);
    setCell(`H${row}`, grandQty);
    setCell(`I${row}`, unitPrice);
    setCell(`J${row}`, grandAmount);
    setCell(`K${row}`, grandFee);
    setCell(`L${row}`, grandTotal);
} else {
    // No 코벤 — just grand total = IPO total
    row++;
    setCell(`H${row}`, ipoAllocatedQty || 0);
    setCell(`J${row}`, ipoTotalAmount);
    setCell(`K${row}`, ipoAllocated.reduce((s, f) => s + f.fee, 0));
    setCell(`L${row}`, ipoAllocated.reduce((s, f) => s + f.totalWithFee, 0));
}

// ── Finalize ──
ws['!merges'] = merges;

let maxRow = 0, maxCol = 0;
Object.keys(ws).forEach(key => {
    if (key.startsWith('!')) return;
    const { r, c } = XLSX.utils.decode_cell(key);
    if (r > maxRow) maxRow = r;
    if (c > maxCol) maxCol = c;
});
ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });

ws['!cols'] = [
    { wch: 3 },   // A: empty
    { wch: 6 },   // B: No
    { wch: 10 },  // C: 기관구분
    { wch: 28 },  // D: 펀드명
    { wch: 10 },  // E: 펀드코드
    { wch: 18 },  // F: 총자산
    { wch: 12 },  // G: 비율
    { wch: 10 },  // H: 수량
    { wch: 10 },  // I: 단가
    { wch: 14 },  // J: 금액
    { wch: 12 },  // K: 수수료
    { wch: 14 },  // L: 금액+수수료
    { wch: 8 },   // M: 확약
    { wch: 22 },  // N: 수탁은행
];

// ── Save ──
let wb;
try {
    wb = XLSX.readFile(targetFile);
} catch (e) {
    console.error(`대상 파일을 열 수 없습니다: ${targetFile}`);
    process.exit(1);
}

if (wb.SheetNames.includes(sheetName)) {
    console.error(`이미 동일한 시트명이 존재합니다: ${sheetName}`);
    process.exit(1);
}

// Insert before 결재방 sheet
const insertIdx = wb.SheetNames.length - 1;
XLSX.utils.book_append_sheet(wb, ws, sheetName);
const names = wb.SheetNames;
const newIdx = names.indexOf(sheetName);
if (newIdx !== insertIdx) {
    names.splice(newIdx, 1);
    names.splice(insertIdx, 0, sheetName);
}

XLSX.writeFile(wb, targetFile);

// Summary
console.log(`✓ 시트 "${sheetName}" 생성 완료`);
console.log(`  파일: ${targetFile}`);
console.log(`  종목: ${stockName}`);
console.log(`  주관사: ${leadManager}`);
console.log(`  확정공모가: ${unitPrice.toLocaleString()}원`);
console.log(`  수수료율: ${(fee * 100).toFixed(1)}%`);
if (ipoAllocatedQty) {
    console.log(`  IPO 배정수량: ${ipoAllocatedQty.toLocaleString()}주`);
    ipoAllocated.forEach(f => {
        console.log(`    - ${shortFundName(f.name)}: ${f.quantity.toLocaleString()}주, 금액 ${f.amount.toLocaleString()}원, 수수료 ${f.fee.toLocaleString()}원`);
    });
}
if (kobenAllocatedQty) {
    console.log(`  코벤 배정수량: ${kobenAllocatedQty.toLocaleString()}주`);
}
