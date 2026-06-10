/**
 * 배분 및 청약 내역 시트를 기존 엑셀 파일에 추가한다.
 * JSZip으로 직접 조작하여 서식(스타일)을 포함한 시트를 생성한다.
 *
 * Usage:
 *   node create_subscription_sheet_v2.js '<JSON_DATA>'
 */

const JSZip = require('jszip');
const fs = require('fs');

const data = JSON.parse(process.argv[2]);

const {
    targetFile, templateSheet, sheetName, stockName, subscriptionDate, baseDate,
    leadManager, isVenture, shareType, unitPrice, feeRate,
    ipoFunds, kobenFunds, ipoAllocatedQty, kobenAllocatedQty,
    custodianBank, lockUp
} = data;

const fee = feeRate || 0.01;
const bank = custodianBank || '넥서스-KB증권-우리은행';

// ── Allocation logic ──

function allocateByRatio(funds, totalQty) {
    const totalAssets = funds.reduce((s, f) => s + f.totalAssets, 0);
    if (totalAssets === 0) return funds.map(f => ({ ...f, ratio: 0, quantity: 0, amount: 0, fee: 0, totalWithFee: 0 }));

    let allocated = funds.map(f => {
        const ratio = f.totalAssets / totalAssets;
        const qty = Math.round(totalQty * ratio);
        return { ...f, ratio, quantity: qty };
    });

    let sumQty = allocated.reduce((s, f) => s + f.quantity, 0);
    let diff = totalQty - sumQty;
    while (diff !== 0) {
        const errors = allocated.map((f, i) => ({ i, err: (totalQty * f.ratio) - f.quantity }));
        if (diff > 0) { errors.sort((a, b) => b.err - a.err); allocated[errors[0].i].quantity += 1; diff--; }
        else { errors.sort((a, b) => a.err - b.err); allocated[errors[0].i].quantity -= 1; diff++; }
    }
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

// ── XML helpers ──

function colToLetter(c) {
    let s = '';
    while (c > 0) { c--; s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26); }
    return s;
}

function cellXml(col, row, value, styleIdx) {
    const ref = `${colToLetter(col)}${row}`;
    const sAttr = styleIdx !== undefined ? ` s="${styleIdx}"` : '';
    if (typeof value === 'number') {
        return `<x:c r="${ref}"${sAttr}><x:v>${value}</x:v></x:c>`;
    } else if (typeof value === 'string' && value !== '') {
        const escaped = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/\r\n/g, '&#10;').replace(/\n/g, '&#10;');
        return `<x:c r="${ref}"${sAttr} t="inlineStr"><x:is><x:t>${escaped}</x:t></x:is></x:c>`;
    }
    return `<x:c r="${ref}"${sAttr}/>`;
}

function rowXml(rowNum, cells, height) {
    const htAttr = height ? ` ht="${height}" customHeight="1"` : '';
    return `<x:row r="${rowNum}"${htAttr} spans="2:14">${cells.join('')}</x:row>`;
}

// ── Style injection ──

function injectStyles(stylesXml) {
    // Parse existing counts
    const fontCount = parseInt((stylesXml.match(/<x:fonts count="(\d+)"/) || [, '0'])[1]);
    const fillCount = parseInt((stylesXml.match(/<x:fills count="(\d+)"/) || [, '0'])[1]);
    const borderCount = parseInt((stylesXml.match(/<x:borders count="(\d+)"/) || [, '0'])[1]);
    const xfCount = parseInt((stylesXml.match(/<x:cellXfs count="(\d+)"/) || [, '0'])[1]);

    // New fonts to add
    const newFonts = [
        // fontId = fontCount+0: Bold 14pt (title)
        `<x:font><x:name val="맑은 고딕"/><x:sz val="14"/><x:b/><x:color rgb="ff000000"/></x:font>`,
        // fontId = fontCount+1: Bold 11pt (headers)
        `<x:font><x:name val="맑은 고딕"/><x:sz val="11"/><x:b/><x:color rgb="ff000000"/></x:font>`,
        // fontId = fontCount+2: Bold 11pt white (section header on dark bg)
        `<x:font><x:name val="맑은 고딕"/><x:sz val="11"/><x:b/><x:color rgb="ffFFFFFF"/></x:font>`,
        // fontId = fontCount+3: Normal 10pt (data)
        `<x:font><x:name val="맑은 고딕"/><x:sz val="10"/><x:color rgb="ff000000"/></x:font>`,
        // fontId = fontCount+4: Bold 10pt (sum)
        `<x:font><x:name val="맑은 고딕"/><x:sz val="10"/><x:b/><x:color rgb="ff000000"/></x:font>`,
    ];

    // New fills to add
    const newFills = [
        // fillId = fillCount+0: Light blue (column headers)
        `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="ffD6E4F0"/><x:bgColor rgb="ffD6E4F0"/></x:patternFill></x:fill>`,
        // fillId = fillCount+1: Dark blue (section headers)
        `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="ff4472C4"/><x:bgColor rgb="ff4472C4"/></x:patternFill></x:fill>`,
        // fillId = fillCount+2: Light yellow (sum rows)
        `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="ffFFF2CC"/><x:bgColor rgb="ffFFF2CC"/></x:patternFill></x:fill>`,
    ];

    // New borders to add
    const newBorders = [
        // borderId = borderCount+0: Thin all sides
        `<x:border><x:left style="thin"><x:color rgb="ff000000"/></x:left><x:right style="thin"><x:color rgb="ff000000"/></x:right><x:top style="thin"><x:color rgb="ff000000"/></x:top><x:bottom style="thin"><x:color rgb="ff000000"/></x:bottom></x:border>`,
        // borderId = borderCount+1: Medium bottom (sum row)
        `<x:border><x:left style="thin"><x:color rgb="ff000000"/></x:left><x:right style="thin"><x:color rgb="ff000000"/></x:right><x:top style="thin"><x:color rgb="ff000000"/></x:top><x:bottom style="medium"><x:color rgb="ff000000"/></x:bottom></x:border>`,
    ];

    const f = {
        title: fontCount,       // bold 14pt
        header: fontCount + 1,  // bold 11pt
        headerWh: fontCount + 2,// bold 11pt white
        data: fontCount + 3,    // normal 10pt
        sum: fontCount + 4,     // bold 10pt
    };
    const fl = {
        lightBlue: fillCount,
        darkBlue: fillCount + 1,
        yellow: fillCount + 2,
    };
    const b = {
        thin: borderCount,
        medBot: borderCount + 1,
    };

    // New cellXf entries (each maps to a style index starting at xfCount)
    const newXfs = [
        // S.TITLE (xfCount+0): bold 14pt, left, no border
        `<x:xf numFmtId="0" fontId="${f.title}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><x:alignment horizontal="left" vertical="center"/></x:xf>`,
        // S.INFO (xfCount+1): normal 10pt, left, no border
        `<x:xf numFmtId="0" fontId="${f.data}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><x:alignment horizontal="left" vertical="center"/></x:xf>`,
        // S.SECTION_HDR (xfCount+2): bold white, dark blue bg, bordered, center
        `<x:xf numFmtId="0" fontId="${f.headerWh}" fillId="${fl.darkBlue}" borderId="${b.thin}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="center" vertical="center"/></x:xf>`,
        // S.COL_HDR (xfCount+3): bold, light blue bg, bordered, center, wrap
        `<x:xf numFmtId="0" fontId="${f.header}" fillId="${fl.lightBlue}" borderId="${b.thin}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="center" vertical="center" wrapText="1"/></x:xf>`,
        // S.DATA_TEXT (xfCount+4): normal, bordered, left
        `<x:xf numFmtId="0" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="left" vertical="center"/></x:xf>`,
        // S.DATA_NUM (xfCount+5): normal, bordered, right, #,##0
        `<x:xf numFmtId="3" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.DATA_PCT (xfCount+6): normal, bordered, right, percent
        `<x:xf numFmtId="10" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.DATA_CENTER (xfCount+7): normal, bordered, center
        `<x:xf numFmtId="0" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="center" vertical="center"/></x:xf>`,
        // S.SUM_NUM (xfCount+8): bold, bordered med bottom, yellow bg, #,##0
        `<x:xf numFmtId="3" fontId="${f.sum}" fillId="${fl.yellow}" borderId="${b.medBot}" xfId="0" applyFont="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.SUM_PCT (xfCount+9): bold, bordered med bottom, yellow bg, percent
        `<x:xf numFmtId="10" fontId="${f.sum}" fillId="${fl.yellow}" borderId="${b.medBot}" xfId="0" applyFont="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.SUMMARY_LABEL (xfCount+10): bold 10pt, left, no border
        `<x:xf numFmtId="0" fontId="${f.sum}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><x:alignment horizontal="left" vertical="center"/></x:xf>`,
        // S.SUMMARY_NUM (xfCount+11): bold 10pt, right, #,##0, no border
        `<x:xf numFmtId="3" fontId="${f.sum}" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.SUMMARY_PCT (xfCount+12): bold 10pt, right, percent, no border
        `<x:xf numFmtId="10" fontId="${f.sum}" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.EMPTY_BORDER (xfCount+13): bordered, empty
        `<x:xf numFmtId="0" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyBorder="1"/>`,
    ];

    // Inject into XML
    let xml = stylesXml;

    // Update fonts
    xml = xml.replace(/<x:fonts count="\d+"/, `<x:fonts count="${fontCount + newFonts.length}"`);
    xml = xml.replace(/<\/x:fonts>/, newFonts.join('') + '</x:fonts>');

    // Update fills
    xml = xml.replace(/<x:fills count="\d+"/, `<x:fills count="${fillCount + newFills.length}"`);
    xml = xml.replace(/<\/x:fills>/, newFills.join('') + '</x:fills>');

    // Update borders
    xml = xml.replace(/<x:borders count="\d+"/, `<x:borders count="${borderCount + newBorders.length}"`);
    xml = xml.replace(/<\/x:borders>/, newBorders.join('') + '</x:borders>');

    // Update cellXfs
    xml = xml.replace(/<x:cellXfs count="\d+"/, `<x:cellXfs count="${xfCount + newXfs.length}"`);
    xml = xml.replace(/<\/x:cellXfs>/, newXfs.join('') + '</x:cellXfs>');

    // Return style indices map
    return {
        xml,
        S: {
            TITLE: xfCount,
            INFO: xfCount + 1,
            SECTION_HDR: xfCount + 2,
            COL_HDR: xfCount + 3,
            DATA_TEXT: xfCount + 4,
            DATA_NUM: xfCount + 5,
            DATA_PCT: xfCount + 6,
            DATA_CENTER: xfCount + 7,
            SUM_NUM: xfCount + 8,
            SUM_PCT: xfCount + 9,
            SUMMARY_LABEL: xfCount + 10,
            SUMMARY_NUM: xfCount + 11,
            SUMMARY_PCT: xfCount + 12,
            EMPTY_BORDER: xfCount + 13,
        }
    };
}

// ── Build sheet ──

async function main() {
    const buf = fs.readFileSync(targetFile);
    const zip = await JSZip.loadAsync(buf);

    // Parse workbook.xml
    let wbXml = await zip.file('xl/workbook.xml').async('string');
    const sheetRegex = /<(?:x:)?sheet\s+([^>]*?)\/?\s*>/g;
    const sheets = [];
    let match;
    while ((match = sheetRegex.exec(wbXml)) !== null) {
        const attrs = match[1];
        const nameMatch = attrs.match(/name="([^"]*)"/);
        const ridMatch = attrs.match(/r:id="([^"]*)"/);
        const idMatch = attrs.match(/sheetId="(\d+)"/);
        if (nameMatch && ridMatch && idMatch) {
            sheets.push({ name: nameMatch[1], rId: ridMatch[1], sheetId: parseInt(idMatch[1]), fullMatch: match[0] });
        }
    }

    // Check duplicate
    if (sheets.find(s => s.name === sheetName)) {
        console.error(`Sheet "${sheetName}" already exists`);
        process.exit(1);
    }

    // Parse rels
    let relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const relRegex = /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*\/?>/g;
    const rels = [];
    while ((match = relRegex.exec(relsXml)) !== null) {
        rels.push({ id: match[1], target: match[2] });
    }

    // Inject styles
    let stylesXml = await zip.file('xl/styles.xml').async('string');
    const { xml: newStylesXml, S } = injectStyles(stylesXml);
    zip.file('xl/styles.xml', newStylesXml);

    // ── Build sheet content ──
    const rows = [];
    const mergeCells = [];
    function addMerge(start, end) { mergeCells.push(`<x:mergeCell ref="${start}:${end}"/>`); }

    // Row 2: Title
    rows.push(rowXml(2, [cellXml(2, 2, `${stockName} 배분 및 청약 내역`, S.TITLE)]));
    addMerge('B2', 'N2');

    // Rows 4-9: Header info
    rows.push(rowXml(4, [cellXml(2, 4, `청약일 : ${subscriptionDate}`, S.INFO)]));
    rows.push(rowXml(5, [cellXml(2, 5, `작성기준일 : ${baseDate}`, S.INFO)]));
    rows.push(rowXml(6, [cellXml(2, 6, '작성자 : 투자운용본부 운용지원팀 최영', S.INFO)]));
    rows.push(rowXml(7, [cellXml(2, 7, `주관사 : ${leadManager}`, S.INFO)]));
    rows.push(rowXml(8, [cellXml(2, 8, `벤처기업 해당여부 : ${isVenture ? '해당' : '미해당'}`, S.INFO)]));
    rows.push(rowXml(9, [cellXml(2, 9, shareType || '신주 100%', S.INFO)]));

    // ── IPO Section ──
    let row = 11;
    rows.push(rowXml(row, [cellXml(2, row, '<펀드별 배분 내역>', S.INFO)]));
    row++;

    // Section header: 펀드정보 / 최종배정
    rows.push(rowXml(row, [
        cellXml(2, row, '펀드정보', S.SECTION_HDR),
        cellXml(3, row, '', S.SECTION_HDR),
        cellXml(4, row, '', S.SECTION_HDR),
        cellXml(5, row, '', S.SECTION_HDR),
        cellXml(6, row, '', S.SECTION_HDR),
        cellXml(7, row, '최종배정', S.SECTION_HDR),
        cellXml(8, row, '', S.SECTION_HDR),
        cellXml(9, row, '', S.SECTION_HDR),
        cellXml(10, row, '', S.SECTION_HDR),
        cellXml(11, row, '', S.SECTION_HDR),
        cellXml(12, row, '', S.SECTION_HDR),
        cellXml(13, row, '', S.SECTION_HDR),
        cellXml(14, row, '', S.SECTION_HDR),
    ]));
    addMerge(`B${row}`, `F${row}`);
    addMerge(`G${row}`, `L${row}`);
    row++;

    // Column headers
    rows.push(rowXml(row, [
        cellXml(2, row, 'No', S.COL_HDR),
        cellXml(3, row, '기관구분', S.COL_HDR),
        cellXml(4, row, '펀드명', S.COL_HDR),
        cellXml(5, row, '펀드코드', S.COL_HDR),
        cellXml(6, row, '총자산 3개월 평균\r\n(원)', S.COL_HDR),
        cellXml(7, row, '비율\r\n(%)', S.COL_HDR),
        cellXml(8, row, '수량\r\n(주)', S.COL_HDR),
        cellXml(9, row, '단가\r\n(원)', S.COL_HDR),
        cellXml(10, row, '금액\r\n(원)', S.COL_HDR),
        cellXml(11, row, '청약수수료', S.COL_HDR),
        cellXml(12, row, '금액+수수료', S.COL_HDR),
        cellXml(13, row, '확약', S.COL_HDR),
        cellXml(14, row, '수탁은행', S.COL_HDR),
    ], 32));
    row++;

    // IPO data rows
    const ipoAllocated = allocateByRatio(ipoFunds || [], ipoAllocatedQty || 0);
    const lockUpStr = normalizedLockUp(lockUp);

    ipoAllocated.forEach((fund, idx) => {
        rows.push(rowXml(row, [
            cellXml(2, row, idx + 1, S.DATA_CENTER),
            cellXml(3, row, '집합투자', S.DATA_CENTER),
            cellXml(4, row, shortFundName(fund.name), S.DATA_TEXT),
            cellXml(5, row, fund.code ? parseInt(fund.code) || fund.code : '', S.DATA_CENTER),
            cellXml(6, row, fund.totalAssets, S.DATA_NUM),
            cellXml(7, row, fund.ratio, S.DATA_PCT),
            cellXml(8, row, fund.quantity, S.DATA_NUM),
            cellXml(9, row, unitPrice, S.DATA_NUM),
            cellXml(10, row, fund.amount, S.DATA_NUM),
            cellXml(11, row, fund.fee, S.DATA_NUM),
            cellXml(12, row, fund.totalWithFee, S.DATA_NUM),
            cellXml(13, row, fund.quantity > 0 ? lockUpStr : '', S.DATA_CENTER),
            cellXml(14, row, bank, S.DATA_TEXT),
        ]));
        row++;
    });

    // IPO totals row
    rows.push(rowXml(row, [
        cellXml(2, row, '', S.EMPTY_BORDER),
        cellXml(3, row, '', S.EMPTY_BORDER),
        cellXml(4, row, '', S.EMPTY_BORDER),
        cellXml(5, row, '', S.EMPTY_BORDER),
        cellXml(6, row, ipoAllocated.reduce((s, f) => s + f.totalAssets, 0), S.SUM_NUM),
        cellXml(7, row, 1, S.SUM_PCT),
        cellXml(8, row, ipoAllocated.reduce((s, f) => s + f.quantity, 0), S.SUM_NUM),
        cellXml(9, row, unitPrice, S.SUM_NUM),
        cellXml(10, row, ipoAllocated.reduce((s, f) => s + f.amount, 0), S.SUM_NUM),
        cellXml(11, row, ipoAllocated.reduce((s, f) => s + f.fee, 0), S.SUM_NUM),
        cellXml(12, row, ipoAllocated.reduce((s, f) => s + f.totalWithFee, 0), S.SUM_NUM),
        cellXml(13, row, '', S.EMPTY_BORDER),
        cellXml(14, row, '', S.EMPTY_BORDER),
    ]));
    row++;

    // "집합투자분 배정수량"
    const ipoTotalAssets = ipoAllocated.reduce((s, f) => s + f.totalAssets, 0);
    const ipoTotalAmount = ipoAllocated.reduce((s, f) => s + f.amount, 0);
    rows.push(rowXml(row, [
        cellXml(2, row, '집합투자분 배정수량', S.SUMMARY_LABEL),
        cellXml(5, row, ipoAllocatedQty || 0, S.SUMMARY_NUM),
        cellXml(7, row, ipoTotalAssets > 0 ? ipoTotalAmount / ipoTotalAssets : 0, S.SUMMARY_PCT),
    ]));
    row++;

    // ── 코벤 Section ──
    if (kobenFunds && kobenFunds.length > 0 && kobenAllocatedQty > 0) {
        row++; // blank row

        // Section header
        rows.push(rowXml(row, [
            cellXml(2, row, '펀드정보', S.SECTION_HDR),
            cellXml(3, row, '', S.SECTION_HDR),
            cellXml(4, row, '', S.SECTION_HDR),
            cellXml(5, row, '', S.SECTION_HDR),
            cellXml(6, row, '', S.SECTION_HDR),
            cellXml(7, row, '최종배정', S.SECTION_HDR),
            cellXml(8, row, '', S.SECTION_HDR),
            cellXml(9, row, '', S.SECTION_HDR),
            cellXml(10, row, '', S.SECTION_HDR),
            cellXml(11, row, '', S.SECTION_HDR),
            cellXml(12, row, '', S.SECTION_HDR),
            cellXml(13, row, '', S.SECTION_HDR),
            cellXml(14, row, '', S.SECTION_HDR),
        ]));
        addMerge(`B${row}`, `F${row}`);
        addMerge(`G${row}`, `L${row}`);
        row++;

        // Column headers
        rows.push(rowXml(row, [
            cellXml(2, row, 'No', S.COL_HDR),
            cellXml(3, row, '기관구분', S.COL_HDR),
            cellXml(4, row, '펀드명', S.COL_HDR),
            cellXml(5, row, '펀드코드', S.COL_HDR),
            cellXml(6, row, '총자산 3개월 평균\r\n(원)', S.COL_HDR),
            cellXml(7, row, '비율\r\n(%)', S.COL_HDR),
            cellXml(8, row, '수량\r\n(주)', S.COL_HDR),
            cellXml(9, row, '단가\r\n(원)', S.COL_HDR),
            cellXml(10, row, '금액\r\n(원)', S.COL_HDR),
            cellXml(11, row, '청약수수료', S.COL_HDR),
            cellXml(12, row, '금액+수수료', S.COL_HDR),
            cellXml(13, row, '확약', S.COL_HDR),
            cellXml(14, row, '수탁은행', S.COL_HDR),
        ], 32));
        row++;

        const kobenAllocated = allocateByRatio(kobenFunds, kobenAllocatedQty);
        kobenAllocated.forEach((fund, idx) => {
            rows.push(rowXml(row, [
                cellXml(2, row, idx + 1, S.DATA_CENTER),
                cellXml(3, row, '집합투자', S.DATA_CENTER),
                cellXml(4, row, shortFundName(fund.name), S.DATA_TEXT),
                cellXml(5, row, fund.code ? parseInt(fund.code) || fund.code : '', S.DATA_CENTER),
                cellXml(6, row, fund.totalAssets, S.DATA_NUM),
                cellXml(7, row, fund.ratio, S.DATA_PCT),
                cellXml(8, row, fund.quantity, S.DATA_NUM),
                cellXml(9, row, unitPrice, S.DATA_NUM),
                cellXml(10, row, fund.amount, S.DATA_NUM),
                cellXml(11, row, fund.fee, S.DATA_NUM),
                cellXml(12, row, fund.totalWithFee, S.DATA_NUM),
                cellXml(13, row, fund.quantity > 0 ? lockUpStr : '', S.DATA_CENTER),
                cellXml(14, row, bank, S.DATA_TEXT),
            ]));
            row++;
        });

        // "코스닥벤처분 배정수량"
        const kobenTotalAssets = kobenAllocated.reduce((s, f) => s + f.totalAssets, 0);
        const kobenTotalAmount = kobenAllocated.reduce((s, f) => s + f.amount, 0);
        rows.push(rowXml(row, [
            cellXml(2, row, '코스닥벤처분 배정수량', S.SUMMARY_LABEL),
            cellXml(5, row, kobenAllocatedQty, S.SUMMARY_NUM),
            cellXml(7, row, kobenTotalAssets > 0 ? kobenTotalAmount / kobenTotalAssets : 0, S.SUMMARY_PCT),
        ]));
        row++;

        // Grand total
        row++;
        const grandAmount = ipoTotalAmount + kobenAllocated.reduce((s, f) => s + f.amount, 0);
        const grandFee = ipoAllocated.reduce((s, f) => s + f.fee, 0) + kobenAllocated.reduce((s, f) => s + f.fee, 0);
        const grandTotal = ipoAllocated.reduce((s, f) => s + f.totalWithFee, 0) + kobenAllocated.reduce((s, f) => s + f.totalWithFee, 0);
        rows.push(rowXml(row, [
            cellXml(2, row, '총 배정수량', S.SUMMARY_LABEL),
            cellXml(8, row, (ipoAllocatedQty || 0) + kobenAllocatedQty, S.SUMMARY_NUM),
            cellXml(9, row, unitPrice, S.SUMMARY_NUM),
            cellXml(10, row, grandAmount, S.SUMMARY_NUM),
            cellXml(11, row, grandFee, S.SUMMARY_NUM),
            cellXml(12, row, grandTotal, S.SUMMARY_NUM),
        ]));
    }

    // ── Column widths ──
    const colsXml = `<x:cols>
<x:col min="1" max="1" width="2" customWidth="1"/>
<x:col min="2" max="2" width="6" customWidth="1"/>
<x:col min="3" max="3" width="10" customWidth="1"/>
<x:col min="4" max="4" width="26" customWidth="1"/>
<x:col min="5" max="5" width="10" customWidth="1"/>
<x:col min="6" max="6" width="18" customWidth="1"/>
<x:col min="7" max="7" width="10" customWidth="1"/>
<x:col min="8" max="8" width="10" customWidth="1"/>
<x:col min="9" max="9" width="10" customWidth="1"/>
<x:col min="10" max="10" width="16" customWidth="1"/>
<x:col min="11" max="11" width="12" customWidth="1"/>
<x:col min="12" max="12" width="16" customWidth="1"/>
<x:col min="13" max="13" width="10" customWidth="1"/>
<x:col min="14" max="14" width="22" customWidth="1"/>
</x:cols>`;

    const mergeXml = mergeCells.length > 0
        ? `<x:mergeCells count="${mergeCells.length}">${mergeCells.join('')}</x:mergeCells>`
        : '';

    const newSheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<x:sheetFormatPr defaultColWidth="9" defaultRowHeight="15"/>
${colsXml}
<x:sheetData>
${rows.join('\n')}
</x:sheetData>
${mergeXml}
</x:worksheet>`;

    // ── Add to zip ──
    const sheetFiles = Object.keys(zip.files).filter(f => f.match(/xl\/worksheets\/sheet\d+\.xml/));
    const maxNum = Math.max(...sheetFiles.map(f => parseInt(f.match(/sheet(\d+)/)[1])));
    const newSheetFile = `sheet${maxNum + 1}.xml`;
    const newSheetId = Math.max(...sheets.map(s => s.sheetId)) + 1;
    const maxRId = Math.max(...rels.map(r => parseInt(r.id.replace('rId', '')) || 0));
    const newRId = `rId${maxRId + 1}`;

    zip.file(`xl/worksheets/${newSheetFile}`, newSheetXml);

    // Update workbook.xml
    const lastSheet = sheets[sheets.length - 1];
    const nsPrefix = lastSheet.fullMatch.startsWith('<x:') ? 'x:' : '';
    const newSheetTag = `<${nsPrefix}sheet name="${sheetName}" sheetId="${newSheetId}" r:id="${newRId}"/>`;
    wbXml = wbXml.replace(lastSheet.fullMatch, newSheetTag + lastSheet.fullMatch);
    zip.file('xl/workbook.xml', wbXml);

    // Update rels
    relsXml = relsXml.replace('</Relationships>',
        `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${newSheetFile}"/>\n</Relationships>`);
    zip.file('xl/_rels/workbook.xml.rels', relsXml);

    // Update [Content_Types].xml
    let ctXml = await zip.file('[Content_Types].xml').async('string');
    ctXml = ctXml.replace('</Types>',
        `<Override PartName="/xl/worksheets/${newSheetFile}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n</Types>`);
    zip.file('[Content_Types].xml', ctXml);

    // Write output
    const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    fs.writeFileSync(targetFile, output);

    // Summary
    console.log(`✓ 시트 "${sheetName}" 생성 완료 (스타일 포함)`);
    console.log(`  파일: ${targetFile}`);
    console.log(`  종목: ${stockName}`);
    console.log(`  주관사: ${leadManager}`);
    console.log(`  확정공모가: ${unitPrice.toLocaleString()}원`);
    if (ipoAllocatedQty) {
        console.log(`  IPO 배정수량: ${ipoAllocatedQty.toLocaleString()}주`);
        ipoAllocated.forEach(f => {
            console.log(`    - ${shortFundName(f.name)}: ${f.quantity.toLocaleString()}주, ${f.amount.toLocaleString()}원, 수수료 ${f.fee.toLocaleString()}원`);
        });
    }
    if (kobenAllocatedQty) {
        console.log(`  코벤 배정수량: ${kobenAllocatedQty.toLocaleString()}주`);
    }
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
