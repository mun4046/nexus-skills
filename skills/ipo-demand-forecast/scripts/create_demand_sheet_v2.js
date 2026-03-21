/**
 * 수요예측 점검 리스트 시트를 기존 엑셀 파일에 추가한다.
 * JSZip으로 직접 조작하여 서식을 포함한다.
 *
 * Usage:
 *   node create_demand_sheet_v2.js '<JSON_DATA>'
 */

const JSZip = require('jszip');
const fs = require('fs');

const data = JSON.parse(process.argv[2]);

const {
    targetFile, sheetName, stockName, writeDate, baseDate,
    leadManager, isVenture, shareType, unitPrice, lockUp,
    ipoFunds, kobenFunds, ipoTotalQuantity, kobenTotalQuantity
} = data;

// ── Helpers ──

function allocateByRatio(funds, totalQty) {
    const totalAssets = funds.reduce((s, f) => s + f.totalAssets, 0);
    if (totalAssets === 0) return funds.map(f => ({ ...f, ratio: 0, quantity: 0, amount: 0 }));

    let allocated = funds.map(f => {
        const ratio = f.totalAssets / totalAssets;
        const qty = Math.round(totalQty * ratio);
        return { ...f, ratio, quantity: qty, amount: qty * unitPrice };
    });

    let sumQty = allocated.reduce((s, f) => s + f.quantity, 0);
    let diff = totalQty - sumQty;
    while (diff !== 0) {
        const errors = allocated.map((f, i) => ({ i, err: (totalQty * f.ratio) - f.quantity }));
        if (diff > 0) { errors.sort((a, b) => b.err - a.err); allocated[errors[0].i].quantity += 1; diff--; }
        else { errors.sort((a, b) => a.err - b.err); allocated[errors[0].i].quantity -= 1; diff++; }
    }
    allocated.forEach(f => { f.amount = f.quantity * unitPrice; });
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
    return `<x:row r="${rowNum}"${htAttr} spans="1:11">${cells.join('')}</x:row>`;
}

// ── Style injection ──

function injectStyles(stylesXml) {
    const fontCount = parseInt((stylesXml.match(/<x:fonts count="(\d+)"/) || [, '0'])[1]);
    const fillCount = parseInt((stylesXml.match(/<x:fills count="(\d+)"/) || [, '0'])[1]);
    const borderCount = parseInt((stylesXml.match(/<x:borders count="(\d+)"/) || [, '0'])[1]);
    const xfCount = parseInt((stylesXml.match(/<x:cellXfs count="(\d+)"/) || [, '0'])[1]);

    const newFonts = [
        `<x:font><x:name val="맑은 고딕"/><x:sz val="14"/><x:b/><x:color rgb="ff000000"/></x:font>`,
        `<x:font><x:name val="맑은 고딕"/><x:sz val="11"/><x:b/><x:color rgb="ff000000"/></x:font>`,
        `<x:font><x:name val="맑은 고딕"/><x:sz val="11"/><x:b/><x:color rgb="ffFFFFFF"/></x:font>`,
        `<x:font><x:name val="맑은 고딕"/><x:sz val="10"/><x:color rgb="ff000000"/></x:font>`,
        `<x:font><x:name val="맑은 고딕"/><x:sz val="10"/><x:b/><x:color rgb="ff000000"/></x:font>`,
    ];

    const newFills = [
        `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="ffD6E4F0"/><x:bgColor rgb="ffD6E4F0"/></x:patternFill></x:fill>`,
        `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="ff4472C4"/><x:bgColor rgb="ff4472C4"/></x:patternFill></x:fill>`,
        `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="ffFFF2CC"/><x:bgColor rgb="ffFFF2CC"/></x:patternFill></x:fill>`,
        `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="ffE2EFDA"/><x:bgColor rgb="ffE2EFDA"/></x:patternFill></x:fill>`,
    ];

    const newBorders = [
        `<x:border><x:left style="thin"><x:color rgb="ff000000"/></x:left><x:right style="thin"><x:color rgb="ff000000"/></x:right><x:top style="thin"><x:color rgb="ff000000"/></x:top><x:bottom style="thin"><x:color rgb="ff000000"/></x:bottom></x:border>`,
        `<x:border><x:left style="thin"><x:color rgb="ff000000"/></x:left><x:right style="thin"><x:color rgb="ff000000"/></x:right><x:top style="thin"><x:color rgb="ff000000"/></x:top><x:bottom style="medium"><x:color rgb="ff000000"/></x:bottom></x:border>`,
    ];

    const f = { title: fontCount, header: fontCount + 1, headerWh: fontCount + 2, data: fontCount + 3, sum: fontCount + 4 };
    const fl = { lightBlue: fillCount, darkBlue: fillCount + 1, yellow: fillCount + 2, green: fillCount + 3 };
    const b = { thin: borderCount, medBot: borderCount + 1 };

    const newXfs = [
        // S.TITLE (xfCount+0)
        `<x:xf numFmtId="0" fontId="${f.title}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><x:alignment horizontal="left" vertical="center"/></x:xf>`,
        // S.INFO (xfCount+1)
        `<x:xf numFmtId="0" fontId="${f.data}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><x:alignment horizontal="left" vertical="center"/></x:xf>`,
        // S.SECTION_HDR (xfCount+2)
        `<x:xf numFmtId="0" fontId="${f.headerWh}" fillId="${fl.darkBlue}" borderId="${b.thin}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="center" vertical="center"/></x:xf>`,
        // S.COL_HDR (xfCount+3)
        `<x:xf numFmtId="0" fontId="${f.header}" fillId="${fl.lightBlue}" borderId="${b.thin}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="center" vertical="center" wrapText="1"/></x:xf>`,
        // S.DATA_TEXT (xfCount+4)
        `<x:xf numFmtId="0" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="left" vertical="center"/></x:xf>`,
        // S.DATA_NUM (xfCount+5)
        `<x:xf numFmtId="3" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.DATA_CENTER (xfCount+6)
        `<x:xf numFmtId="0" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="center" vertical="center"/></x:xf>`,
        // S.SUM_NUM (xfCount+7)
        `<x:xf numFmtId="3" fontId="${f.sum}" fillId="${fl.yellow}" borderId="${b.medBot}" xfId="0" applyFont="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.EMPTY_BORDER (xfCount+8)
        `<x:xf numFmtId="0" fontId="${f.data}" fillId="0" borderId="${b.thin}" xfId="0" applyFont="1" applyBorder="1"/>`,
        // S.PART_LABEL (xfCount+9): participation summary label, green bg
        `<x:xf numFmtId="0" fontId="${f.header}" fillId="${fl.green}" borderId="${b.thin}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="center" vertical="center"/></x:xf>`,
        // S.PART_VALUE (xfCount+10): participation summary value, green bg
        `<x:xf numFmtId="3" fontId="${f.header}" fillId="${fl.green}" borderId="${b.thin}" xfId="0" applyFont="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.PART_VALUE_TEXT (xfCount+11): participation summary text value
        `<x:xf numFmtId="0" fontId="${f.header}" fillId="${fl.green}" borderId="${b.thin}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><x:alignment horizontal="right" vertical="center"/></x:xf>`,
        // S.INFO_BOLD (xfCount+12): bold info text
        `<x:xf numFmtId="0" fontId="${f.header}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><x:alignment horizontal="left" vertical="center"/></x:xf>`,
    ];

    let xml = stylesXml;
    xml = xml.replace(/<x:fonts count="\d+"/, `<x:fonts count="${fontCount + newFonts.length}"`);
    xml = xml.replace(/<\/x:fonts>/, newFonts.join('') + '</x:fonts>');
    xml = xml.replace(/<x:fills count="\d+"/, `<x:fills count="${fillCount + newFills.length}"`);
    xml = xml.replace(/<\/x:fills>/, newFills.join('') + '</x:fills>');
    xml = xml.replace(/<x:borders count="\d+"/, `<x:borders count="${borderCount + newBorders.length}"`);
    xml = xml.replace(/<\/x:borders>/, newBorders.join('') + '</x:borders>');
    xml = xml.replace(/<x:cellXfs count="\d+"/, `<x:cellXfs count="${xfCount + newXfs.length}"`);
    xml = xml.replace(/<\/x:cellXfs>/, newXfs.join('') + '</x:cellXfs>');

    return {
        xml,
        S: {
            TITLE: xfCount,
            INFO: xfCount + 1,
            SECTION_HDR: xfCount + 2,
            COL_HDR: xfCount + 3,
            DATA_TEXT: xfCount + 4,
            DATA_NUM: xfCount + 5,
            DATA_CENTER: xfCount + 6,
            SUM_NUM: xfCount + 7,
            EMPTY_BORDER: xfCount + 8,
            PART_LABEL: xfCount + 9,
            PART_VALUE: xfCount + 10,
            PART_VALUE_TEXT: xfCount + 11,
            INFO_BOLD: xfCount + 12,
        }
    };
}

// ── Main ──

async function main() {
    const buf = fs.readFileSync(targetFile);
    const zip = await JSZip.loadAsync(buf);

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

    if (sheets.find(s => s.name === sheetName)) {
        console.error(`이미 동일한 시트명이 존재합니다: ${sheetName}`);
        process.exit(1);
    }

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

    // ── Build sheet ──
    const rows = [];
    const mergeCells = [];
    function addMerge(start, end) { mergeCells.push(`<x:mergeCell ref="${start}:${end}"/>`); }

    // Row 2: Title
    rows.push(rowXml(2, [cellXml(1, 2, `${stockName} 수요예측 점검 리스트`, S.TITLE)]));
    addMerge('A2', 'K2');

    // Row 4-9: Metadata
    rows.push(rowXml(4, [cellXml(1, 4, `작성일 : ${writeDate}`, S.INFO)]));
    rows.push(rowXml(5, [cellXml(1, 5, `작성기준일 : ${baseDate}`, S.INFO)]));
    rows.push(rowXml(6, [cellXml(1, 6, '작성자 : 투자운용본부 운용지원팀 최영', S.INFO)]));
    rows.push(rowXml(7, [cellXml(1, 7, `주관사 : ${leadManager}`, S.INFO)]));
    rows.push(rowXml(8, [cellXml(1, 8, `벤처기업 해당여부 : ${isVenture ? '해당' : '미해당'}`, S.INFO)]));
    rows.push(rowXml(9, [cellXml(1, 9, shareType || '신주 100%', S.INFO)]));

    // Row 11: Section header
    rows.push(rowXml(11, [cellXml(1, 11, '*기관투자자 참여수량', S.INFO_BOLD)]));

    // Rows 12-17: Participation summary
    rows.push(rowXml(12, [
        cellXml(1, 12, 'IPO', S.PART_LABEL), cellXml(2, 12, '', S.PART_LABEL), cellXml(3, 12, '', S.PART_LABEL),
        cellXml(4, 12, ipoTotalQuantity || 0, S.PART_VALUE), cellXml(5, 12, '', S.PART_LABEL),
    ]));
    addMerge('A12', 'C12'); addMerge('D12', 'E12');

    rows.push(rowXml(13, [
        cellXml(1, 13, '코벤', S.PART_LABEL), cellXml(2, 13, '', S.PART_LABEL), cellXml(3, 13, '', S.PART_LABEL),
        cellXml(4, 13, kobenTotalQuantity || 0, S.PART_VALUE), cellXml(5, 13, '', S.PART_LABEL),
    ]));
    addMerge('A13', 'C13'); addMerge('D13', 'E13');

    rows.push(rowXml(14, [
        cellXml(1, 14, '하이일드', S.PART_LABEL), cellXml(2, 14, '', S.PART_LABEL), cellXml(3, 14, '', S.PART_LABEL),
        cellXml(4, 14, '', S.PART_LABEL), cellXml(5, 14, '', S.PART_LABEL),
    ]));
    addMerge('A14', 'C14'); addMerge('D14', 'E14');

    rows.push(rowXml(15, [
        cellXml(1, 15, '하이일드', S.PART_LABEL), cellXml(2, 15, '', S.PART_LABEL), cellXml(3, 15, '', S.PART_LABEL),
        cellXml(4, 15, 'N/A', S.PART_VALUE_TEXT), cellXml(5, 15, '', S.PART_LABEL),
    ]));
    addMerge('A15', 'C15'); addMerge('D15', 'E15');

    rows.push(rowXml(16, [
        cellXml(1, 16, '일임', S.PART_LABEL), cellXml(2, 16, '', S.PART_LABEL), cellXml(3, 16, '', S.PART_LABEL),
        cellXml(4, 16, '', S.PART_LABEL), cellXml(5, 16, '', S.PART_LABEL),
    ]));
    addMerge('A16', 'C16'); addMerge('D16', 'E16');

    rows.push(rowXml(17, [
        cellXml(1, 17, '수요예측 참여 공모가액', S.PART_LABEL), cellXml(2, 17, '', S.PART_LABEL), cellXml(3, 17, '', S.PART_LABEL),
        cellXml(4, 17, unitPrice, S.PART_VALUE), cellXml(5, 17, '', S.PART_LABEL),
    ]));
    addMerge('A17', 'C17'); addMerge('D17', 'E17');

    // Row 20-21: Table headers
    rows.push(rowXml(20, [
        cellXml(1, 20, '펀드정보', S.SECTION_HDR), cellXml(2, 20, '', S.SECTION_HDR),
        cellXml(3, 20, '', S.SECTION_HDR), cellXml(4, 20, '', S.SECTION_HDR), cellXml(5, 20, '', S.SECTION_HDR),
        cellXml(6, 20, '수요예측 내역', S.SECTION_HDR), cellXml(7, 20, '', S.SECTION_HDR),
        cellXml(8, 20, '', S.SECTION_HDR), cellXml(9, 20, '', S.SECTION_HDR), cellXml(10, 20, '', S.SECTION_HDR),
        cellXml(11, 20, '비고', S.SECTION_HDR),
    ]));
    addMerge('A20', 'E20'); addMerge('F20', 'J20'); addMerge('K20', 'K21');

    rows.push(rowXml(21, [
        cellXml(1, 21, '연번', S.COL_HDR),
        cellXml(2, 21, '기관구분', S.COL_HDR),
        cellXml(3, 21, '펀드명', S.COL_HDR),
        cellXml(4, 21, '펀드코드', S.COL_HDR),
        cellXml(5, 21, '총자산 3개월 평균\r\n(FoF제외)', S.COL_HDR),
        cellXml(6, 21, '수량\r\n(주)', S.COL_HDR),
        cellXml(7, 21, '단가\r\n(원)', S.COL_HDR),
        cellXml(8, 21, '금액\r\n(원)', S.COL_HDR),
        cellXml(9, 21, '확약여부', S.COL_HDR),
        cellXml(10, 21, 'NAV한도내\r\n참여여부', S.COL_HDR),
        cellXml(11, 21, '', S.COL_HDR),
    ], 32));

    // ── IPO Fund Rows ──
    const ipoAllocated = allocateByRatio(ipoFunds || [], ipoTotalQuantity || 0);
    const lockUpStr = normalizedLockUp(lockUp);
    let row = 22;

    ipoAllocated.forEach((fund, idx) => {
        const navStatus = fund.totalAssets > 0 && fund.amount <= fund.totalAssets ? '한도내참여' : (fund.totalAssets === 0 ? '' : '한도이상');
        rows.push(rowXml(row, [
            cellXml(1, row, idx + 1, S.DATA_CENTER),
            cellXml(2, row, '집합투자', S.DATA_CENTER),
            cellXml(3, row, shortFundName(fund.name), S.DATA_TEXT),
            cellXml(4, row, fund.code ? parseInt(fund.code) || fund.code : '', S.DATA_CENTER),
            cellXml(5, row, fund.totalAssets, S.DATA_NUM),
            cellXml(6, row, fund.quantity, S.DATA_NUM),
            cellXml(7, row, unitPrice, S.DATA_NUM),
            cellXml(8, row, fund.amount, S.DATA_NUM),
            cellXml(9, row, lockUpStr, S.DATA_CENTER),
            cellXml(10, row, navStatus, S.DATA_CENTER),
            cellXml(11, row, '', S.EMPTY_BORDER),
        ]));
        row++;
    });

    // IPO sum row
    rows.push(rowXml(row, [
        cellXml(1, row, '', S.EMPTY_BORDER), cellXml(2, row, '', S.EMPTY_BORDER),
        cellXml(3, row, '', S.EMPTY_BORDER), cellXml(4, row, '', S.EMPTY_BORDER),
        cellXml(5, row, ipoAllocated.reduce((s, f) => s + f.totalAssets, 0), S.SUM_NUM),
        cellXml(6, row, ipoAllocated.reduce((s, f) => s + f.quantity, 0), S.SUM_NUM),
        cellXml(7, row, '', S.EMPTY_BORDER),
        cellXml(8, row, ipoAllocated.reduce((s, f) => s + f.amount, 0), S.SUM_NUM),
        cellXml(9, row, '', S.EMPTY_BORDER), cellXml(10, row, '', S.EMPTY_BORDER),
        cellXml(11, row, '', S.EMPTY_BORDER),
    ]));
    row++;

    // ── 코벤 Fund Rows ──
    if (kobenFunds && kobenFunds.length > 0 && kobenTotalQuantity > 0) {
        row++; // blank row
        const kobenAllocated = allocateByRatio(kobenFunds, kobenTotalQuantity);
        kobenAllocated.forEach((fund, idx) => {
            rows.push(rowXml(row, [
                cellXml(1, row, idx + 1, S.DATA_CENTER),
                cellXml(2, row, '코스닥벤처', S.DATA_CENTER),
                cellXml(3, row, shortFundName(fund.name), S.DATA_TEXT),
                cellXml(4, row, fund.code ? parseInt(fund.code) || fund.code : '', S.DATA_CENTER),
                cellXml(5, row, fund.totalAssets, S.DATA_NUM),
                cellXml(6, row, fund.quantity, S.DATA_NUM),
                cellXml(7, row, unitPrice, S.DATA_NUM),
                cellXml(8, row, fund.amount, S.DATA_NUM),
                cellXml(9, row, lockUpStr, S.DATA_CENTER),
                cellXml(10, row, fund.totalAssets > 0 && fund.amount <= fund.totalAssets ? '한도내참여' : '', S.DATA_CENTER),
                cellXml(11, row, '', S.EMPTY_BORDER),
            ]));
            row++;
        });

        // Koben sum row
        rows.push(rowXml(row, [
            cellXml(1, row, '', S.EMPTY_BORDER), cellXml(2, row, '', S.EMPTY_BORDER),
            cellXml(3, row, '', S.EMPTY_BORDER), cellXml(4, row, '', S.EMPTY_BORDER),
            cellXml(5, row, kobenAllocated.reduce((s, f) => s + f.totalAssets, 0), S.SUM_NUM),
            cellXml(6, row, kobenAllocated.reduce((s, f) => s + f.quantity, 0), S.SUM_NUM),
            cellXml(7, row, '', S.EMPTY_BORDER),
            cellXml(8, row, kobenAllocated.reduce((s, f) => s + f.amount, 0), S.SUM_NUM),
            cellXml(9, row, '', S.EMPTY_BORDER), cellXml(10, row, '', S.EMPTY_BORDER),
            cellXml(11, row, '', S.EMPTY_BORDER),
        ]));
    }

    // ── Column widths ──
    const colsXml = `<x:cols>
<x:col min="1" max="1" width="6" customWidth="1"/>
<x:col min="2" max="2" width="10" customWidth="1"/>
<x:col min="3" max="3" width="28" customWidth="1"/>
<x:col min="4" max="4" width="10" customWidth="1"/>
<x:col min="5" max="5" width="18" customWidth="1"/>
<x:col min="6" max="6" width="12" customWidth="1"/>
<x:col min="7" max="7" width="10" customWidth="1"/>
<x:col min="8" max="8" width="16" customWidth="1"/>
<x:col min="9" max="9" width="10" customWidth="1"/>
<x:col min="10" max="10" width="14" customWidth="1"/>
<x:col min="11" max="11" width="10" customWidth="1"/>
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

    const lastSheet = sheets[sheets.length - 1];
    const nsPrefix = lastSheet.fullMatch.startsWith('<x:') ? 'x:' : '';
    const newSheetTag = `<${nsPrefix}sheet name="${sheetName}" sheetId="${newSheetId}" r:id="${newRId}"/>`;
    wbXml = wbXml.replace(lastSheet.fullMatch, newSheetTag + lastSheet.fullMatch);
    zip.file('xl/workbook.xml', wbXml);

    relsXml = relsXml.replace('</Relationships>',
        `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${newSheetFile}"/>\n</Relationships>`);
    zip.file('xl/_rels/workbook.xml.rels', relsXml);

    let ctXml = await zip.file('[Content_Types].xml').async('string');
    ctXml = ctXml.replace('</Types>',
        `<Override PartName="/xl/worksheets/${newSheetFile}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n</Types>`);
    zip.file('[Content_Types].xml', ctXml);

    const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    fs.writeFileSync(targetFile, output);

    console.log(`✓ 시트 "${sheetName}" 생성 완료`);
    console.log(`  파일: ${targetFile}`);
    console.log(`  종목: ${stockName}`);
    console.log(`  주관사: ${leadManager}`);
    console.log(`  참여가격: ${unitPrice.toLocaleString()}원`);
    if (ipoTotalQuantity) {
        console.log(`  IPO 참여수량: ${ipoTotalQuantity.toLocaleString()}주`);
        ipoAllocated.forEach(f => {
            console.log(`    - ${shortFundName(f.name)}: ${f.quantity.toLocaleString()}주 (${f.amount.toLocaleString()}원)`);
        });
    }
    if (kobenTotalQuantity) {
        console.log(`  코벤 참여수량: ${kobenTotalQuantity.toLocaleString()}주`);
    }
}

main().catch(e => { console.error('Error:', e.message); process.exit(1); });
