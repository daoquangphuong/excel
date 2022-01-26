const fs = require('fs');
const path = require('path');
const Excel = require('exceljs');

const ROOT = path.resolve(__dirname);
const INPUT_PATH = path.resolve(ROOT, 'input');
const OUTPUT_PATH = path.resolve(ROOT, 'output');

if(!fs.existsSync(INPUT_PATH)){
    fs.mkdirSync(INPUT_PATH);
}

if(!fs.existsSync(OUTPUT_PATH)){
    fs.mkdirSync(OUTPUT_PATH);
}

const processFile = async (filename) => {
    const workbook = new Excel.Workbook();
    console.log('reading', filename)
    await workbook.xlsx.readFile(path.resolve(INPUT_PATH, filename));
    const kgMinData = [];
    workbook.eachSheet(function (worksheet, sheetId) {
        if (worksheet.state !== 'visible') {
            return;
        }
        if (!worksheet.name.trim().match(/^\d+-\d+$/)) {
            return;
        }
        // if (worksheet.name !== '03-05') {
        //     return;
        // }
        console.log('processing', worksheet.name);
        const block1 = {
            is: false,
            start: 0,
            end: 0,
        };
        worksheet.eachRow((row, rowNumber) => {
            const fullRowText = JSON.stringify(row.values);
            if (fullRowText.toLowerCase().indexOf('total chemical input') > -1) {
                block1.is = true;
            }
            if (block1.is) {
                if (!block1.start) {
                    if (row.values[1] > 0) {
                        block1.start = rowNumber;
                    }
                } else if (!block1.end) {
                    if (!row.values[1]) {
                        block1.end = rowNumber;
                    }
                }
            }
            // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
        })
        kgMinData.push({
            name: worksheet.name,
            start: block1.start,
            end: block1.end,
        });
    });

    const sheet = workbook.addWorksheet('KG.min');
    sheet.getCell('A1').value = 'DATE'
    sheet.getCell('B1').value = 'FOAM GRADE'
    sheet.getCell('C1').value = 'HARDEST'
    sheet.getCell('D1').value = 'FOAM TYPE'
    sheet.getCell('E1').value = 'KG'
    sheet.getCell('F1').value = 'MIN'
    sheet.getCell('G1').value = 'KG*MIN'

    let l = 1;
    kgMinData.forEach((s) => {
        const sr = `'${s.name}'`;
        for (let i = s.start; i < s.end; i++) {
            l += 1;
            if (i === s.start) {
                sheet.getCell(`A${l}`).value = s.name;
            }
            sheet.getCell(`B${l}`).value = {formula: `${sr}!B${i}`};
            sheet.getCell(`C${l}`).value = {formula: `1*RIGHT(SUMPRODUCT(MID(0&B${l}, LARGE(INDEX(ISNUMBER(--MID(B${l}, ROW(INDIRECT("1:"&LEN(B${l}))), 1)) * ROW(INDIRECT("1:"&LEN(B${l}))), 0), ROW(INDIRECT("1:"&LEN(B${l}))))+1, 1) * 10^ROW(INDIRECT("1:"&LEN(B${l})))/10), 2)`};
            if (i === s.start) {
                const textJoin = [];
                for (let j = l; j < l + s.end - s.start; j++) {
                    textJoin.push(`B${j}`)
                }
                sheet.getCell(`D${l}`).value = {formula: textJoin.join(`&"/"&`)};
            }
            sheet.getCell(`E${l}`).value = {formula: `1*RIGHT(LEFT(FORMULATEXT(${sr}!E${i}),FIND("*",FORMULATEXT(${sr}!E${i}))-1),LEN(LEFT(FORMULATEXT(${sr}!E${i}),FIND("*",FORMULATEXT(${sr}!E${i}))-1))-1)`};
            sheet.getCell(`F${l}`).value = {formula: `1*RIGHT(FORMULATEXT(${sr}!E${i}),LEN(FORMULATEXT(${sr}!E${i}))-FIND("*",FORMULATEXT(${sr}!E${i})))`};
            sheet.getCell(`G${l}`).value = {formula: `${sr}!E${i}`};
        }
    })

    console.log('saving', filename)
    await workbook.xlsx.writeFile(path.resolve(OUTPUT_PATH, filename));
}

const main = async () => {
    const files = fs.readdirSync(INPUT_PATH);
    let seq = Promise.resolve();
    files.forEach((file) => {
        if (!file.endsWith('.xlsx')) {
            return;
        }
        // if(file !== 'FOAMING LOSS MILLENNIUM -1.2021 0k.xlsx'){
        //     return;
        // }
        seq = seq.then(async () => {
            await processFile(file)
        })
    })
    await seq;
    console.log('---------')
    console.log('OK - DONE')
    console.log('---------')
}

main().catch(console.error);