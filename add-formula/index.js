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
        const block2 = {
            is: false,
            start: 0,
            end: 0,
        };
        const block3 = {
            is: false,
            start: 0,
            end: 0,
        };
        worksheet.eachRow((row, rowNumber) => {
            const fullRowText = JSON.stringify(row.values);
            if (fullRowText.toLowerCase().indexOf('good blocks produced weight') > -1) {
                block2.is = true;
            }
            if (block2.is) {
                if (!block2.start) {
                    if (row.values[1] == 1) {
                        block2.start = rowNumber;
                    }
                } else if (!block2.end) {
                    if (!row.values[1]) {
                        block2.end = rowNumber;
                    }
                }
            }
            if (fullRowText.toLowerCase().indexOf('rejected blocks weight') > -1) {
                block3.is = true;
            }
            if (block3.is) {
                if (!block3.start) {
                    if (row.values[1] == 1) {
                        block3.start = rowNumber;
                    }
                } else if (!block3.end) {
                    if (!row.values[1]) {
                        block3.end = rowNumber;
                    }
                }
            }
            // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
        })
        const isNew = worksheet.getCell(`Q${block2.start-1}`).value === `ACTUAL WT`;
        const c = (v) => {
            if(!isNew){
                return v;
            }
            return String.fromCharCode(v.charCodeAt(0) + 1)
        }
        for (let i = block2.start; i < block2.end; i++) {
            //
            worksheet.getCell(`F${i}`).value = {formula: `VALUE(IF(RIGHT(B${i}, 1)="%", RIGHT(B${i}, 3), 0))`}
            worksheet.getCell(`G${i}`).value = {formula: `+${c('Q')}${i}*C${i}/100*D${i}/100*E${i}/100`}
            worksheet.getCell(`H${i}`).value = {formula: `IF(F${i}=0, G${i}, F${i}*G${i})`}
        }
        for (let i = block3.start; i < block3.end; i++) {
            worksheet.getCell(`F${i}`).value = {formula: `1-VALUE(IF(RIGHT(B${i}, 1)="%", RIGHT(B${i}, 3), 0))`}
            worksheet.getCell(`G${i}`).value = {formula: `+Q${i}*C${i}/100*D${i}/100*E${i}/100`}
            worksheet.getCell(`H${i}`).value = {formula: `IF(F${i}=0, G${i}, F${i}*G${i})`}
        }
        block2.end -= 1;
        const start = block2.start;
        const end = block2.end;
        worksheet.getCell(`${c('R')}${end + 2}`).value = {formula: `SUMIFS($${c('S')}$${start}:$${c('S')}$${end},$${c('R')}$${start}:$${c('R')}$${end},">=3")`}
        worksheet.getCell(`${c('S')}${end + 2}`).value = {formula: `SUM(${c('S')}${start}:${c('S')}${end})`}
        worksheet.getCell(`${c('T')}${end + 2}`).value = {formula: `SUM(${c('T')}${start}:${c('T')}${end})`}
        worksheet.getCell(`${c('S')}${end + 4}`).value = `KG`;
        worksheet.getCell(`${c('T')}${end + 4}`).value = `MRT`;
        worksheet.getCell(`${c('R')}${end + 5}`).value = `LONG BUN`;
        worksheet.getCell(`${c('S')}${end + 5}`).value = {formula: `SUMIFS($J$${start}:$J$${end},$${c('R')}$${start}:$${c('R')}$${end},">=3")`}
        worksheet.getCell(`${c('R')}${end + 6}`).value = `BLOCK`;
        worksheet.getCell(`${c('S')}${end + 6}`).value = {formula: `J${end + 2}-${c('S')}${end + 5}`}
        worksheet.getCell(`${c('S')}${end + 7}`).value = {formula: `+I${end + 2}`}
        worksheet.getCell(`${c('T')}${end + 6}`).value = {formula: `+SUMIFS($${c('T')}$${start}:$${c('T')}$${end},$${c('R')}$${start}:$${c('R')}$${end},">=3")`}
    });
    console.log('saving', filename)
    await workbook.xlsx.writeFile(path.resolve(OUTPUT_PATH, filename));
}

const main = async () => {
    const files = fs.readdirSync(INPUT_PATH);
    let seq = Promise.resolve();
    files.forEach((file) => {
        if(!file.endsWith('.xlsx')){
            return;
        }
        seq = seq.then(async () => {
            await processFile(file)
        })
    })
    await seq;
}

main().catch(console.error);