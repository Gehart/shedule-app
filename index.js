var xlsx_node = require('node-xlsx'),
    xlsx = require('xlsx'),
    fs = require('fs');

const xlsFile = 'short_shedule.xlsx';
var workbook = xlsx.readFile(xlsFile, {
    raw: false,
    cellText: false,
    cellHTML: false,
    cellStyles: false,
    cellDates: true
});

const firstSheetName = workbook.SheetNames[0];
const workingSheet = workbook.Sheets[firstSheetName];

// writeFile('sheet.json');
function writeFile(outputFile) {
    const wbSheet = JSON.stringify(workingSheet); 
    fs.writeFile(outputFile, wbSheet, 'utf8', (err) => {
        if (err) throw err;
        console.log('file was saved!');
    });
}

main();
function main() {
    const cellValue = getCellValue({c:1,r:2});
    console.log('cellValue=',cellValue);
}

function getCellValue(cellAddress) {
    const docMerges = workingSheet['!merges'];
    let cellValue = workingSheet[numberToCharAddress(cellAddress.c) + '' + (cellAddress.r + 1)];

    if (!!cellValue) {
        console.log('вариант по хорошему');
        return cellValue.v;
    }
    else {
        console.log('вариант по плохому');
        for (merge of docMerges) {
            // если попадает в границы диапазона одного из !merges
            if ((cellAddress.c >= merge.s.c && cellAddress.c <= merge.e.c) &&
                (cellAddress.r >= merge.s.r && cellAddress.r <= merge.e.r)) 
            {
                // console.log('merges=',merge);
                // console.log(numberToCharAddress(merge.s.c) + '' + (merge.s.r + 1));
                cellValue = workingSheet[numberToCharAddress(merge.s.c) + '' + (merge.s.r + 1)]; 
            }
        }
        if (!cellValue) {
            return '';
        }
        else {
            return cellValue
        }
    }
}

function numberToCharAddress(n) {
    var ACode = 'A'.charCodeAt(0);
    var ZCode = 'Z'.charCodeAt(0);
    var len = ZCode - ACode + 1;

    var charAddress = "";
    while (n >= 0) {
        charAddress = String.fromCharCode(n % len + ACode) + charAddress;
        n = Math.floor(n / len) - 1;
    }
    return charAddress;
}
// // Parse a file
// const workSheetsFromFile = xlsx_node.parse(`${__dirname}/${xlsFile}`);
// const jsonSheet = JSON.stringify(workSheetsFromFile[0]);

const sheduleBase = {
    dayOfWeek: 0,
    nOfLesson: 1,
    timeOfLesson: 2,
    evenOdd: 3,
    group: {
        "s": {
            "c": 4,
            "r": 0
        },
        "e": {
            "c": 5,
            "r": 0
        }
    },
    subgroup: 5,
    typeOfLesson: 6, 
    classroom: 7,
    startRowOfSheet: '0',
    endRowOfSheet: '14'
    // dayOfWeek: 'A',
    // nOfLesson: 'B',
    // timeOfLesson: 'C',
    // evenOdd: 'D',
    // group: {
    //     "s": {
    //         "c": 4,
    //         "r": 0
    //     },
    //     "e": {
    //         "c": 5,
    //         "r": 0
    //     }
    // },
    // subgroup: 'F',
    // typeOfLesson: 'G', 
    // classroom: 'H',
};


/*
найти столбец подгруппы
идти сверху вниз, проверяя на !merges
определить столбцы дня недели, столбца пар, времени, чет/нечет, тип занятий, аудитория

неделя
    четная
        день
            пары
                номер пары
                пара 
                время пары ?
                тип занятий
                аудитория

    нечетная
        день
            пары
                номер пары
                пара
                время пары ?
                тип занятий
                аудитория
*/

