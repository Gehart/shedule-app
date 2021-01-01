const xlsx_node = require('node-xlsx'),
      xlsx      = require('xlsx'),
      fs        = require('fs');

const resourcesDir = 'resources/';
const xlsFile = 'short_shedule.xlsx';
const workbook = xlsx.readFile(resourcesDir + xlsFile, {
    raw: false,
    cellText: false,
    celHTML: false,
    cellStyles: false,
    cellDates: true
});

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
    subgroup: 4,
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

const firstSheetName = workbook.SheetNames[0];
const workingSheet = workbook.Sheets[firstSheetName];

main();
function main() {
    const parsedDay = parseDay();
}

function writeFile(outputFile : string, object: any) {
    const wbSheet = JSON.stringify(object); 
    fs.writeFile(outputFile, wbSheet, 'utf8', (err) => {
        if (err) {
            throw err;
        }
    });
}

interface Shedule {
    odd?: {
        monday?: Lesson[],
        tuesday?: Lesson[],
        wednesday?: Lesson[],
        thursday?: Lesson[],
        friday?: Lesson[],
        sunday?: Lesson[]
    },
    even?: {
        monday?: Lesson[],
        tuesday?: Lesson[],
        wednesday?: Lesson[],
        thursday?: Lesson[],
        friday?: Lesson[],
        sunday?: Lesson[]
    }
}

function parseDay() {
    const startRowOfDay = 1;
    const endRowOfDay = 14;
    const day: Shedule = {
        odd: {
            friday: []
        },
        even: {
            friday: []
        }
    };

    for (let i = 0; i < endRowOfDay - startRowOfDay + 1; i++) {
        const currentRow = i + startRowOfDay;
        const cellValue = getCellValue({c: sheduleBase.subgroup, r: currentRow});
        // console.log(cellValue);
        if (!cellValue) continue; 

        const lesson : Lesson = {};
        lesson.name = cellValue.split(/\s+/).join(' ');
        const nOfLesson = Math.floor(i / 2);
        // console.log('nOfLesson', lesson.nOfLesson);

        lesson.type = getCellValue({c: sheduleBase.typeOfLesson, r: currentRow});
        lesson.classroom = getCellValue({c: sheduleBase.classroom, r: currentRow});

        if (i % 2 === 0) {
            day.odd.friday[nOfLesson] = lesson;
        }
        else {
            day.even.friday[nOfLesson] = lesson;
        }
    }


    writeFile('out/result.json', day);
}
// неделя
//     четная
//         день
//             пары
//                 номер пары
//                 пара 
//                 время пары ?
//                 тип занятий
//                 аудитория
//     нечетная
//         день
//             пары
//                 номер пары
//                 пара
//                 время пары ?
//                 тип занятий
//                 аудитория

interface Lesson {
    // nOfLesson?: number,
    name?: string,
    type?: string,
    classroom?: string
}

interface CellAddress {
    c: number,
    r: number
}

// получить значение в ячейке, даже если ячейка смежная
function getCellValue(cellAddress: CellAddress) : string {
    const docMerges = workingSheet['!merges'];
    let cellValue = workingSheet[numberToCharAddress(cellAddress.c) + '' + (cellAddress.r + 1)];

    if (!!cellValue) {
        // console.log('вариант по хорошему');
        // console.log("type of cell", typeof(cellValue.v));
        return (cellValue.v + '').trim();
    }
    else {
        // console.log('вариант по плохому');
        for (let merge of docMerges) {
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
            return (cellValue.v + '').trim();
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

