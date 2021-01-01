const xlsx_node = require('node-xlsx'),
      xlsx      = require('xlsx'),
      fs        = require('fs');

// const xlsFile = 'short_shedule.xlsx';
const resourcesDir = 'resources/';
const xlsFile = 'univ_shedule.xls';
const workbook = xlsx.readFile(resourcesDir + xlsFile, {
    raw: false,
    cellText: false,
    celHTML: false,
    cellStyles: false,
    cellDates: true
});

// объект с адресами основных колонок - времени занятий, группы, аудитории
const sheduleBaseColumns = {
    nOfSheet: 2,
    dayOfWeek: 0,
    nOfLesson: 1,
    timeOfLesson: 2,
    evenOdd: 3,
    group: <ColumnRange>{
        "start": 12,
        "end": 13 
    },
    subgroup: 13,
    typeOfLesson: 14, 
    classroom: 15,
    startRowOfSheet: 8,
    endRowOfSheet: 90,

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

const sheetName = workbook.SheetNames[sheduleBaseColumns.nOfSheet];
const workingSheet = workbook.Sheets[sheetName];

const dayNameOfWeek = ['monday','tuesday','wednesday','thursday','friday','saturday'];
main();
function main() {
    const dayRanges: RowRange[] = findDaysRanges();
    const parsedDays = dayRanges.map((el, i) => parseDay(el, dayNameOfWeek[i]));
    const combined = joinParcedDays(parsedDays);
    writeFile('out/result.json', combined);
}

// TODO: попробовать с иммутабельностью позже
function joinParcedDays(parsedDays: Shedule[]) {
    let combined: Shedule = {odd:{},even:{}};
    parsedDays.map( el => {
        combined.odd = Object.assign(combined.odd, el.odd);
        combined.even = Object.assign(combined.even, el.even);
    });
    return combined;
}

function writeFile(outputFile : string, object: any) {
    const wbSheet = JSON.stringify(object); 
    fs.writeFile(outputFile, wbSheet, 'utf8', (err: Error) => {
        if (err) {
            throw err;
        }
    });
}

function findDaysRanges(dayNameColumn: number = 0): RowRange[] {
    const docMerges = workingSheet['!merges'];
    const dayMerges = docMerges.filter(el => el.s.c === dayNameColumn && el.e.c === dayNameColumn)
        .map(el => { 
            return <RowRange>{
                start: el.s.r,
                end: el.e.r
            }
        })
        .reverse(); 

    return dayMerges;
}

interface Shedule {
    odd: {
        monday?,
        tuesday?,
        wednesday?,
        thursday?,
        friday?,
        sunday?
    },
    even: {
        monday?,
        tuesday?,
        wednesday?,
        thursday?,
        friday?,
        sunday?
    }
}

interface RowRange {
    start: number,
    end: number
}

interface ColumnRange {
    start: number,
    end: number
}

// TODO: сделать нормальный тип возврата
function parseDay(rowRange: RowRange, dayName: string): Shedule {
    const startRowOfDay = rowRange.start;
    const endRowOfDay = rowRange.end;
    const day: Shedule = {
        odd: { },
        even: { }
    };

    for (let i = 0; i < endRowOfDay - startRowOfDay + 1; i++) {
        const currentRow = i + startRowOfDay;
        // console.log('cur row = ', currentRow);
        
        const cellValue = getCellValue({c: sheduleBaseColumns.subgroup, r: currentRow});
        // console.log(cellValue);
        if (!cellValue) continue; 

        const lesson : Lesson = {};
        lesson.name = cellValue.split(/\s+/).join(' ');
        const nOfLesson = Math.floor(i / 2);
        // console.log('nOfLesson', lesson.nOfLesson);

        lesson.type = getCellValue({c: sheduleBaseColumns.typeOfLesson, r: currentRow});
        lesson.classroom = getCellValue({c: sheduleBaseColumns.classroom, r: currentRow});
        
        if (i % 2 === 0) {
            if (!day.odd.hasOwnProperty(dayName)) {
                day.odd[dayName] = {};
            }
            if (!day.odd[dayName].hasOwnProperty(nOfLesson)) {
                day.odd[dayName][nOfLesson] = {};
            }
            day.odd[dayName][nOfLesson] = lesson;
        }
        else {
            if (!day.even.hasOwnProperty(dayName)) {
                day.even[dayName] = {};
            }
            if (!day.even[dayName].hasOwnProperty(nOfLesson)) {
                day.even[dayName][nOfLesson] = {};
            }
            day.even[dayName][nOfLesson] = lesson;
        }
    }
    // writeFile('out/result.json', day);
    return day;
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
        return (cellValue.v + '').trim();
    }
    else {
        // проверяем, является ли ячейка "частью" другой ячейки
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

function numberToCharAddress(n: number) {
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

