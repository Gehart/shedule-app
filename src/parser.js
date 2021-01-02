"use strict";
const xlsx_node = require('node-xlsx'), xlsx = require('xlsx'), fs = require('fs');
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
const baseColumns = {
    nOfSheet: 2,
    dayOfWeek: 0,
    nOfLesson: 1,
    timeOfLesson: 2,
    evenOdd: 3,
    group: {
        "start": 12,
        "end": 13
    },
    subgroup: 13,
    typeOfLesson: 14,
    classroom: 15,
    startRowOfSheet: 8,
    endRowOfSheet: 90,
};
const sheetName = workbook.SheetNames[baseColumns.nOfSheet];
const workingSheet = workbook.Sheets[sheetName];
const dayNameOfWeek = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
main();
function main() {
    const dayRanges = findDaysRanges();
    const parsedDays = dayRanges.map((el, i) => parseDay(el, dayNameOfWeek[i]));
    const combined = joinParcedDays(parsedDays);
    writeFile('out/result.json', combined);
}
function joinParcedDays(parsedDays) {
    return parsedDays.reduce((combined, current) => {
        Object.assign(combined.odd, current.odd);
        Object.assign(combined.even, current.even);
        return combined;
    });
}
function writeFile(outputFile, object) {
    const wbSheet = JSON.stringify(object);
    fs.writeFile(outputFile, wbSheet, 'utf8', (err) => {
        if (err) {
            throw err;
        }
    });
}
function findDaysRanges(dayNameColumn = 0) {
    const docMerges = workingSheet['!merges'];
    const dayMerges = docMerges.filter(el => el.s.c === dayNameColumn && el.e.c === dayNameColumn)
        .map(el => {
        return {
            start: el.s.r,
            end: el.e.r
        };
    })
        .reverse();
    return dayMerges;
}
// TODO: узнать про нормальный способ конструктора (с указанием типа для строгого typescript)
function Shedule() {
    this.odd = {};
    this.even = {};
}
function parseDay(rowRange, dayName) {
    const startRowOfDay = rowRange.start;
    const endRowOfDay = rowRange.end;
    let day = new Shedule();
    for (let i = 0; i < endRowOfDay - startRowOfDay + 1; i++) {
        const currentRow = i + startRowOfDay;
        const cellValue = getCellValue({ c: baseColumns.subgroup, r: currentRow });
        // если ячейка пустая, пропускаем
        if (typeof cellValue === undefined || !cellValue)
            continue;
        let lesson = {};
        lesson.name = cellValue;
        lesson.type = getCellValue({ c: baseColumns.typeOfLesson, r: currentRow });
        lesson.classroom = getCellValue({ c: baseColumns.classroom, r: currentRow });
        // проверяем на общие пары на потоке.
        if (lesson.name === lesson.type) {
            lesson = getCommonLessonInfo({ c: baseColumns.subgroup, r: currentRow });
        }
        day = addLessonToDay(day, lesson, dayName, i);
    }
    return day;
}
function getCommonLessonInfo(address) {
    let lesson = {};
    const docMerges = workingSheet['!merges'];
    let cellValue = workingSheet[numberToCharAddress(address.c) + '' + (address.r + 1)];
    let lessonRange;
    // найдем диапазон
    for (let merge of docMerges) {
        // если попадает в границы диапазона одного из !merges
        if ((address.c >= merge.s.c && address.c <= merge.e.c) &&
            (address.r >= merge.s.r && address.r <= merge.e.r)) {
            // cellValue = workingSheet[numberToCharAddress(merge.s.c) + '' + (merge.s.r + 1)];
            lessonRange = { start: merge.s.c, end: merge.e.c };
        }
    }
    lesson.name = getCellValue(address);
    lesson.type = getCellValue({ r: address.r, c: lessonRange.end + 1 });
    lesson.classroom = getCellValue({ r: address.r, c: lessonRange.end + 2 });
    return lesson;
}
// TODO: подумать над тем, чтобы сделать это красивее. update: да, надо
function addLessonToDay(day, lesson, dayName, index) {
    // создаем копию объекта
    let newDay = JSON.parse(JSON.stringify(day));
    const nOfLesson = Math.floor(index / 2);
    // присваивание свойств новому объекту. При необходимости свойства создаются
    if (index % 2 === 0) {
        if (!newDay.odd.hasOwnProperty(dayName)) {
            newDay.odd[dayName] = {};
        }
        if (!newDay.odd[dayName].hasOwnProperty(nOfLesson)) {
            newDay.odd[dayName][nOfLesson] = {};
        }
        newDay.odd[dayName][nOfLesson] = lesson;
    }
    else {
        if (!newDay.even.hasOwnProperty(dayName)) {
            newDay.even[dayName] = {};
        }
        if (!newDay.even[dayName].hasOwnProperty(nOfLesson)) {
            newDay.even[dayName][nOfLesson] = {};
        }
        newDay.even[dayName][nOfLesson] = lesson;
    }
    return newDay;
}
// получить значение в ячейке, даже если ячейка смежная
function getCellValue(cellAddress) {
    const docMerges = workingSheet['!merges'];
    let cellValue = workingSheet[numberToCharAddress(cellAddress.c) + '' + (cellAddress.r + 1)];
    // если в ячейке есть значение
    if (typeof cellValue != undefined && !!cellValue) {
        // возвращаем значение, избавляясь от пробелов
        return (cellValue.v + '').trim().split(/\s+/).join(' ');
    }
    else {
        // проверяем, является ли ячейка "частью" другой ячейки
        for (let merge of docMerges) {
            // если попадает в границы диапазона одного из !merges
            if ((cellAddress.c >= merge.s.c && cellAddress.c <= merge.e.c) &&
                (cellAddress.r >= merge.s.r && cellAddress.r <= merge.e.r)) {
                cellValue = workingSheet[numberToCharAddress(merge.s.c) + '' + (merge.s.r + 1)];
            }
        }
        if (!cellValue) {
            return '';
        }
        else {
            return (cellValue.v + '').trim().split(/\s+/).join(' ');
        }
    }
}
function numberToCharAddress(n) {
    const ACode = 'A'.charCodeAt(0);
    const ZCode = 'Z'.charCodeAt(0);
    const len = ZCode - ACode + 1;
    let charAddress = "";
    while (n >= 0) {
        charAddress = String.fromCharCode(n % len + ACode) + charAddress;
        n = Math.floor(n / len) - 1;
    }
    return charAddress;
}
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
