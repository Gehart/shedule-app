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
// TODO: находить все эти столбцы автоматически
// объект с адресами основных колонок - времени занятий, группы, аудитории
const baseInfoOfSheet = {
    nOfSheet: 2,
    dayOfWeek: 0,
    nOfLesson: 1,
    timeOfLesson: 2,
    evenOdd: 3,
    groupNameRow: 8,
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
const sheetName = workbook.SheetNames[baseInfoOfSheet.nOfSheet];
const workingSheet = workbook.Sheets[sheetName];
const mergesInSheet = workingSheet['!merges'];
const dayNameOfWeek = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
// regex для имени группы
// [а-яА-ЯёЁ]{2,7}\/[а-яА-ЯёЁ]*-[а-яА-ЯёЁ0-9-]*
main();
function main() {
    const dayRanges = findDaysRanges();
    // строка с названиями групп
    baseInfoOfSheet.groupNameRow = dayRanges[0].start - 1;
    baseInfoOfSheet.group = findGroupColumns();
    const parsedDays = dayRanges.map((el, i) => parseDay(el, dayNameOfWeek[i]));
    const combined = joinParcedDays(parsedDays);
    writeFile('out/result.json', combined);
}
function findGroupColumns() {
    const rangeInSheet = workingSheet["!ref"];
    // const lastCell = rangeInSheet.split(':')[1];
    const lastColumnInLetters = rangeInSheet
        .split(':')[1] // последняя ячейка
        .split('')
        .filter(el => /[A-Ra-r]/.test(el)) // отделяем буквы (адрес колонки) от чисел
        .join('');
    // TODO: найти последнюю колонку и научить переводить буквы в число, а потом найти, наконец, группу
    const lastColumn = charToNumberAddress(lastColumnInLetters);
    const rowValues = getValuesFromColumnRangeInRow(baseInfoOfSheet.groupNameRow, { start: 0, end: lastColumn });
    return { start: 0, end: 10 };
}
function getValuesFromColumnRangeInRow(row, range) {
    // TODO: получить массив мерджев в строке
    for (let merge of mergesInSheet) {
        if (row === merge.s.c === merge.e.c) {
        }
    }
    console.log('row', row);
    const rowMerges = mergesInSheet.filter(merge => row === merge.s.r && row === merge.e.r);
    console.log("rowMerges", rowMerges);
    // (cellAddress.r >= merge.s.r && cellAddress.r <= merge.e.r)) 
    // TODO: в каждом из них взять значение
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
        .reverse(); // reverse, так как merges считываются в обратном порядке.
    return dayMerges;
}
function findDayNameColumn(docMerges) {
    for (let nColumn = 0; nColumn < 5; nColumn++) {
        let hasMergedCells = docMerges.filter(el => el.s.c === nColumn && el.e.c === nColumn)
            .findIndex(el => {
            return (el.e.r - el.s.r) >= 7;
        });
        if (hasMergedCells !== -1) {
            return nColumn;
        }
    }
    return null;
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
        const cellValue = getCellValue({ c: baseInfoOfSheet.subgroup, r: currentRow });
        // если ячейка пустая, пропускаем
        if (typeof cellValue === undefined || !cellValue)
            continue;
        let lesson = {};
        lesson.name = cellValue;
        lesson.type = getCellValue({ c: baseInfoOfSheet.typeOfLesson, r: currentRow });
        lesson.classroom = getCellValue({ c: baseInfoOfSheet.classroom, r: currentRow });
        // проверяем на общие пары на потоке.
        if (lesson.name === lesson.type) {
            lesson = getCommonLessonInfo({ c: baseInfoOfSheet.subgroup, r: currentRow });
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
// преобразует адресс вида 'AA' в адрес 25
function charToNumberAddress(charAddress) {
    const alphabetLength = 26;
    const numberAdr = charAddress.toUpperCase().split('')
        .map(el => {
        return el.charCodeAt(0) - 'A'.charCodeAt(0);
    })
        .reduce((prev, curr, ind, array) => {
        return prev + ((curr + 1) * Math.pow(alphabetLength, array.length - 1 - ind));
    }, 0);
    return numberAdr - 1;
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
