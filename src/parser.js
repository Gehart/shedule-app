"use strict";
const xlsx_node = require('node-xlsx'), xlsx = require('xlsx'), fs = require('fs');
const BaseBookInfo = {
    workbook: {},
    sheetName: '',
    workingSheet: {},
    mergesInSheet: {}
};
// объект с адресами основных колонок - времени занятий, группы, аудитории
const BaseInfoOfSheet = {
    nOfSheet: 0,
    dayName: 0,
    groupNameRow: 0,
    group: {},
    subgroup: 0,
    typeOfLesson: 0,
    classroom: 0,
    startRowOfSheet: 0,
    endRowOfSheet: 0,
};
main();
function main() {
    const resourcesDir = 'resources/';
    const xlsFile = 'univ_shedule.xls';
    try {
        const data = readASheduleFile(resourcesDir + xlsFile);
        const parsedShedule = parse(data, 4, 'ИВТ/б-17-1-о', 0);
        writeFile('out/result.json', parsedShedule);
    }
    catch (e) {
        console.error(e);
    }
}
function parse(data, course, groupName, subgroup) {
    BaseBookInfo.workbook = data;
    const sheetNames = BaseBookInfo.workbook.SheetNames.filter(el => el.includes(course + 'к'));
    const dayNameOfWeek = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    let parsedDays;
    for (let sheet of sheetNames) {
        try {
            BaseBookInfo.sheetName = sheet;
            BaseBookInfo.workingSheet = BaseBookInfo.workbook.Sheets[BaseBookInfo.sheetName];
            BaseBookInfo.mergesInSheet = BaseBookInfo.workingSheet['!merges'];
            BaseInfoOfSheet.dayName = findDayNameColumn();
            const dayRanges = findDaysRanges(BaseInfoOfSheet.dayName);
            // найдем столбец группы
            setGroup(groupName, dayRanges);
            if (BaseInfoOfSheet.group === null) {
                continue;
            }
            BaseInfoOfSheet.subgroup = (subgroup === 0) ? BaseInfoOfSheet.group.start : BaseInfoOfSheet.group.end;
            BaseInfoOfSheet.typeOfLesson = BaseInfoOfSheet.group.end + 1;
            BaseInfoOfSheet.classroom = BaseInfoOfSheet.group.end + 2;
            parsedDays = dayRanges.map((el, i) => parseDay(el, dayNameOfWeek[i]));
            if (typeof parsedDays !== undefined)
                break;
        }
        catch (e) {
            continue;
        }
    }
    if (typeof parsedDays === 'undefined') {
        throw new Error("Не нашли группy " + groupName + " в файле");
    }
    return joinParcedDays(parsedDays);
}
function readASheduleFile(fileName) {
    return xlsx.readFile(fileName, {
        raw: false,
        cellText: false,
        celHTML: false,
        cellStyles: false,
        cellDates: true
    });
}
function setGroup(groupName, dayRanges) {
    // строка с названиями групп. Обычно распологается на строку выше, чем названия дней недели.
    // цикл для возможность того, что названия группы располагаются не там, где ожидалось
    for (let i = 0; i < 2; i++) {
        BaseInfoOfSheet.groupNameRow = dayRanges[0].start - i - 1;
        BaseInfoOfSheet.group = findGroupColumnRange(groupName);
        if (BaseInfoOfSheet.group !== null)
            break;
    }
}
function findGroupColumnRange(groupName) {
    const lastColumnInSheet = findLastColumnInSheet();
    const rowValues = getValuesFromColumnRangeInRow(BaseInfoOfSheet.groupNameRow, { start: 0, end: lastColumnInSheet });
    const group = rowValues.find(el => el.value === groupName);
    if (typeof group === "undefined") {
        return null;
    }
    const groupMerge = getMergeAroundCell({ r: BaseInfoOfSheet.groupNameRow, c: group.column });
    return {
        start: groupMerge.s.c,
        end: groupMerge.e.c
    };
}
function findLastColumnInSheet() {
    const rangeInSheet = BaseBookInfo.workingSheet["!ref"];
    const lastColumnInLetters = rangeInSheet
        .split(':')[1] // последняя ячейка
        .split('')
        .filter(el => /[A-Ra-r]/.test(el)) // отделяем буквы (адрес колонки) от чисел
        .join('');
    return charToNumberAddress(lastColumnInLetters);
}
function getValuesFromColumnRangeInRow(row, range) {
    let rowValues = [];
    for (let i = range.start; i < range.end; i++) {
        rowValues.push({ column: i, value: getStrictCellValue({ r: row, c: i }) });
    }
    return rowValues.filter(el => el.value);
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
function compareRowRange(a, b) {
    return a.start >= b.start;
}
function findDaysRanges(dayNameColumn) {
    const docMerges = BaseBookInfo.workingSheet['!merges'];
    const dayMerges = docMerges.filter(el => el.s.c === dayNameColumn && el.e.c === dayNameColumn)
        .filter(el => el.e.r - el.s.r >= 7)
        .map(el => {
        return {
            start: el.s.r,
            end: el.e.r
        };
    })
        .sort(compareRowRange); // диапазоны иногда приходят в неправильном порядке
    return dayMerges;
}
function findDayNameColumn() {
    for (let nColumn = 0; nColumn < 5; nColumn++) {
        let hasMergedCells = BaseBookInfo.mergesInSheet.filter(el => el.s.c === nColumn && el.e.c === nColumn)
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
        const cellValue = getCellValue({ c: BaseInfoOfSheet.subgroup, r: currentRow });
        // если ячейка пустая, пропускаем
        if (typeof cellValue === undefined || !cellValue)
            continue;
        let lesson = {};
        lesson.name = cellValue;
        if (!lesson.name.includes("ВОЕННАЯ КАФЕДРА") &&
            !lesson.name.includes("ВОЕННО-УЧЕБНЫЙ ЦЕНТР") &&
            !lesson.name.includes("ОБЩЕУНИВЕРСИТЕТСКИЙ ПУЛ")) {
            lesson.type = getCellValue({ c: BaseInfoOfSheet.typeOfLesson, r: currentRow });
            lesson.classroom = getCellValue({ c: BaseInfoOfSheet.classroom, r: currentRow });
        }
        else {
            // нет смысла полностью обрабатывать дни военной кафедры и общий пул.
            lesson.type = '';
            lesson.classroom = '';
        }
        // проверяем на общие пары на потоке.
        if (lesson.name === lesson.type) {
            lesson = getCommonLessonInfo({ c: BaseInfoOfSheet.subgroup, r: currentRow });
        }
        day = addLessonToDay(day, lesson, dayName, i);
    }
    return day;
}
function getCommonLessonInfo(address) {
    let lesson = {};
    const docMerges = BaseBookInfo.workingSheet['!merges'];
    let cellValue = BaseBookInfo.workingSheet[numberToCharAddress(address.c) + '' + (address.r + 1)];
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
    const docMerges = BaseBookInfo.workingSheet['!merges'];
    let cellValue = BaseBookInfo.workingSheet[numberToCharAddress(cellAddress.c) + '' + (cellAddress.r + 1)];
    // если в ячейке есть значение
    if (typeof cellValue != undefined && !!cellValue) {
        // возвращаем значение, избавляясь от пробелов
        return (cellValue.v + '').trim().split(/\s+/).join(' ');
    }
    else {
        // проверяем, является ли ячейка "частью" другой ячейки
        let merge = getMergeAroundCell(cellAddress);
        if (merge) {
            cellValue = BaseBookInfo.workingSheet[numberToCharAddress(merge.s.c) + '' + (merge.s.r + 1)];
        }
        return (!cellValue) ? '' : (cellValue.v + '').trim().split(/\s+/).join(' ');
    }
}
function getMergeAroundCell(cellAddress) {
    const docMerges = BaseBookInfo.mergesInSheet;
    for (let merge of docMerges) {
        // если попадает в границы диапазона одного из !merges
        if ((cellAddress.c >= merge.s.c && cellAddress.c <= merge.e.c) &&
            (cellAddress.r >= merge.s.r && cellAddress.r <= merge.e.r)) {
            // cellValue = workingSheet[numberToCharAddress(merge.s.c) + '' + (merge.s.r + 1)];
            return merge;
        }
    }
    return null;
}
// получить значение ячейки без учитывания смежных ячеек
function getStrictCellValue(cellAddress) {
    let cellValue = BaseBookInfo.workingSheet[numberToCharAddress(cellAddress.c) + '' + (cellAddress.r + 1)];
    if (typeof cellValue != undefined && !!cellValue) {
        // возвращаем значение, избавляясь от пробелов
        return (cellValue.v + '').trim().split(/\s+/).join(' ');
    }
    else {
        return cellValue;
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
