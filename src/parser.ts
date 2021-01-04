const xlsx_node = require('node-xlsx'),
      xlsx      = require('xlsx'),
      fs        = require('fs');


const BaseBookInfo = {
    workbook: <any>{},
    sheetName: '',
    workingSheet: {},
    mergesInSheet: <Merge[]>{}
}

// объект с адресами основных колонок - времени занятий, группы, аудитории
const BaseInfoOfSheet = {
    nOfSheet: 2,
    dayName: 0,
    groupNameRow: 8,
    group: <ColumnRange>{
        "start": 12,
        "end": 13 
    },
    subgroup: 13,
    typeOfLesson: 14, 
    classroom: 15,
    startRowOfSheet: -1,
    endRowOfSheet: 90,
};


main();
function main() {
    const resourcesDir = 'resources/';
    const xlsFile = 'univ_shedule.xls';
    try {
        const parsedShedule = parse(resourcesDir + xlsFile, 2, 'ИС/б-19-2о', 0);
        writeFile('out/result.json', parsedShedule);
    }
    catch(e){
        console.error(e);
    }
}

function readASheduleFile(fileName: string) {
    return xlsx.readFile(fileName, {
        raw: false,
        cellText: false,
        celHTML: false,
        cellStyles: false,
        cellDates: true
    });
}

function parse(fileName: string, course: number, groupName: string, subgroup: 0 | 1) {
    BaseBookInfo.workbook = readASheduleFile(fileName);
    setBaseBookInfo(course);

    BaseInfoOfSheet.dayName = findDayNameColumn();
    const dayRanges: RowRange[] = findDaysRanges(BaseInfoOfSheet.dayName);
    // строка с названиями групп. Обычно распологается на строку выше, чем названия дней недели.
    BaseInfoOfSheet.groupNameRow = dayRanges[0].start - 1;
    BaseInfoOfSheet.group = findGroupColumnRange(groupName);
    BaseInfoOfSheet.subgroup = (subgroup === 0) ? BaseInfoOfSheet.group.start : BaseInfoOfSheet.group.end;
    BaseInfoOfSheet.typeOfLesson = BaseInfoOfSheet.group.end + 1; 
    BaseInfoOfSheet.classroom = BaseInfoOfSheet.group.end + 2;

    const dayNameOfWeek = ['monday','tuesday','wednesday','thursday','friday','saturday'];
    const parsedDays = dayRanges.map((el, i) => parseDay(el, dayNameOfWeek[i]));
    
    return joinParcedDays(parsedDays);
}

function setBaseBookInfo(course: number) {
    BaseBookInfo.sheetName = BaseBookInfo.workbook.SheetNames.find(el => el.includes(course + 'к'));
    BaseBookInfo.workingSheet = BaseBookInfo.workbook.Sheets[BaseBookInfo.sheetName];
    BaseBookInfo.mergesInSheet = BaseBookInfo.workingSheet['!merges'];
}

function findGroupColumnRange(groupName: string): ColumnRange {
    const lastColumnInSheet = findLastColumnInSheet();
    const rowValues = getValuesFromColumnRangeInRow(BaseInfoOfSheet.groupNameRow, {start: 0, end: lastColumnInSheet});
    
    const group = rowValues.find(el => el.value === groupName);
    
    if (typeof group === "undefined") {
        throw new Error("Не нашли группy "+groupName+" на листе \""+ BaseBookInfo.sheetName + "\"");
    }

    const groupMerge = getMergeAroundCell({r: BaseInfoOfSheet.groupNameRow, c: group.column});
    return {
        start: groupMerge.s.c, 
        end: groupMerge.e.c
    };
}

function findLastColumnInSheet() {
    const rangeInSheet = BaseBookInfo.workingSheet["!ref"];
    const lastColumnInLetters = rangeInSheet
        .split(':')[1]   // последняя ячейка
        .split('')
        .filter(el => /[A-Ra-r]/.test(el))   // отделяем буквы (адрес колонки) от чисел
        .join('');
    
    return charToNumberAddress(lastColumnInLetters);
}

function getValuesFromColumnRangeInRow(row: number, range: ColumnRange) {
    let rowValues = [];
    for(let i = range.start; i < range.end; i++) {
        rowValues.push({column: i, value: getStrictCellValue({r: row, c: i})});
    }

    return rowValues.filter(el => el.value);
}

function joinParcedDays(parsedDays: Shedule[]) {
    return parsedDays.reduce((combined, current) => {
        Object.assign(combined.odd, current.odd);
        Object.assign(combined.even, current.even);
        return combined;
    });
}

function writeFile(outputFile : string, object: any) {
    const wbSheet = JSON.stringify(object); 
    fs.writeFile(outputFile, wbSheet, 'utf8', (err: Error) => {
        if (err) {
            throw err;
        }
    });
}

function compareRowRange(a: RowRange, b: RowRange) {
    return a.start >= b.start;
}

function findDaysRanges(dayNameColumn: number): RowRange[] {
    const docMerges = BaseBookInfo.workingSheet['!merges'];
    const dayMerges = docMerges.filter(el => el.s.c === dayNameColumn && el.e.c === dayNameColumn)
        .map(el => { 
            return <RowRange>{
                start: el.s.r,
                end:   el.e.r
            }
        })
        .sort(compareRowRange); // диапазоны иногда приходят в неправильном порядке
        
    return dayMerges;
}

function findDayNameColumn(): number | null {
    for(let nColumn = 0; nColumn < 5; nColumn++) {
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

// TODO: узнать про нормальный способ конструктора (с указанием типа для строгого typescript)
function Shedule() {
    this.odd = {};
    this.even = {};
}

interface RowRange {
    start: number,
    end: number
}

interface ColumnRange {
    start: number,
    end: number
}

function parseDay(rowRange: RowRange, dayName: string): Shedule {
    const startRowOfDay = rowRange.start;
    const endRowOfDay = rowRange.end;
    let day: Shedule = new Shedule();

    for (let i = 0; i < endRowOfDay - startRowOfDay + 1; i++) {
        const currentRow = i + startRowOfDay;
        
        const cellValue = getCellValue({c: BaseInfoOfSheet.subgroup, r: currentRow});
        // если ячейка пустая, пропускаем
        if (typeof cellValue === undefined || !cellValue) continue; 

        let lesson: Lesson = {};
        lesson.name = cellValue;
        if (!lesson.name.includes("ВОЕННАЯ КАФЕДРА") && 
            !lesson.name.includes("ОБЩЕУНИВЕРСИТЕТСКИЙ ПУЛ")) 
        {
            lesson.type = getCellValue({ c: BaseInfoOfSheet.typeOfLesson, r: currentRow });
            lesson.classroom = getCellValue({ c: BaseInfoOfSheet.classroom, r: currentRow });
        }
        else {
            // нет смысла полностью обрабатывать дни военной кафедры и общий пул.
            lesson.type = '';
            lesson.classroom = '';
            day = addLessonToDay(day, lesson, dayName, i);
            // return day;
        }

        // проверяем на общие пары на потоке.
        if (lesson.name === lesson.type) {
            lesson = getCommonLessonInfo({c: BaseInfoOfSheet.subgroup, r: currentRow});
        }
        
        day = addLessonToDay(day, lesson, dayName, i);
    }
    return day;
}

function getCommonLessonInfo(address: CellAddress): Lesson {
    let lesson: Lesson = {};
    const docMerges = BaseBookInfo.workingSheet['!merges'];
    let cellValue = BaseBookInfo.workingSheet[numberToCharAddress(address.c) + '' + (address.r + 1)];
    let lessonRange: ColumnRange;
    // найдем диапазон
    for (let merge of docMerges) {
        // если попадает в границы диапазона одного из !merges
        if ((address.c >= merge.s.c && address.c <= merge.e.c) &&
            (address.r >= merge.s.r && address.r <= merge.e.r)) {
            // cellValue = workingSheet[numberToCharAddress(merge.s.c) + '' + (merge.s.r + 1)];
            lessonRange = {start: merge.s.c, end: merge.e.c};
        }
    }
    lesson.name = getCellValue(address);
    lesson.type = getCellValue({r: address.r, c: lessonRange.end + 1});
    lesson.classroom = getCellValue({r: address.r, c: lessonRange.end + 2});
    return lesson;
}

// TODO: подумать над тем, чтобы сделать это красивее. update: да, надо
function addLessonToDay(day: Shedule, lesson: Lesson, dayName: string, index: number): Shedule {
    // создаем копию объекта
    let newDay: Shedule = JSON.parse(JSON.stringify(day));
    
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
        if(merge) {
            cellValue = BaseBookInfo.workingSheet[numberToCharAddress(merge.s.c) + '' + (merge.s.r + 1)]; 
        }
        return (!cellValue) ? '' : (cellValue.v + '').trim().split(/\s+/).join(' ');
    }
}

interface Merge {
    s: CellAddress,
    e: CellAddress
}

function getMergeAroundCell(cellAddress: CellAddress): Merge | null {
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
function getStrictCellValue(cellAddress: CellAddress): string {
    let cellValue = BaseBookInfo.workingSheet[numberToCharAddress(cellAddress.c) + '' + (cellAddress.r + 1)];
    if (typeof cellValue != undefined && !!cellValue) {
        // возвращаем значение, избавляясь от пробелов
        return (cellValue.v + '').trim().split(/\s+/).join(' ');
    }
    else {
        return cellValue;
    }
}

function numberToCharAddress(n: number) {
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
function charToNumberAddress(charAddress: string): number {
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