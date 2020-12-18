var xlsx_node = require('node-xlsx'),
    xlsx = require('xlsx'),
    fs = require('fs');

const xlsFile = 'short_shedule.xlsx';
// const fileBuffer = fs.readFile('./'+ xlsFile, 'utf8', ()=>{console.log("somethisng");});
var workbook = xlsx.readFile(xlsFile, {
    // type: 'buffer',
    raw: false,
    cellHTML: false,
    cellStyles: false,
    sheets: 0
});
const firstSheet = workbook.SheetNames[0];
const wbSheet = JSON.stringify(workbook.Sheets[firstSheet]);
fs.writeFile('sheet.json', wbSheet, 'utf8', (err) => {
    if (err) throw err;
    console.log('file was saved!');
});

// console.log(workbook);

// Parse a buffer
// const worksheetsfrombuffer = xlsx_node.parse(fs.readfilesync(`${__dirname}/univ_shedule.xls`));

// Parse a file
const workSheetsFromFile = xlsx_node.parse(`${__dirname}/${xlsFile}`);
const jsonSheet = JSON.stringify(workSheetsFromFile[0]);
// console.log(workSheetsFromFile[0]);
// console.log(jsonSheet);
// fs.writeFile('sheet.json', jsonSheet, 'utf8', ()=>{});