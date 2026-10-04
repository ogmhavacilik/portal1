const XLSX = require('xlsx');
const path = require('path');
const PERSONNEL_EXCEL_FILE = path.resolve(process.cwd(), 'data/personnel_list.xlsx');
const workbook = XLSX.readFile(PERSONNEL_EXCEL_FILE);
console.log('Sheet Names:', workbook.SheetNames);
const sheetName = workbook.SheetNames.includes("AT-802 PERSONEL VERİSİ") 
    ? "AT-802 PERSONEL VERİSİ" 
    : (workbook.SheetNames.includes("PERSONEL LİSTESİ") ? "PERSONEL LİSTESİ" : workbook.SheetNames[0]);
console.log('Sheet Name:', sheetName);
const sheet = workbook.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet);
console.log('Rows:', JSON.stringify(rows, null, 2));
