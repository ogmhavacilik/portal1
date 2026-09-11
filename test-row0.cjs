const XLSX = require('xlsx');

const workbook = XLSX.readFile("test.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
console.log(rawRows[0]);
