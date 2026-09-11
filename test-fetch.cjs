const XLSX = require('xlsx');

async function test() {
  const fileId = "1_a8BrcQlerc9qLJUbITFl4qKafgZZOxU";
  const url = `https://script.google.com/macros/s/AKfycbxBVLlhvSrYDsIY5-Z8-RQ4f2-kTrHbLrZN3Bk7hB3AtQotugE9xqXRscIZd6ruinlqTg/exec?action=readExcelFromDrive&fileId=${fileId}`;
  
  const res = await fetch(url);
  const data = await res.json();
  const workbook = XLSX.read(data.base64, { type: 'base64' });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

  // unmergeAndFillWorksheet
  if (worksheet['!merges']) {
    worksheet['!merges'].forEach(range => {
      const startCellRef = XLSX.utils.encode_cell(range.s);
      const startCell = worksheet[startCellRef];
      if (!startCell) return;
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          if (R === range.s.r && C === range.s.c) continue;
          const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
          worksheet[cellRef] = { ...startCell };
        }
      }
    });
  }

  const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
  for (let i=0; i<5; i++) {
    console.log(`Row ${i}:`, rawRows[i]);
  }
}

test();
