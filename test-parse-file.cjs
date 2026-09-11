const XLSX = require('xlsx');

const workbook = XLSX.readFile("test.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

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

let headerRowIdx = 0;
let maxScore = -1;
for (let r = 0; r < rawRows.length; r++) {
  const rCells = rawRows[r] || [];
  let score = 0;
  for (const c of rCells) {
    const s = String(c || '').trim().toUpperCase();
    if (s.includes("SIRA") || s === "NO" || s === "NO.") score += 2;
    if (s.includes("TEÇHİZAT") || s.includes("TECHIZAT") || s.includes("MALZEME") || s.includes("ÜRÜN") || s.includes("PLAKA")) score += 3;
    if (s.includes("P/N") || s.includes("PARÇA") || s.includes("PARCA") || s.includes("MODEL")) score += 2;
    if (s.includes("S/N") || s.includes("SERİ") || s.includes("SERI")) score += 2;
  }
  if (score > maxScore && score >= 4) {
    maxScore = score;
    headerRowIdx = r;
  }
}

console.log("headerRowIdx:", headerRowIdx);

const rawHeaders = (rawRows[headerRowIdx] || []).map(h => String(h || '').trim().toUpperCase());
const finalHeaders = rawHeaders.map((h, hIdx) => h || `KOLON ${hIdx + 1}`);

const findColIdx = (keywords, excludeKeywords = []) => {
  return finalHeaders.findIndex(h => {
    const upper = h.toUpperCase().trim();
    const hasKey = keywords.some(k => upper.includes(k.toUpperCase()));
    const hasExclude = excludeKeywords.some(ex => upper.includes(ex.toUpperCase()));
    return hasKey && !hasExclude;
  });
};

const siraColIdx = findColIdx(["SIRA", "NO."], ["SERİ", "SERI", "PARÇA", "PARCA", "P/N", "MODEL"]);
const nameColIdx = findColIdx(
  ["TEÇHİZAT", "TECHİZAT", "MALZEME", "ARAÇ", "ARAC", "PLAKA", "ÜRÜN", "URUN", "EKİPMAN", "TANIM", "NAME"],
  ["FİRMA", "FIRMA", "KONTROL", "BAKIM", "YAPAN"]
);
const pnColIdx = findColIdx(["P/N", "PN", "PARÇA NO", "PARCA NO", "MODEL", "PART NUMBER", "PART NO"]);
const snColIdx = findColIdx(["S/N", "SN", "SERİ NO", "SERI NO", "SERİ", "SERI", "SERIAL"], ["SIRA"]);
const miktarColIdx = findColIdx(["MİKTAR", "MIKTAR", "KAPASİTE", "KAPASITE", "ADET", "QTY", "QUANTITY"]);
const locColIdx = findColIdx(["BULUNDUĞU", "BULUNDUGU", "LOKASYON", "KONUM", "YER", "RAF", "DEPO", "LOCATION"]);

console.log({siraColIdx, nameColIdx, pnColIdx, snColIdx, miktarColIdx, locColIdx});

for (let r = headerRowIdx + 1; r < headerRowIdx + 4; r++) {
  const rawRow = rawRows[r] || [];
  
  const getVal = (idx, fallback) => {
    if (idx !== -1 && rawRow[idx] !== undefined) return String(rawRow[idx]).trim();
    return fallback;
  };

  const rawSira = getVal(siraColIdx, "");
  const rawName = getVal(nameColIdx, rawRow[1] ? String(rawRow[1]).trim() : (rawRow[0] ? String(rawRow[0]).trim() : ""));
  const rawPn = getVal(pnColIdx, rawRow[2] ? String(rawRow[2]).trim() : "");
  const rawSn = getVal(snColIdx, rawRow[3] ? String(rawRow[3]).trim() : "-");
  const rawMiktar = getVal(miktarColIdx, rawRow[4] ? String(rawRow[4]).trim() : "1");
  const rawLoc = getVal(locColIdx, rawRow[5] ? String(rawRow[5]).trim() : "");

  console.log(`Row ${r} output -> Sira: ${rawSira}, Name: ${rawName}, PN: ${rawPn}, Miktar: ${rawMiktar}, Loc: ${rawLoc}`);
}
