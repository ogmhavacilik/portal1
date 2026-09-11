const rawRows = [
  ["HANGAR YER DESTEK TEÇHİZATLARI"],
  [
    "SIRA NO",
    "TEÇHİZAT ADI",
    "PARÇA NO (P/N) / MODEL",
    "SERİ NO (S/N)",
    "MİKTAR / KAPASİTE",
    "BULUNDUĞU YER",
    "DURUMU",
    "BAKIMA TABİ",
    "SON KONTROL / KALİBRASYON / BAKIM",
    "GELECEK KONTROL / KALİBRASYON / BAKIM",
    "SON KONTROLÜ YAPAN FİRMA",
    "AÇIKLAMA",
    "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
  ]
];

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
    if (s.includes("MİKTAR") || s.includes("MIKTAR") || s.includes("KAPASİTE")) score += 2;
    if (s.includes("BULUNDUĞU") || s.includes("LOKASYON") || s.includes("KONUM")) score += 2;
    if (s.includes("DURUM")) score += 2;
    if (s.includes("KALİBRASYON") || s.includes("BAKIM") || s.includes("KONTROL")) score += 2;
  }
  console.log(`Row ${r} score: ${score}`);
  if (score > maxScore && score >= 4) {
    maxScore = score;
    headerRowIdx = r;
  }
}

console.log("final headerRowIdx:", headerRowIdx);

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

console.log("sira", findColIdx(["SIRA", "NO."], ["SERİ", "SERI", "PARÇA", "PARCA", "P/N", "MODEL"]));
console.log("name", findColIdx(["TEÇHİZAT", "TECHİZAT", "MALZEME", "ARAÇ", "ARAC", "PLAKA", "ÜRÜN", "URUN", "EKİPMAN", "TANIM", "NAME"], ["FİRMA", "FIRMA", "KONTROL", "BAKIM", "YAPAN"]));
console.log("pn", findColIdx(["P/N", "PN", "PARÇA NO", "PARCA NO", "MODEL", "PART NUMBER", "PART NO"]));
console.log("sn", findColIdx(["S/N", "SN", "SERİ NO", "SERI NO", "SERİ", "SERI", "SERIAL"], ["SIRA"]));
console.log("miktar", findColIdx(["MİKTAR", "MIKTAR", "KAPASİTE", "KAPASITE", "ADET", "QTY", "QUANTITY"]));
console.log("loc", findColIdx(["BULUNDUĞU", "BULUNDUGU", "LOKASYON", "KONUM", "YER", "RAF", "DEPO", "LOCATION"]));
console.log("durum", findColIdx(["DURUM", "DURUMU", "STATUS", "FAALİYET"]));
console.log("kalibTabi", findColIdx(["KALİBRASYONA TABİ", "KALIBRASYONA TABI", "BAKIMA TABİ", "BAKIMA TABI", "TABİ Mİ", "TABI MI", "ÖMÜRLÜ", "OMURLU"]));

