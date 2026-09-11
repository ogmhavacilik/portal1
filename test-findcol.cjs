const finalHeaders = [ 'SIRA NO',
  'TEÇHİZAT ADI',
  'PARÇA NO (P/N) / MODEL',
  'SERİ NO (S/N)',
  'MİKTAR / KAPASİTE',
  'BULUNDUĞU YER',
  'DURUMU',
  'BAKIMA TABİ',
  'SON KONTROL / KALİBRASYON / BAKIM',
  'GELECEK KONTROL / KALİBRASYON / BAKIM',
  'SON KONTROLÜ YAPAN FİRMA',
  'AÇIKLAMA',
  '90 GÜN UYARISI MAİL GÖNDERİM TARİHİ' ];

const findColIdx = (keywords, excludeKeywords = []) => {
  return finalHeaders.findIndex(h => {
    const upper = h.toUpperCase().trim();
    const hasKey = keywords.some(k => upper.includes(k.toUpperCase()));
    const hasExclude = excludeKeywords.some(ex => upper.includes(ex.toUpperCase()));
    return hasKey && !hasExclude;
  });
};

const nameColIdx = findColIdx(
  ["TEÇHİZAT", "TECHİZAT", "MALZEME", "ARAÇ", "ARAC", "PLAKA", "ÜRÜN", "URUN", "EKİPMAN", "TANIM", "NAME"],
  ["FİRMA", "FIRMA", "KONTROL", "BAKIM", "YAPAN"]
);

console.log("nameColIdx:", nameColIdx);
