const headers = [
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
];

const findColIdx = (keywords, excludeKeywords = []) => {
  return headers.findIndex(h => {
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
console.log("sonBakim", findColIdx(["SON KONTROL", "SON BAKIM", "SON KALİBRASYON", "SON TEST", "SON MUAYENE", "SON KM", "YAPILAN KONTROL"]));
console.log("gelecekBakim", findColIdx(["GELECEK KONTROL", "GELECEK BAKIM", "GELECEK KALİBRASYON", "BİR SONRAKİ", "SONRAKİ BAKIM", "ÖMÜR BİTİŞ", "OMUR BITIS", "SON KULLANMA", "EXPIRY"]));
console.log("firma", findColIdx(["KONTROLÜ YAPAN", "KONTROLU YAPAN", "YAPAN FİRMA", "YAPAN FIRMA", "FİRMA", "FIRMA", "TEDARİK", "TEDARIK", "SERVİS", "VENDOR", "SUPPLIER"]));
