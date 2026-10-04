import * as XLSX from 'xlsx';
import {
  cleanAndFormatDateString,
  isHeaderLikeRow,
  detectHeaderRowIndex,
  unmergeAndFillWorksheet,
  groupMultiLocationRowsHelper
} from './driveExcelSync';

export interface ParseExcelOptions {
  systemType: 'yer_destek' | 'depo' | 'kara_araclari';
  subSection?: string;
  unitKey?: string;
}

/**
 * Birebir Ortak Excel Yükleme ve Çözümleme Fonksiyonu
 * Yer Destek, Özel Aletler, Depo Sarf, Depo Kimyasal ve Kara Araçları sistemlerinde
 * aynı yüksek performanslı algoritma, birleştirilmiş hücre çözme ve veri eşleme ile çalışır.
 */
export const parseUploadedExcelFile = async (
  file: File,
  options: ParseExcelOptions
): Promise<string[][]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error("Excel dosyasında çalışma sayfası bulunamadı.");
        }

        // İlk sayfayı al veya hedef birim / kategori adı geçen sayfayı seç
        let targetSheetName = workbook.SheetNames[0];
        if (options.subSection && workbook.SheetNames.length > 1) {
          const matchSub = workbook.SheetNames.find(s => 
            s.toLowerCase().includes(options.subSection!.toLowerCase())
          );
          if (matchSub) targetSheetName = matchSub;
        } else if (options.unitKey && workbook.SheetNames.length > 1) {
          const match = workbook.SheetNames.find(s => 
            s.toLowerCase().includes(options.unitKey!.toLowerCase())
          );
          if (match) targetSheetName = match;
        }

        const worksheet = workbook.Sheets[targetSheetName];
        if (!worksheet) {
          throw new Error("Çalışma sayfası okunamadı.");
        }

        // 1. Birleştirilmiş hücreleri unroll yap
        unmergeAndFillWorksheet(worksheet);

        // 2. Ham satırları oku
        const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: "", raw: false });
        if (!rawRows || rawRows.length === 0) {
          throw new Error("Excel dosyası boş.");
        }

        // 3. Akıllı Başlık Tespiti
        const headerRowIdx = detectHeaderRowIndex(rawRows);
        const rawHeaders = (rawRows[headerRowIdx] || []).map(h => String(h || '').trim().toUpperCase());
        const finalHeaders = rawHeaders.map((h, idx) => h || `KOLON ${idx + 1}`);

        const findColIdx = (keywords: string[], excludeKeywords: string[] = []): number => {
          return finalHeaders.findIndex(h => {
            const upper = h.toUpperCase().trim();
            const hasKey = keywords.some(k => upper.includes(k.toUpperCase()));
            const hasExclude = excludeKeywords.some(ex => upper.includes(ex.toUpperCase()));
            return hasKey && !hasExclude;
          });
        };

        const siraColIdx = findColIdx(["SIRA", "NO.", "S.N.", "S.NO", "S.NU"], ["SERİ", "SERI", "PARÇA", "PARCA", "P/N", "MODEL"]);
        const nameColIdx = findColIdx(
          ["TEÇHİZAT", "TECHİZAT", "MALZEME", "ARAÇ", "ARAC", "PLAKA", "ÜRÜN", "URUN", "EKİPMAN", "TANIM", "NAME", "ALET", "ITEM", "DESCRIPTION"],
          ["FİRMA", "FIRMA", "KONTROL", "BAKIM", "YAPAN", "SIRA"]
        );
        const pnColIdx = findColIdx(["P/N", "PN", "PARÇA NO", "PARCA NO", "MODEL", "PART NUMBER", "PART NO"]);
        const snColIdx = findColIdx(["S/N", "SN", "SERİ NO", "SERI NO", "SERİ", "SERI", "SERIAL", "LOT", "NOTLAR"], ["SIRA"]);
        const miktarColIdx = findColIdx(["MİKTAR", "MIKTAR", "KAPASİTE", "KAPASITE", "ADET", "QTY", "QUANTITY", "GELEN", "STOK"]);
        const locColIdx = findColIdx(["BULUNDUĞU", "BULUNDUGU", "LOKASYON", "KONUM", "YER", "RAF", "DEPO", "LOCATION"]);
        const durumColIdx = findColIdx(["DURUM", "DURUMU", "STATUS", "FAALİYET"]);
        const kalibTabiColIdx = findColIdx(["KALİBRASYONA TABİ", "KALIBRASYONA TABI", "BAKIMA TABİ", "BAKIMA TABI", "TABİ Mİ", "TABI MI", "TABİ", "TABI", "ÖMÜRLÜ", "OMURLU", "RAF ÖMRÜ", "RAF OMRU"]);
        const sonBakimColIdx = findColIdx(["SON KONTROL", "SON BAKIM", "SON KALİBRASYON", "SON TEST", "SON MUAYENE", "SON KM", "YAPILAN KONTROL", "SON TARİH", "GİRİŞ", "GIRIS", "İMAL", "IMAL"]);
        const gelecekBakimColIdx = findColIdx(["GELECEK KONTROL", "GELECEK BAKIM", "GELECEK KALİBRASYON", "BİR SONRAKİ", "SONRAKİ BAKIM", "SONRAKİ KONTROL", "ÖMÜR BİTİŞ", "OMUR BITIS", "SON KULLANMA", "SKT", "EXPIRY"]);
        const firmaColIdx = findColIdx(["KONTROLÜ YAPAN", "KONTROLU YAPAN", "YAPAN FİRMA", "YAPAN FIRMA", "FİRMA", "FIRMA", "TEDARİK", "TEDARIK", "SERVİS", "VENDOR", "SUPPLIER"]);
        const aciklamaColIdx = findColIdx(["AÇIKLAMA", "ACIKLAMA", "AÇIKLAMALAR", "ACIKLAMALAR", "NOT", "NOTLAR", "DESCRIPTION", "REMARKS", "DETAY"]);
        const mailColIdx = findColIdx(["MAİL GÖNDERİM", "MAIL GONDERIM", "90 GÜN", "90 GUN", "E-POSTA", "MAIL", "MAİL"]);

        const isDepo = options.systemType === 'depo' || options.subSection === 'depo_sarf' || options.subSection === 'depo_kimyasal';
        const isKara = options.systemType === 'kara_araclari';
        const finalSection = options.subSection || (isDepo ? 'depo_sarf' : isKara ? 'kara_araclari' : 'yer_destek');

        const rawParsedRows: string[][] = [];

        for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
          const rawRow = rawRows[r] || [];
          const isRowEmpty = rawRow.every(cell => String(cell || '').trim() === '');
          if (isRowEmpty) continue;
          if (isHeaderLikeRow(rawRow)) continue;

          const getVal = (idx: number, fallback = "") => {
            if (idx >= 0 && rawRow[idx] !== undefined && rawRow[idx] !== null) {
              const valStr = String(rawRow[idx]).trim();
              if (valStr !== "") return valStr;
            }
            return fallback;
          };

          const rawSira = getVal(siraColIdx, siraColIdx < 0 && rawRow[0] !== undefined ? String(rawRow[0]).trim() : "");
          const rawName = getVal(nameColIdx, nameColIdx < 0 && rawRow[1] !== undefined ? String(rawRow[1]).trim() : (rawRow[0] && !/^\d+$/.test(String(rawRow[0]).trim()) ? String(rawRow[0]).trim() : ""));
          const rawPn = getVal(pnColIdx, pnColIdx < 0 && rawRow[2] !== undefined ? String(rawRow[2]).trim() : "");
          const rawSn = getVal(snColIdx, snColIdx < 0 && rawRow[3] !== undefined ? String(rawRow[3]).trim() : "-");
          const rawMiktar = getVal(miktarColIdx, miktarColIdx < 0 && rawRow[4] !== undefined ? String(rawRow[4]).trim() : "1");
          const rawLoc = getVal(locColIdx, locColIdx < 0 && rawRow[5] !== undefined ? String(rawRow[5]).trim() : (isDepo ? "DEPO" : "Y/D HANGAR"));
          const rawDurum = getVal(durumColIdx, durumColIdx < 0 && rawRow[6] !== undefined ? String(rawRow[6]).trim() : "FAAL");
          
          let rawKalib = getVal(kalibTabiColIdx, kalibTabiColIdx < 0 && rawRow[7] !== undefined ? String(rawRow[7]).trim() : "");
          if (!rawKalib) {
            rawKalib = isDepo ? "EVET" : "EVET";
          }

          const rawSonBakim = cleanAndFormatDateString(getVal(sonBakimColIdx, sonBakimColIdx < 0 && rawRow[8] !== undefined ? String(rawRow[8]).trim() : ""));
          const rawGelecekBakim = cleanAndFormatDateString(getVal(gelecekBakimColIdx, gelecekBakimColIdx < 0 && rawRow[9] !== undefined ? String(rawRow[9]).trim() : ""));
          const rawFirma = getVal(firmaColIdx, firmaColIdx < 0 && rawRow[10] !== undefined ? String(rawRow[10]).trim() : "-");
          
          let rawAciklama = "";
          if (aciklamaColIdx >= 0 && rawRow[aciklamaColIdx] !== undefined) {
            rawAciklama = String(rawRow[aciklamaColIdx]).trim();
          } else {
            const fallbackIdx = finalHeaders.length >= 12 ? 11 : 10;
            if (rawRow[fallbackIdx] !== undefined) {
              rawAciklama = String(rawRow[fallbackIdx]).trim();
            }
          }

          const rawMail = getVal(mailColIdx, "");

          let targetRow: string[] = [];
          if (isDepo) {
            // DEPO STANDART 14 KOLONLU MATRİS FORMATI:
            // 0: SIRA NO, 1: MALZEME / PARÇA ADI, 2: P/N, 3: S/N, 4: MİKTAR, 5: BULUNDUĞU YER,
            // 6: DURUMU, 7: ÖMÜRLÜ PARÇA MI?, 8: ÖMÜR BİTİŞ TARİHİ, 9: TEDARİK EDİLEN FİRMA,
            // 10: AÇIKLAMA, 11: 90 GÜN MAİL, 12: BÖLÜM ETİKETİ, 13: BÖLÜM ETİKETİ
            targetRow = [
              rawSira,
              rawName,
              rawPn || "-",
              rawSn || "-",
              rawMiktar || "1",
              rawLoc || "DEPO",
              rawDurum || "FAAL",
              rawKalib || "EVET",
              rawGelecekBakim || "-",
              rawFirma || "-",
              rawAciklama,
              rawMail,
              finalSection,
              finalSection
            ];
          } else if (isKara) {
            // KARA ARAÇLARI FORMATI:
            // 0: SIRA NO, 1: PLAKA, 2: MARKA, 3: MODEL, 4: BULUNDUĞU YER, 5: SON KM, 6: DURUMU,
            // 7: BAKIMA TABİ, 8: SON KONTROL, 9: GELECEK KONTROL, 10: FİRMA, 11: AÇIKLAMA, 12: MAİL, 13: BÖLÜM
            targetRow = [
              rawSira,
              rawName,
              "-",
              rawPn || "-",
              rawLoc || "ANKARA",
              rawMiktar || rawSonBakim || "",
              rawDurum || "FAAL",
              rawKalib || "EVET",
              rawSonBakim,
              rawGelecekBakim,
              rawFirma,
              rawAciklama,
              rawMail,
              "kara_araclari"
            ];
          } else {
            // YER DESTEK & ÖZEL ALETLER FORMATI:
            // 0: SIRA NO, 1: TEÇHİZAT ADI, 2: P/N, 3: S/N, 4: MİKTAR, 5: BULUNDUĞU YER,
            // 6: DURUMU, 7: KALİBRASYONA TABİ, 8: SON KONTROL, 9: GELECEK KONTROL, 10: FİRMA, 11: AÇIKLAMA,
            // 12: 90 GÜN MAİL, 13: BÖLÜM ETİKETİ
            targetRow = [
              rawSira,
              rawName,
              rawPn || "-",
              rawSn || "-",
              rawMiktar || "1",
              rawLoc || "Y/D HANGAR",
              rawDurum || "FAAL",
              rawKalib || "EVET",
              rawSonBakim,
              rawGelecekBakim,
              rawFirma,
              rawAciklama,
              rawMail,
              finalSection
            ];
          }

          if (targetRow[1] || targetRow[2] || (targetRow[3] && targetRow[3] !== "-") || targetRow[5]) {
            rawParsedRows.push(targetRow);
          }
        }

        // 4. Alt lokasyon ve birleştirilmiş hücreleri grupla
        const grouped = groupMultiLocationRowsHelper(
          rawParsedRows,
          1,
          isKara ? 4 : 5,
          isKara ? -1 : 4,
          0,
          isKara ? 3 : 2,
          isKara ? -1 : 3
        );

        // 5. Sıra no 1'den başlayarak ardışık indexle ve kategori etiketini sabitle
        const finalResult = grouped.map((r, i) => {
          const rowCopy = [...r];
          rowCopy[0] = String(i + 1);
          while (rowCopy.length < 14) rowCopy.push("");
          rowCopy[12] = finalSection;
          rowCopy[13] = finalSection;
          return rowCopy;
        });

        resolve(finalResult);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};
