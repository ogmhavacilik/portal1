with open('src/components/DepoManagementModal.tsx', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add handleDownloadFilteredExcel function right before handleExcelUpload
target_anchor = "  // EXCEL YÜKLE - MEVCUT VERİ YÜKLEME GÜNCELLEME İLE BİREBİR UYUMLU VE DRIVE ENTEGRASYONU"

new_func = """  // FİLTRELENMİŞ EXCEL İNDİR - YALNIZCA SEÇİLİ BÖLGE SÜTUNLARINI VE FİLTRELENMİŞ SATIRLARI İNDİRİR
  const handleDownloadFilteredExcel = () => {
    const activeUnit = (currentUnit || 'at802').toUpperCase();
    const reg = (filterRegion || 'ALL').toUpperCase().trim();
    const isKimyasal = currentDepoType === 'kimyasal';
    const rowsToExport = isKimyasal ? filteredKimyasalRows : filteredSarfRows;

    if (rowsToExport.length === 0) {
      alert("İndirilecek filtrelenmiş veri bulunamadı. Lütfen filtrelerinizi kontrol ediniz.");
      return;
    }

    const aoaData: (string | number)[][] = [];

    if (isKimyasal) {
      // Kimyasal Depo Başlıkları (Yalnızca seçili bölge sütunlarını içerir)
      const headerRow: string[] = [
        "No",
        "Description",
        "P/N - SPAIN DISCRIPTION",
        "MİKTAR",
        "QTY SÖZLEŞME TOPLAMI",
        "PİYASA GELEN",
        "SPAIN GELEN",
        "GENEL TOPLAM",
        "TOPLAM STOK"
      ];

      if (reg === 'ALL' || reg === 'ANKARA') {
        headerRow.push("ANKARA ÇIKAN", "ANKARA MEVCUT");
      }
      if (reg === 'ALL' || reg === 'KARAİN' || reg === 'KARAIN') {
        headerRow.push("KARAİN TRANSFER", "KARAİN ÇIKAN", "KARAİN MEVCUT");
      }
      if (reg === 'ALL' || reg === 'ÇANAKKALE' || reg === 'CANAKKALE') {
        headerRow.push("ÇANAKKALE TRANSFER", "ÇANAKKALE ÇIKAN", "ÇANAKKALE MEVCUT");
      }
      if (reg === 'ALL' || reg === 'MİLAS' || reg === 'MILAS') {
        headerRow.push("MİLAS TRANSFER", "MİLAS ÇIKAN", "MİLAS MEVCUT");
      }
      if (reg === 'ALL' || reg === 'BURSA') {
        headerRow.push("BURSA TRANSFER", "BURSA ÇIKAN", "BURSA MEVCUT");
      }

      headerRow.push("AÇIKLAMA", "RAF ÖMRÜ VAR MI?", "RAF ÖMRÜ BİTİŞ TARİHİ");
      aoaData.push(headerRow);

      // Veri Satırları
      rowsToExport.forEach((item, idx) => {
        const row: (string | number)[] = [
          idx + 1,
          item.description || item.name || '',
          item.spainDescription || item.partNumber || item.pn || '-',
          item.miktarQty || 'Adet',
          item.sozlesmeToplami || '0',
          item.piyasa || 0,
          item.spainGelen || 0,
          item.genelToplam || 0,
          item.toplamStok || 0
        ];

        if (reg === 'ALL' || reg === 'ANKARA') {
          row.push(item.ankaraCikan || 0, item.ankaraMevcut || 0);
        }
        if (reg === 'ALL' || reg === 'KARAİN' || reg === 'KARAIN') {
          row.push(item.karainTransfer || 0, item.karainCikan || 0, item.karainMevcut || 0);
        }
        if (reg === 'ALL' || reg === 'ÇANAKKALE' || reg === 'CANAKKALE') {
          row.push(item.canakkaleTransfer || 0, item.canakkaleCikan || 0, item.canakkaleMevcut || 0);
        }
        if (reg === 'ALL' || reg === 'MİLAS' || reg === 'MILAS') {
          row.push(item.milasTransfer || 0, item.milasCikan || 0, item.milasMevcut || 0);
        }
        if (reg === 'ALL' || reg === 'BURSA') {
          row.push(item.bursaTransfer || 0, item.bursaCikan || 0, item.bursaMevcut || 0);
        }

        const hasLife = item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET');
        row.push(
          item.serialAndNotes || item.sn || '-',
          hasLife ? 'EVET' : 'HAYIR',
          hasLife && item.shelfLifeDate && item.shelfLifeDate !== '-' ? item.shelfLifeDate : '-'
        );

        aoaData.push(row);
      });
    } else {
      // Sarf & Parça Deposu Başlıkları (Yalnızca seçili bölge sütunlarını içerir)
      const headerRow: string[] = [
        "SIRA NO",
        "DESCRIPTION",
        "PART NUMBER",
        "SERİ NUMBER - MÜKERRER NO - AÇIKLAMA",
        "LOKASYON NO",
        "GELEN",
        "TOPLAM STOK"
      ];

      if (reg === 'ALL' || reg === 'ANKARA') {
        headerRow.push("ANKARA ÇIKAN", "ANKARA MEVCUT");
      }
      if (reg === 'ALL' || reg === 'KARAİN' || reg === 'KARAIN') {
        headerRow.push("KARAİN TRANSFER", "KARAİN ÇIKAN", "KARAİN MEVCUT");
      }
      if (reg === 'ALL' || reg === 'ÇANAKKALE' || reg === 'CANAKKALE') {
        headerRow.push("ÇANAKKALE TRANSFER", "ÇANAKKALE ÇIKAN", "ÇANAKKALE MEVCUT");
      }
      if (reg === 'ALL' || reg === 'MİLAS' || reg === 'MILAS') {
        headerRow.push("MİLAS TRANSFER", "MİLAS ÇIKAN", "MİLAS MEVCUT");
      }
      if (reg === 'ALL' || reg === 'BURSA') {
        headerRow.push("BURSA TRANSFER", "BURSA ÇIKAN", "BURSA MEVCUT");
      }

      if (reg === 'ALL') {
        headerRow.push("MUAYENE GİDEN", "MUAYENE GELEN", "MUAYENE TOPLAM");
      }

      headerRow.push("RAF ÖMRÜ VAR MI?", "RAF ÖMRÜ BİTİŞ TARİHİ");
      aoaData.push(headerRow);

      // Veri Satırları
      rowsToExport.forEach((item, idx) => {
        const row: (string | number)[] = [
          idx + 1,
          item.description || item.name || '',
          item.partNumber || item.pn || '-',
          item.serialAndNotes || item.sn || '-',
          item.lokasyonNo || item.location || '-',
          item.gelen || 0,
          item.toplamStok || 0
        ];

        if (reg === 'ALL' || reg === 'ANKARA') {
          row.push(item.ankaraCikan || 0, item.ankaraMevcut || 0);
        }
        if (reg === 'ALL' || reg === 'KARAİN' || reg === 'KARAIN') {
          row.push(item.karainTransfer || 0, item.karainCikan || 0, item.karainMevcut || 0);
        }
        if (reg === 'ALL' || reg === 'ÇANAKKALE' || reg === 'CANAKKALE') {
          row.push(item.canakkaleTransfer || 0, item.canakkaleCikan || 0, item.canakkaleMevcut || 0);
        }
        if (reg === 'ALL' || reg === 'MİLAS' || reg === 'MILAS') {
          row.push(item.milasTransfer || 0, item.milasCikan || 0, item.milasMevcut || 0);
        }
        if (reg === 'ALL' || reg === 'BURSA') {
          row.push(item.bursaTransfer || 0, item.bursaCikan || 0, item.bursaMevcut || 0);
        }

        if (reg === 'ALL') {
          row.push(item.muayeneGiden || 0, item.muayeneGelen || 0, item.muayeneToplam || 0);
        }

        const hasLife = item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET');
        row.push(
          hasLife ? 'EVET' : 'HAYIR',
          hasLife && item.shelfLifeDate && item.shelfLifeDate !== '-' ? item.shelfLifeDate : '-'
        );

        aoaData.push(row);
      });
    }

    const ws = XLSX.utils.aoa_to_sheet(aoaData);
    const wb = XLSX.utils.book_new();
    const sheetName = isKimyasal ? "KIMYASAL_FILTRELENMIS" : "SARF_FILTRELENMIS";
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    let fileName = `${activeUnit}_${currentDepoType.toUpperCase()}`;
    if (reg !== 'ALL') {
      fileName += `_${reg}`;
    }
    if (selectedLocations.length > 0) {
      fileName += `_LOKASYON_FILTRELENMIS`;
    }
    fileName += `_${new Date().toISOString().slice(0, 10)}.xlsx`;

    XLSX.writeFile(wb, fileName);
    showNotification(`✅ ${rowsToExport.length} filtrelenmiş malzeme "${fileName}" olarak bilgisayarınıza indirildi.`);
  };\n\n`;

if (target_anchor in content && !content.includes('handleDownloadFilteredExcel')) {
    content = content.replace(target_anchor, new_func + target_anchor);
    print("Added handleDownloadFilteredExcel function");
}

# 2. Add button in top bar next to Drive Canlı Senkronize Et
old_bar_buttons = """            <button 
              type="button"
              onClick={() => loadLiveDriveExcel(true)}
              disabled={isLoadingDrive}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-[#0b3d1d] hover:bg-emerald-900 disabled:opacity-50 text-white rounded-xl text-xs font-bold transition shadow-xs cursor-pointer active:scale-95"
              title="Google Drive üzerindeki Excel dosyasından anında canlı verileri tazeler"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLoadingDrive ? 'animate-spin' : ''}`} />
              <span>{isLoadingDrive ? 'Drive Bağlanıyor...' : 'Drive Canlı Senkronize Et'}</span>
            </button>"""

new_bar_buttons = """            <button 
              type="button"
              onClick={handleDownloadFilteredExcel}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white rounded-xl text-xs font-black transition shadow-xs cursor-pointer active:scale-95 font-sans"
              title="Ekrandaki filtrelenmiş malzemeleri ve seçili bölgeye ait sütunları Excel olarak bilgisayarınıza indirir"
            >
              <Download className="w-3.5 h-3.5 text-emerald-200" />
              <span>Filtrelenmiş Excel İndir</span>
            </button>

            <button 
              type="button"
              onClick={() => loadLiveDriveExcel(true)}
              disabled={isLoadingDrive}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-[#0b3d1d] hover:bg-emerald-900 disabled:opacity-50 text-white rounded-xl text-xs font-bold transition shadow-xs cursor-pointer active:scale-95"
              title="Google Drive üzerindeki Excel dosyasından anında canlı verileri tazeler"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLoadingDrive ? 'animate-spin' : ''}`} />
              <span>{isLoadingDrive ? 'Drive Bağlanıyor...' : 'Drive Canlı Senkronize Et'}</span>
            </button>"""

if old_bar_buttons in content:
    content = content.replace(old_bar_buttons, new_bar_buttons)
    print("Added Filtrelenmiş Excel İndir button to top bar")

with open('src/components/DepoManagementModal.tsx', 'w', encoding='utf-8') as f:
    f.write(content)
