/**
 * ==============================================================================
 * HAVA ARAÇLARI BAKIM TEKNİK ŞUBE MÜDÜRLÜĞÜ - E-TABLO SENKRONİZASYON MOTORU (V6)
 * ==============================================================================
 * Dosya Adı: komut.gs
 * Google Drive Hedef Klasör ID (PDF): 1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP
 * 
 * Bu script Google E-Tablo script editörüne yapıştırılmalı ve "Web Uygulaması" 
 * olarak tüm kullanıcıların (Herkes / Anyone) erişimine açık şekilde yayınlanmalıdır.
 * ==============================================================================
 */

/**
 * E-Tablo açıldığında özel yönetim menüsünü üst bara ekler.
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('🚁 HA BAKIM SENKRON')
      .addItem('📥 Çevrimdışı Portala Excel Olarak İndir', 'exportAndDownloadExcel')
      .addSeparator()
      .addItem('⏰ Günlük E-Posta Hatırlatıcıyı Etkinleştir', 'createDailyTrigger')
      .addSeparator()
      .addItem('💡 Entegrasyon Kılavuzu', 'showIntegrationGuide')
      .addToUi();
  } catch (err) {
    Logger.log("Arayüz (UI) bağlamına erişilemedi: " + err.toString());
  }
}

/**
 * Gelen POST isteklerini karşılar ve canlı verileri e-tabloya kaydeder, PDF'leri Drive'a yükler.
 */
function doPost(e) {
  var response = { status: "success", timestamp: new Date().toLocaleTimeString('tr-TR') };
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Her işlem öncesi Gün Takip sayfasının varlığından emin olalım
    ensureGunTakipSheet(ss);
    
    var now = new Date();
    var pad = function(n) { return String(n).padStart(2, '0'); };
    var formattedDate = pad(now.getDate()) + "." + pad(now.getMonth() + 1) + "." + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes());

    if (action === "updateSheet") {
      var sheetName = postData.sheetName;
      var data = postData.data;
      var sheet = getSheetWithFallback(ss, sheetName);
      sheet.clear();
      
      if (data && data.length > 0) {
        if (Array.isArray(data[0])) {
          var maxCols = 0;
          data.forEach(function(row) {
            if (Array.isArray(row) && row.length > maxCols) { maxCols = row.length; }
          });
          
          if (maxCols > 0) {
            var values = data.map(function(row) {
              var newRow = [];
              for (var i = 0; i < maxCols; i++) {
                var cellVal = (row && row[i] !== undefined && row[i] !== null) ? row[i] : "";
                if (typeof cellVal === "string" && !isNaN(cellVal) && cellVal.trim() !== "") {
                  var trimmed = cellVal.trim();
                  if (!/^0\d+/.test(trimmed) && trimmed.length > 0) {
                    cellVal = Number(trimmed);
                  }
                }
                newRow.push(cellVal);
              }
              return newRow;
            });
            sheet.getRange(1, 1, values.length, maxCols).setValues(values);
            
            // Format sheet based on type
            if (sheet.getName().toLowerCase().indexOf("1-gorevlendirme") !== -1) {
              formatGorevlendirmeSheet(sheet);
            } else {
              formatGeneralSheet(sheet);
            }
          }
        }
      }
      
      recordLastUpdate(ss, getUnitTitleByPrefix(sheetName), formattedDate);
      response.message = sheet.getName() + " sayfası başarıyla güncellendi.";

    } else if (action === "updatePdfTimestamp") {
      var formId = Number(postData.formId);
      var month = postData.month;
      var unitName = postData.unitName || getUnitTitleById(formId, month);
      
      recordLastUpdate(ss, unitName, formattedDate);
      response.message = "'" + unitName + "' planlaması için güncelleme tarihi başarıyla e-tabloya kaydedildi.";
      response.uploadDate = formattedDate;

    } else if (action === "uploadPdfToDrive") {
      var fileName = postData.fileName;
      var base64Data = postData.base64Data;
      var formId = postData.formId ? Number(postData.formId) : null;
      var month = postData.month || null;
      
      var folderId = postData.folderId;
      if (!folderId) {
        var fileNameLower = String(fileName || "").toLowerCase();
        var monthStr = String(month || "");
        var formIdNum = formId ? Number(formId) : null;
        if (fileNameLower.indexOf("tech_img_") === 0 || monthStr === "Teçhizat Takip" || formIdNum === 6) {
          folderId = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"; // Teçhizat Takip Klasörü
        } else {
          folderId = "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0"; // Form Kayıtları Klasörü
        }
      }
      var folder = DriveApp.getFolderById(folderId);
      
      // Aynı isimde dosya varsa üzerine yazmak için eski dosyayı Çöp Kutusuna (trash) gönder
      var existingFiles = folder.getFilesByName(fileName);
      while (existingFiles.hasNext()) {
        var file = existingFiles.next();
        try {
          file.setTrashed(true);
        } catch (delErr) {
          Logger.log("Eski dosya temizlenemedi: " + delErr.toString());
        }
      }
      
      var mimeType = postData.mimeType;
      if (!mimeType) {
        var fileNameLower = fileName.toLowerCase();
        if (fileNameLower.indexOf(".xlsx") !== -1) mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        else if (fileNameLower.indexOf(".xls") !== -1) mimeType = "application/vnd.ms-excel";
        else if (fileNameLower.indexOf(".csv") !== -1) mimeType = "text/csv";
        else mimeType = "application/pdf";
      }
      
      var decoded = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(decoded, mimeType, fileName);
      var newFile = folder.createFile(blob);
      
      // Hangi hava aracının hangi ayı olduğunu tespit et ve güncelleme tarihini yaz
      var derivedFormId = formId;
      var derivedMonth = month;
      
      if (!derivedFormId) {
        derivedFormId = 21;
        var fileNameLower = fileName.toLowerCase();
        if (fileNameLower.indexOf("bell429") !== -1) derivedFormId = 21;
        else if (fileNameLower.indexOf("t70") !== -1) derivedFormId = 22;
        else if (fileNameLower.indexOf("at802") !== -1) derivedFormId = 23;
        else if (fileNameLower.indexOf("gorevlendirme") !== -1) derivedFormId = 1;
        else if (fileNameLower.indexOf("yetki") !== -1) derivedFormId = 3;
        else if (fileNameLower.indexOf("bilgi") !== -1) derivedFormId = 5;
        else if (fileNameLower.indexOf("ucus") !== -1) derivedFormId = 6;
      }
      
      if (!derivedMonth) {
        derivedMonth = "Genel Plan";
        var fileNameLower = fileName.toLowerCase();
        var months = ["haziran", "temmuz", "agustos", "eylul", "ekim", "kasim", "aralik", "ocak", "subat", "mart", "nisan", "mayis"];
        for (var mIdx = 0; mIdx < months.length; mIdx++) {
          if (fileNameLower.indexOf(months[mIdx]) !== -1) {
            derivedMonth = months[mIdx].charAt(0).toUpperCase() + months[mIdx].slice(1) + " 2026";
            break;
          }
        }
      }
      
      var unitName = postData.unitName || getUnitTitleById(derivedFormId, derivedMonth);
      recordLastUpdate(ss, unitName, formattedDate);
      
      response.message = "PDF belgesi '" + fileName + "' adıyla Google Drive'a başarıyla yüklendi ve e-tablo güncelleme tarihi yenilendi.";
      response.fileId = newFile.getId();
      response.viewUrl = "https://drive.google.com/file/d/" + newFile.getId() + "/preview";
      response.unitName = unitName;
      response.uploadDate = formattedDate;

    } else if (action === "updateMultiSheets") {
      var prefix = postData.prefix;
      var sheetsList = postData.sheets;
      var tempSheet = ss.insertSheet("__temp_ha_bakim_" + Math.floor(Math.random() * 10000));
      
      var sheetsInDoc = ss.getSheets();
      var prefixLower = prefix.toLowerCase();
      for (var i = sheetsInDoc.length - 1; i >= 0; i--) {
        var shName = sheetsInDoc[i].getName();
        var nameLower = shName.toLowerCase();
        if (nameLower === prefixLower || nameLower.indexOf(prefixLower + "-") === 0) {
          try {
            ss.deleteSheet(sheetsInDoc[i]);
          } catch (delErr) {}
        }
      }
      
      var createdList = [];
      for (var shIdx = 0; shIdx < sheetsList.length; shIdx++) {
        var sheetObj = sheetsList[shIdx];
        var newSheet = ss.insertSheet(sheetObj.name);
        createdList.push(sheetObj.name);
        
        var sheetData = sheetObj.data;
        if (sheetData && sheetData.length > 0) {
          var maxCols = 0;
          sheetData.forEach(function(r) {
            if (Array.isArray(r) && r.length > maxCols) { maxCols = r.length; }
          });
          
          if (maxCols > 0) {
            var values = sheetData.map(function(row) {
              var newRow = [];
              for (var c = 0; c < maxCols; c++) {
                var cellVal = (row && row[c] !== undefined && row[c] !== null) ? row[c] : "";
                if (typeof cellVal === "string" && !isNaN(cellVal) && cellVal.trim() !== "") {
                  var trimmed = cellVal.trim();
                  if (!/^0\d+/.test(trimmed) && trimmed.length > 0) {
                    cellVal = Number(trimmed);
                  }
                }
                newRow.push(cellVal);
              }
              return newRow;
            });
            newSheet.getRange(1, 1, values.length, maxCols).setValues(values);
            
            if (prefixLower.indexOf("1-gorevlendirme") !== -1) {
              formatGorevlendirmeSheet(newSheet);
            } else {
              formatGeneralSheet(newSheet);
            }
          }
        }
      }
      
      try {
        ss.deleteSheet(tempSheet);
      } catch (e) {}
      
      recordLastUpdate(ss, getUnitTitleByPrefix(prefix), formattedDate);
      response.message = "Toplam " + sheetsList.length + " sayfa '" + prefix + "' ön ekiyle güncellendi.";
      response.updatedSheets = createdList;
    } else if (action === "updateTumTechizat") {
      var unitLabel = postData.unitLabel;
      var data = postData.data;
      var targetSs = null;
      var targetSsId = postData.spreadsheetId;
      if (targetSsId) {
        try {
          targetSs = SpreadsheetApp.openById(targetSsId);
        } catch (err) {
          Logger.log("openById failed: " + err.toString());
        }
      }
      if (!targetSs) {
        var fallbackId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
        try {
          targetSs = SpreadsheetApp.openById(fallbackId);
        } catch (err) {
          Logger.log("openById fallback failed: " + err.toString());
        }
      }
      if (!targetSs) {
        targetSs = ss;
      }
      var sheet = getSheetWithFallback(targetSs, "TÜM TECHİZAT");
      
      var lastCol = sheet.getLastColumn();
      var headers = [
        "AİT OLDUĞU BİRİM",
        "SIRA NO",
        "TEÇHİZAT ADI",
        "ARAÇ PLAKASI / TANIMI",
        "PARÇA NO (P/N)",
        "MODEL",
        "SERİ NO (S/N)",
        "MİKTAR / KAPASİTE",
        "SON KM Sİ",
        "BULUNDUĞU YER",
        "DURUMU",
        "SON KONTROL / BAKIM",
        "GELECEK KONTROL / BAKIM",
        "SON KONTROLÜ YAPAN FİRMA",
        "AÇIKLAMA",
        "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ",
        "60 GÜN UYARISI MAİL GÖNDERİM TARİHİ",
        "30 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
      ];
      
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(headers);
        sheet.getRange("A1:R1").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff").setHorizontalAlignment("center");
      } else {
        var currentLastCol = sheet.getLastColumn();
        if (currentLastCol < 18) {
          // Reset headers to guarantee all 18 columns exist
          sheet.getRange(1, 1, 1, 18).setValues([headers]);
          sheet.getRange("A1:R1").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff").setHorizontalAlignment("center");
        }
      }
      
      var lastRow = sheet.getLastRow();
      var emailSentTracker = {}; // key: "Birim_CihazAdı_SeriNo" -> { mail90: "", mail60: "", mail30: "", gelecekBakim: "" }
      
      if (lastRow > 1) {
        var numCols = sheet.getLastColumn();
        if (numCols >= 18) {
          var entireRange = sheet.getRange(2, 1, lastRow - 1, numCols);
          var entireValues = entireRange.getValues();
          for (var r = 0; r < entireValues.length; r++) {
            var rowUnit = String(entireValues[r][0]).trim().toUpperCase();
            var isKara = (rowUnit === "KARA ARAÇLARI" || rowUnit.indexOf("KARA") !== -1);
            
            var eqName = isKara ? String(entireValues[r][3]).trim().toUpperCase() : String(entireValues[r][2]).trim().toUpperCase();
            var serNo = isKara ? "" : String(entireValues[r][6]).trim().toUpperCase();
            var savedGelecekBakim = (entireValues[r][12] !== undefined && entireValues[r][12] !== null) ? String(entireValues[r][12]).trim() : "";
            
            var m90 = (entireValues[r][15] !== undefined && entireValues[r][15] !== null) ? String(entireValues[r][15]).trim() : "";
            var m60 = (entireValues[r][16] !== undefined && entireValues[r][16] !== null) ? String(entireValues[r][16]).trim() : "";
            var m30 = (entireValues[r][17] !== undefined && entireValues[r][17] !== null) ? String(entireValues[r][17]).trim() : "";
            
            var trackerKey = rowUnit + "_" + eqName + "_" + serNo;
            emailSentTracker[trackerKey] = { mail90: m90, mail60: m60, mail30: m30, gelecekBakim: savedGelecekBakim };
          }
        }
        
        // Delete rows matching current unit
        var sheetValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var r = sheetValues.length - 1; r >= 0; r--) {
          var rowUnit = String(sheetValues[r][0]).trim().toUpperCase();
          if (rowUnit === unitLabel.toUpperCase()) {
            sheet.deleteRow(r + 2);
          }
        }
      }
      
      if (data && data.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        var numRows = data.length;
        var numCols = 18;
        
        var valuesToInsert = data.map(function(row) {
          var newRow = [];
          var rowUnit = String(row[0]).trim().toUpperCase();
          var isKara = (rowUnit === "KARA ARAÇLARI" || rowUnit.indexOf("KARA") !== -1);
          
          if (isKara) {
            // [unitLabel, SIRA_NO, PLAKA, MODEL, YER, km, DURUMU, SON_KONTROL, GELECEK_KONTROL, FİRMA, AÇIKLAMA, MAIL]
            newRow.push("KARA ARAÇLARI"); // AİT OLDUĞU BİRİM
            newRow.push(row[1] !== undefined && row[1] !== null ? String(row[1]) : ""); // SIRA NO
            newRow.push("-"); // TEÇHİZAT ADI
            newRow.push(row[2] !== undefined && row[2] !== null ? String(row[2]) : ""); // ARAÇ PLAKASI / TANIMI
            newRow.push("-"); // PARÇA NO (P/N)
            newRow.push(row[3] !== undefined && row[3] !== null ? String(row[3]) : ""); // MODEL
            newRow.push("-"); // SERİ NO (S/N)
            newRow.push("-"); // MİKTAR / KAPASİTE
            newRow.push(row[5] !== undefined && row[5] !== null ? String(row[5]) : ""); // SON KM Sİ
            newRow.push(row[4] !== undefined && row[4] !== null ? String(row[4]) : ""); // BULUNDUĞU YER
            newRow.push(row[6] !== undefined && row[6] !== null ? String(row[6]) : ""); // DURUMU
            newRow.push(row[7] !== undefined && row[7] !== null ? String(row[7]) : ""); // SON KONTROL / BAKIM
            newRow.push(row[8] !== undefined && row[8] !== null ? String(row[8]) : ""); // GELECEK KONTROL / BAKIM
            newRow.push(row[9] !== undefined && row[9] !== null ? String(row[9]) : ""); // SON KONTROLÜ YAPAN FİRMA
            newRow.push(row[10] !== undefined && row[10] !== null ? String(row[10]) : ""); // AÇIKLAMA
          } else {
            // [unitLabel, SIRA_NO, TEÇHİZAT ADI, PARÇA_NO, SERİ_NO, MİKTAR, YER, DURUMU, SON_KONTROL, GELECEK_KONTROL, FİRMA, AÇIKLAMA]
            newRow.push(rowUnit); // AİT OLDUĞU BİRİM
            newRow.push(row[1] !== undefined && row[1] !== null ? String(row[1]) : ""); // SIRA NO
            newRow.push(row[2] !== undefined && row[2] !== null ? String(row[2]) : ""); // TEÇHİZAT ADI
            newRow.push("-"); // ARAÇ PLAKASI / TANIMI
            newRow.push(row[3] !== undefined && row[3] !== null ? String(row[3]) : ""); // PARÇA NO (P/N)
            newRow.push("-"); // MODEL
            newRow.push(row[4] !== undefined && row[4] !== null ? String(row[4]) : ""); // SERİ NO (S/N)
            newRow.push(row[5] !== undefined && row[5] !== null ? String(row[5]) : ""); // MİKTAR / KAPASİTE
            newRow.push("-"); // SON KM Sİ
            newRow.push(row[6] !== undefined && row[6] !== null ? String(row[6]) : ""); // BULUNDUĞU YER
            newRow.push(row[7] !== undefined && row[7] !== null ? String(row[7]) : ""); // DURUMU
            newRow.push(row[8] !== undefined && row[8] !== null ? String(row[8]) : ""); // SON KONTROL / BAKIM
            newRow.push(row[9] !== undefined && row[9] !== null ? String(row[9]) : ""); // GELECEK KONTROL / BAKIM
            newRow.push(row[10] !== undefined && row[10] !== null ? String(row[10]) : ""); // SON KONTROLÜ YAPAN FİRMA
            newRow.push(row[11] !== undefined && row[11] !== null ? String(row[11]) : ""); // AÇIKLAMA
          }
          
          var eqName = isKara ? String(row[2] || "").trim().toUpperCase() : String(row[2] || "").trim().toUpperCase();
          var serNo = isKara ? "" : String(row[4] || "").trim().toUpperCase();
          var newGelecekBakim = isKara ? String(row[8] || "").trim() : String(row[9] || "").trim();
          var trackerKey = rowUnit + "_" + eqName + "_" + serNo;
          
          var saved = emailSentTracker[trackerKey] || { mail90: "", mail60: "", mail30: "", gelecekBakim: "" };
          
          if (saved.gelecekBakim && saved.gelecekBakim !== newGelecekBakim) {
            newRow.push("");
            newRow.push("");
            newRow.push("");
          } else {
            newRow.push(saved.mail90);
            newRow.push(saved.mail60);
            newRow.push(saved.mail30);
          }
          
          return newRow;
        });
        
        sheet.getRange(startRow, 1, numRows, numCols).setValues(valuesToInsert);
      }
      
      formatGeneralSheet(sheet);
      
      var targetUnitTitle = unitLabel.toUpperCase();
      if (targetUnitTitle === "BELL 429") targetUnitTitle = "BELL 429 YER DESTEK TEÇHİZATLARI";
      else if (targetUnitTitle === "AT-802F" || targetUnitTitle === "AT-802") targetUnitTitle = "AT-802F YER DESTEK TEÇHİZATLARI";
      else if (targetUnitTitle === "T-70 YER DESTEK" || targetUnitTitle === "T-70") targetUnitTitle = "T-70 YER DESTEK TEÇHİZATI";
      else if (targetUnitTitle === "T-70 BUMBİ BACKET") targetUnitTitle = "T-70 BUMBİ BACKET TEÇHİZATI";
      else if (targetUnitTitle === "B-360") targetUnitTitle = "B-360 YER DESTEK TEÇHİZATLARI";
      else if (targetUnitTitle === "C-650") targetUnitTitle = "C-650 YER DESTEK TEÇHİZATLARI";
      else if (targetUnitTitle === "HANGAR YER DESTEK" || targetUnitTitle === "HANGAR") targetUnitTitle = "HANGAR YER DESTEK TEÇHİZATLARI";
      
      recordLastUpdate(targetSs, targetUnitTitle, formattedDate);
      
      response.message = "TÜM TECHİZAT sayfasındaki '" + unitLabel + "' verileri başarıyla güncellendi.";
    } else if (action === "updateGunTakip") {
      var data = postData.data; // Array of [SORUMLU BİRİM, ADI SOYADI, E-POSTA ADRESİ]
      var targetSs = null;
      var targetSsId = postData.spreadsheetId;
      if (targetSsId) {
        try {
          targetSs = SpreadsheetApp.openById(targetSsId);
        } catch (err) {
          Logger.log("openById failed: " + err.toString());
        }
      }
      if (!targetSs) {
        var fallbackId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
        try {
          targetSs = SpreadsheetApp.openById(fallbackId);
        } catch (err) {
          Logger.log("openById fallback failed: " + err.toString());
        }
      }
      if (!targetSs) {
        targetSs = ss;
      }
      var sheet = getSheetWithFallback(targetSs, "GÜN TAKİP");
      
      var existingDates = {}; // key: BİRİM (UPPERCASE) -> { mail90: string, mail60: string, mail30: string }
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var numColsExisting = sheet.getLastColumn();
        var entireValues = sheet.getRange(2, 1, lastRow - 1, numColsExisting).getValues();
        for (var r = 0; r < entireValues.length; r++) {
          var birim = String(entireValues[r][0]).trim().toUpperCase();
          if (birim) {
            existingDates[birim] = {
              mail90: (entireValues[r][3] !== undefined && entireValues[r][3] !== null) ? String(entireValues[r][3]).trim() : "",
              mail60: (entireValues[r][4] !== undefined && entireValues[r][4] !== null) ? String(entireValues[r][4]).trim() : "",
              mail30: (entireValues[r][5] !== undefined && entireValues[r][5] !== null) ? String(entireValues[r][5]).trim() : ""
            };
          }
        }
      }
      
      sheet.clear();
      
      var headers = [
        "SORUMLU BİRİM", 
        "ADI SOYADI", 
        "E-POSTA ADRESİ", 
        "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ", 
        "60 GÜN UYARISI MAİL GÖNDERİM TARİHİ", 
        "30 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
      ];
      sheet.appendRow(headers);
      sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff").setHorizontalAlignment("center");
      
      if (data && data.length > 0) {
        var startRow = 2;
        var numRows = data.length;
        var numCols = 6;
        
        var valuesToInsert = data.map(function(row) {
          var rowBirim = row[0] !== undefined && row[0] !== null ? String(row[0]).trim() : "";
          var rowBirimUpper = rowBirim.toUpperCase();
          var saved = existingDates[rowBirimUpper] || { mail90: "", mail60: "", mail30: "" };
          
          return [
            rowBirim,
            row[1] !== undefined && row[1] !== null ? String(row[1]).trim() : "",
            row[2] !== undefined && row[2] !== null ? String(row[2]).trim() : "",
            saved.mail90,
            saved.mail60,
            saved.mail30
          ];
        });
        
        sheet.getRange(startRow, 1, numRows, numCols).setValues(valuesToInsert);
      }
      
      formatGeneralSheet(sheet);
      recordLastUpdate(targetSs, "GÜN TAKİP", formattedDate);
      
      // Mail hatırlatıcıyı otomatik tetikleme devre dışı bırakıldı (kullanıcı isteğiyle artık sadece zamanlanmış tetikleyici çalışacak)
      // processEquipmentReminders(ss, formattedDate);
      
      response.message = "GÜN TAKİP (Sorumlu Birim Mail Ayarları) başarıyla güncellendi.";
    } else if (action === "appendEbysTable") {
      var ebysNo = String(postData.ebysNo || "").trim();
      var tableData = postData.data; // Array of [Ait Olduğu Birim, Teçhizat Adı, Parça No, Seri No, Miktar, Son Kontrolü Yapan Firma, Açıklama]
      
      var targetSs = null;
      var targetSsId = postData.spreadsheetId;
      if (targetSsId) {
        try {
          targetSs = SpreadsheetApp.openById(targetSsId);
        } catch (err) {
          Logger.log("openById failed: " + err.toString());
        }
      }
      if (!targetSs) {
        var fallbackId = "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ";
        try {
          targetSs = SpreadsheetApp.openById(fallbackId);
        } catch (err) {
          Logger.log("openById fallback failed: " + err.toString());
        }
      }
      if (!targetSs) {
        targetSs = ss;
      }
      
      var sheet = targetSs.getSheetByName("Sayfa1") || targetSs.getSheets()[0];
      if (!sheet) {
        sheet = targetSs.insertSheet("Sayfa1");
      }
      
      var lastRow = getActualLastRow(sheet);
      var startRow = lastRow === 0 ? 1 : lastRow + 2; // leave exactly 1 blank row
      
      // 1. Write the EBYS Title Row
      sheet.getRange(startRow, 1).setValue("EBYS NO:");
      sheet.getRange(startRow, 2).setValue(ebysNo);
      sheet.getRange(startRow, 3).setValue(String(postData.talepTuru || "MALZEME").trim().toUpperCase());
      
      var titleRange = sheet.getRange(startRow, 1, 1, 2);
      titleRange.setBackground("#ffff00").setFontWeight("bold").setHorizontalAlignment("left");
      
      var matRange = sheet.getRange(startRow, 3);
      matRange.setBackground("#93c5fd").setFontWeight("bold").setHorizontalAlignment("center");
      
      // 2. Write Headers Row
      var headers = ["SIRA NU.", "MALZEME ADI", "PARÇA NUMARASI", "BİRİM", "İSTEK MİKTARI", "TESLİM TARİHİ"];
      sheet.getRange(startRow + 1, 1, 1, 6).setValues([headers]);
      
      var headerRange = sheet.getRange(startRow + 1, 1, 1, 6);
      headerRange.setFontWeight("bold").setBackground("#cbd5e1").setFontColor("#000000").setHorizontalAlignment("center");
      
      // 3. Write Data Rows
      if (tableData && tableData.length > 0) {
        var values = tableData.map(function(row, index) {
          var techName = row[1] || "";
          var parcaNo = row[2] || "-";
          var miktar = row[4] || "1";
          
          // Auto-detect unit (Birim) based on Turkish words in equipment name
          var birim = "ADET";
          var nameLower = techName.toLowerCase();
          if (nameLower.indexOf("set") !== -1 || nameLower.indexOf("takim") !== -1 || nameLower.indexOf("istasyonu") !== -1) {
            birim = "SET";
          } else if (nameLower.indexOf("litre") !== -1 || nameLower.indexOf(" lt") !== -1) {
            birim = "LİTRE";
          } else if (nameLower.indexOf("kutu") !== -1) {
            birim = "KUTU";
          } else if (nameLower.indexOf("paket") !== -1) {
            birim = "PAKET";
          } else if (nameLower.indexOf("rulo") !== -1) {
            birim = "RULO";
          } else if (nameLower.indexOf("metre") !== -1 || nameLower.indexOf(" mt") !== -1) {
            birim = "METRE";
          } else if (nameLower.indexOf("kilo") !== -1 || nameLower.indexOf("kg") !== -1) {
            birim = "KG";
          }
          
          return [
            String(index + 1), // SIRA NU.
            techName,          // MALZEME ADI
            parcaNo,           // PARÇA NUMARASI
            birim,             // BİRİM
            miktar,            // İSTEK MİKTARI
            ""                 // TESLİM TARİHİ (boş bırakılır)
          ];
        });
        sheet.getRange(startRow + 2, 1, values.length, 6).setValues(values);
        
        var dataRange = sheet.getRange(startRow + 2, 1, values.length, 6);
        dataRange.setHorizontalAlignment("center");
        
        // Left align Malzeme Adi for better readability
        sheet.getRange(startRow + 2, 2, values.length, 1).setHorizontalAlignment("left");
      }
      
      // 4. Style entire block
      var blockHeight = 2 + (tableData ? tableData.length : 0);
      var totalRange = sheet.getRange(startRow, 1, blockHeight, 6);
      totalRange.setFontFamily("Calibri").setFontSize(11).setVerticalAlignment("middle");
      
      // Apply thin inside borders and thick outside borders
      try {
        totalRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        totalRange.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      } catch (borderErr) {
        totalRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }
      
      // Auto-resize columns
      for (var colIdx = 1; colIdx <= 6; colIdx++) {
        sheet.autoResizeColumn(colIdx);
        var width = sheet.getColumnWidth(colIdx);
        if (colIdx === 2) { // MALZEME ADI column can be wider
          if (width < 250) sheet.setColumnWidth(colIdx, 250);
          else if (width > 450) sheet.setColumnWidth(colIdx, 450);
        } else {
          if (width < 100) sheet.setColumnWidth(colIdx, 100);
          else if (width > 200) sheet.setColumnWidth(colIdx, 200);
        }
      }
      
      response.message = "EBYS Tablosu Sayfa1 sayfasına başarıyla oluşturuldu.";

    } else if (action === "updateEbysFirma") {
      var ebysNo = String(postData.ebysNo || "").trim();
      var partNo = String(postData.parcaNo || "").trim();
      var seriNo = String(postData.seriNo || "").trim();
      var newFirma = String(postData.firma || "").trim();
      
      var ssId = ss.getId();
      var ssName = ss.getName().toUpperCase();
      var isTaskline = (ssId === "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ" || ssName.indexOf("TASKLINE") !== -1);
      
      var sheet;
      if (isTaskline) {
        sheet = ss.getSheetByName("Sayfa1") || ss.getSheets()[0];
      } else {
        sheet = getSheetWithFallback(ss, "TASKLINE-PARÇA LİSTESİ");
      }
      
      var lastRow = getActualLastRow(sheet);
      var updated = false;
      if (lastRow > 0) {
        var allValues = sheet.getRange(1, 1, lastRow, 7).getValues();
        var currentEbysNo = "";
        
        for (var i = 0; i < allValues.length; i++) {
          var row = allValues[i];
          var cellA = String(row[0] || "").trim();
          
          if (cellA.indexOf("EBYS NO:") === 0) {
            var cellB = String(row[1] || "").trim();
            if (cellB) {
              currentEbysNo = cellB;
            } else {
              currentEbysNo = cellA.replace("EBYS NO:", "").trim();
            }
            continue;
          }
          
          if (currentEbysNo === ebysNo) {
            var rowPartNo = String(row[2] || "").trim();
            var rowSeriNo = String(row[3] || "").trim();
            
            if (rowPartNo === partNo && rowSeriNo === seriNo) {
              sheet.getRange(i + 1, 6).setValue(newFirma);
              updateTechizatFirmaInAllTechizat(ss, partNo, seriNo, newFirma);
              updated = true;
              break;
            }
          }
        }
      }
      
      if (updated) {
        response.message = "Firma başarıyla güncellendi ve teçhizat listesiyle senkronize edildi.";
      } else {
        response.status = "error";
        response.message = "Eşleşen kayıt bulunamadı: " + ebysNo + ", " + partNo + ", " + seriNo;
      }

    } else if (action === "getDemands") {
      var targetSs = null;
      var targetSsId = postData.spreadsheetId;
      if (targetSsId) {
        try {
          targetSs = SpreadsheetApp.openById(targetSsId);
        } catch (err) {
          Logger.log("openById failed: " + err.toString());
        }
      }
      if (!targetSs) {
        var fallbackId = "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ";
        try {
          targetSs = SpreadsheetApp.openById(fallbackId);
        } catch (err) {
          Logger.log("openById fallback failed: " + err.toString());
        }
      }
      if (!targetSs) {
        targetSs = ss;
      }
      var sheet = targetSs.getSheetByName("İşlemdeki Talepler") || targetSs.getSheetByName("Talepler") || targetSs.getSheetByName("Demands") || targetSs.getSheetByName("TASKLINE-PARÇA LİSTESİ") || targetSs.getSheetByName("Sayfa1") || targetSs.getSheets()[0];
      var rows = [];
      if (sheet) {
        var sheetName = sheet.getName();
        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        
        if (sheetName === "TASKLINE-PARÇA LİSTESİ" || sheetName === "Sayfa1" || sheetName.toLowerCase() === "sayfa1" || sheetName.toLowerCase() === "taskline-parca listesi") {
          if (lastRow > 0 && lastCol > 0) {
            var allValues = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 3)).getDisplayValues();
            var seenEbys = {};
            for (var i = 0; i < allValues.length; i++) {
              var row = allValues[i];
              var cellA = String(row[0] || "").trim();
              var cellB = String(row[1] || "").trim();
              var cellC = String(row[2] || "").trim();
              
              if (cellA.indexOf("EBYS NO:") === 0) {
                var eNo = "";
                if (cellB) {
                  eNo = cellB;
                } else {
                  eNo = cellA.replace("EBYS NO:", "").trim();
                }
                
                var digits = eNo.trim();
                if (digits && !seenEbys[digits]) {
                  seenEbys[digits] = true;
                  var talepTuru = cellC || "MALZEME";
                  rows.push({
                    "SIRA NO": String(rows.length + 1),
                    "Başlık": "EBYS Talep " + digits,
                    "Açıklama": talepTuru + " Talebi (" + digits + ")",
                    "Talep Türü": talepTuru,
                    "EBYS NO": digits
                  });
                }
              }
            }
          }
        } else {
          if (lastRow > 1 && lastCol > 0) {
            var allValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
            var headers = allValues[0].map(function(h) { return String(h || "").trim().toUpperCase(); });
            
            var siraIdx = -1, basIdx = -1, aciIdx = -1, turIdx = -1, ebyIdx = -1;
            for (var col = 0; col < headers.length; col++) {
              var h = headers[col];
              if (h.indexOf("SIRA") !== -1 || h === "S.N." || h === "S.NU.") siraIdx = col;
              else if (h.indexOf("BAŞLIK") !== -1 || h.indexOf("BASLIK") !== -1 || h === "KONU") basIdx = col;
              else if (h.indexOf("AÇIKLAMA") !== -1 || h.indexOf("ACIKLAMA") !== -1) aciIdx = col;
              else if (h.indexOf("TÜR") !== -1 || h.indexOf("TURU") !== -1 || h === "TİP") turIdx = col;
              else if (h.indexOf("EBYS") !== -1 && h.indexOf("TARİH") === -1 && h.indexOf("TARIH") === -1 && h.indexOf("DATE") === -1) ebyIdx = col;
            }
            
            for (var rIdx = 1; rIdx < allValues.length; rIdx++) {
              var rowVal = allValues[rIdx];
              var eNo = "";
              if (ebyIdx !== -1) {
                eNo = String(rowVal[ebyIdx] || "").trim();
              }
              if (!eNo || eNo.toLowerCase() === "n/a" || eNo.toLowerCase() === "na") {
                if (rowVal.length > 7) {
                  eNo = String(rowVal[7] || "").trim();
                }
              }
              if (!eNo) continue;
              
              rows.push({
                "SIRA NO": siraIdx !== -1 ? String(rowVal[siraIdx] || "").trim() : String(rIdx),
                "Başlık": basIdx !== -1 ? String(rowVal[basIdx] || "").trim() : ("Talep " + eNo),
                "Açıklama": aciIdx !== -1 ? String(rowVal[aciIdx] || "").trim() : "",
                "Talep Türü": turIdx !== -1 ? String(rowVal[turIdx] || "").trim() : "MALZEME",
                "EBYS NO": eNo
              });
            }
          }
        }
      }
      
      if (rows.length === 0) {
        rows = [
          {
            "SIRA NO": "1",
            "Başlık": "Bell 429 Test Cihazı Kalibrasyon Talebi",
            "Açıklama": "Teknik Şube kalibrasyon laboratuvarı cihaz kalibrasyon işlem talebi",
            "Talep Türü": "KALİBRASYON",
            "EBYS NO": "2408159"
          },
          {
            "SIRA NO": "2",
            "Başlık": "T-70 Telsiz Ölçüm Ekipmanı Periyodik Kontrol",
            "Açıklama": "Aviyonik atölyesi telsiz test cihazının periyodik kontrol işlemi",
            "Talep Türü": "BAKIM",
            "EBYS NO": "2408160"
          },
          {
            "SIRA NO": "3",
            "Başlık": "AT-802F Yangın Söndürme Kit Kontrolü",
            "Açıklama": "Muğla Yangın söndürme kit ekipmanının yıllık bakımı",
            "Talep Türü": "BAKIM",
            "EBYS NO": "2408161"
          },
          {
            "SIRA NO": "4",
            "Başlık": "Hangar Güç Ünitesi Yıllık Kalibrasyon",
            "Açıklama": "Yer destek cihazı jeneratör kalibrasyon ve test işlemi",
            "Talep Türü": "KALİBRASYON",
            "EBYS NO": "2408162"
          }
        ];
      }
      response.data = rows;
      response.message = "Demands list retrieved successfully.";

    } else {
      response.status = "error";
      response.message = "Geçersiz doPost aksiyonu: " + action;
    }
  } catch (err) {
    response.status = "error";
    response.message = err.toString();
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Gelen GET isteklerini karşılar ve e-tablodaki verileri JSON formatında döndürür.
 */
function doGet(e) {
  var response = { status: "success", timestamp: new Date().toLocaleTimeString('tr-TR') };
  try {
    var action = e.parameter.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Her işlem öncesi Gün Takip sayfasının varlığından emin olalım
    ensureGunTakipSheet(ss);
    
    if (action === "getSheets") {
      response.sheets = ss.getSheets().map(function(s) {
        return { name: s.getName(), id: s.getSheetId() };
      });
      response.message = "Sayfalar başarıyla listelendi.";

    } else if (action === "listPdfsFromDrive") {
      var list = [];
      try {
        var folderId = e.parameter.folderId || "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0";
        var folder = DriveApp.getFolderById(folderId);
        var files = folder.getFiles();
        while (files.hasNext()) {
          var file = files.next();
          var name = file.getName();
          if (name.toLowerCase().endsWith(".pdf")) {
            list.push({
              name: name,
              id: file.getId(),
              viewUrl: "https://drive.google.com/file/d/" + file.getId() + "/preview",
              lastUpdated: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss")
            });
          }
        }
      } catch (driveErr) {
        Logger.log("Drive erişim hatası: " + driveErr.toString());
      }
      response.data = list;
      response.message = list.length + " adet planlama PDF'i listelendi.";

    } else if (action === "listImagesFromDrive") {
      var imagesMap = {};
      try {
        var folderId = e.parameter.folderId || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var folder = DriveApp.getFolderById(folderId);
        var files = folder.getFiles();
        while (files.hasNext()) {
          var file = files.next();
          var name = file.getName();
          if (name.toLowerCase().indexOf("tech_img_") === 0) {
            var cleanKey = name.substring("tech_img_".length);
            var dotIdx = cleanKey.lastIndexOf(".");
            if (dotIdx !== -1) {
              cleanKey = cleanKey.substring(0, dotIdx);
            }
            var parts = cleanKey.split("_");
            if (parts.length > 1) {
              var lastPart = parts[parts.length - 1];
              if (!isNaN(lastPart) && lastPart.length >= 10) {
                parts.pop();
                cleanKey = parts.join("_");
              }
            }
            imagesMap[cleanKey] = "https://drive.google.com/file/d/" + file.getId() + "/preview";
          }
        }
      } catch (driveErr) {
        Logger.log("Drive resim listeleme hatası: " + driveErr.toString());
      }
      response.images = imagesMap;
      response.message = Object.keys(imagesMap).length + " adet teçhizat resmi listelendi.";

    } else if (action === "getPdfBase64") {
      try {
        var fileId = e.parameter.fileId;
        var file = DriveApp.getFileById(fileId);
        var blob = file.getBlob();
        var bytes = blob.getBytes();
        var base64 = Utilities.base64Encode(bytes);
        response.base64 = base64;
        response.name = file.getName();
        response.message = "PDF başarıyla base64 formatına kodlandı.";
      } catch (pdfErr) {
        response.status = "error";
        response.message = "PDF kodlama hatası: " + pdfErr.toString();
      }

    } else if (action === "filterSheets") {
      var prefix = e.parameter.prefix;
      var currentSheets = ss.getSheets();
      var matchedCount = 0;
      var firstMatched = null;
      
      for (var s = 0; s < currentSheets.length; s++) {
        var shName = currentSheets[s].getName();
        if (shName === prefix || shName.indexOf(prefix + "-") === 0) {
          currentSheets[s].showSheet();
          if (!firstMatched) firstMatched = currentSheets[s];
          matchedCount++;
        }
      }
      
      if (firstMatched) {
        ss.setActiveSheet(firstMatched);
        firstMatched.activate();
      }
      
      for (var s = 0; s < currentSheets.length; s++) {
        var shName = currentSheets[s].getName();
        var isMatch = (shName === prefix || shName.indexOf(prefix + "-") === 0);
        if (!isMatch) {
          try { currentSheets[s].hideSheet(); } catch(err) {}
        }
      }
      
      response.message = prefix + " sayfaları başarıyla gösterildi. Diğer sayfalar kalabalığı önlemek için gizlendi.";
      response.sheets = ss.getSheets().map(function(s) {
        return { name: s.getName(), id: s.getSheetId() };
      });

    } else if (action === "readSheet") {
      var sheetName = e.parameter.sheetName;
      var targetSs = null;
      var targetSsId = e.parameter.spreadsheetId;
      if (targetSsId) {
        try {
          targetSs = SpreadsheetApp.openById(targetSsId);
        } catch (err) {
          Logger.log("openById failed: " + err.toString());
        }
      }
      if (!targetSs) {
        var fallbackId = "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ";
        try {
          targetSs = SpreadsheetApp.openById(fallbackId);
        } catch (err) {
          Logger.log("openById fallback failed: " + err.toString());
        }
      }
      if (!targetSs) {
        targetSs = ss;
      }
      var sheet = getSheetWithFallback(targetSs, sheetName);
      var rows = [];
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      
      if (sheetName === "TASKLINE-PARÇA LİSTESİ" || sheetName === "Sayfa1" || (sheetName && sheetName.toLowerCase() === "sayfa1") || (sheetName && sheetName.toLowerCase() === "taskline-parca listesi")) {
        if (lastRow > 0 && lastCol > 0) {
          var allValues = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 7)).getDisplayValues();
          var currentEbysNo = "";
          
          for (var i = 0; i < allValues.length; i++) {
            var row = allValues[i];
            var cellA = String(row[0] || "").trim();
            var cellB = String(row[1] || "").trim();
            var cellC = String(row[2] || "").trim();
            
            if (cellA.indexOf("EBYS NO:") === 0) {
              if (cellB) {
                currentEbysNo = cellB;
              } else {
                currentEbysNo = cellA.replace("EBYS NO:", "").trim();
              }
              continue;
            }
            
            if (cellA === "AİT OLDUĞU BİRİM" || cellA === "AIT OLDUGU BIRIM" || cellA === "SIRA NU." || cellA === "EBYS NO") {
              continue;
            }
            
            if (!cellA && !cellB && !cellC) {
              continue;
            }
            
            if (currentEbysNo) {
              rows.push({
                "EBYS NO": currentEbysNo,
                "ebysNo": currentEbysNo,
                "AİT OLDUĞU BİRİM": cellA,
                "aitOlduguBirim": cellA,
                "MALZEME ADI": cellB,
                "malzemeAdi": cellB,
                "PARÇA NUMARASI": cellC,
                "parcaNumarasi": cellC,
                "SERİ NO (S/N)": String(row[3] || "").trim(),
                "seriNo": String(row[3] || "").trim(),
                "MİKTAR / KAPASİTE": String(row[4] || "").trim(),
                "miktarKapasite": String(row[4] || "").trim(),
                "KONTROLÜ YAPAN FİRMA": String(row[5] || "").trim(),
                "kontroluYapanFirma": String(row[5] || "").trim(),
                "AÇIKLAMA": String(row[6] || "").trim(),
                "aciklama": String(row[6] || "").trim()
              });
            }
          }
        }
      } else if (lastRow > 0 && lastCol > 0) {
        var allValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
        var rawHeaders = allValues[0];
        var headers = [];
        var isGorevlendirme = (sheet.getName().toLowerCase().indexOf("1-gorevlendirme") !== -1);
        
        for (var j = 0; j < lastCol; j++) {
          if (isGorevlendirme) {
            headers.push(String.fromCharCode(65 + j));
          } else {
            var h = String(rawHeaders[j]).trim();
            headers.push(h === "" ? "Kolon_" + (j + 1) : h);
          }
        }
        
        var startIdx = isGorevlendirme ? 0 : 1;
        for (var i = startIdx; i < allValues.length; i++) {
          var rowObj = {};
          var hasValue = false;
          for (var j = 0; j < headers.length; j++) {
            var header = headers[j];
            var val = allValues[i][j];
            if (val instanceof Date) {
              if (sheet.getName().toLowerCase().indexOf("tarih") !== -1) {
                rowObj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
              } else {
                var hours = val.getHours();
                var minutes = val.getMinutes();
                var seconds = val.getSeconds();
                if (hours !== 0 || minutes !== 0 || seconds !== 0) {
                  rowObj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
                } else {
                  rowObj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
                }
              }
            } else {
              rowObj[header] = (val !== undefined && val !== null) ? String(val) : "";
            }
            if (rowObj[header].trim() !== "") hasValue = true;
          }
          if (hasValue) rows.push(rowObj);
        }
      }
      
      response.data = rows;
      response.sheetName = sheet.getName();
      response.rowCount = rows.length;
      response.message = sheet.getName() + " verileri okundu.";

    } else {
      response.message = "Hava Araçları Senkronizasyon Bağlantısı Aktif.";
      response.sheetCount = ss.getNumSheets();
      response.sheets = ss.getSheets().map(function(s) {
        return { name: s.getName(), id: s.getSheetId() };
      });
    }
  } catch (err) {
    response.status = "error";
    response.message = err.toString();
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Güncelleme tarihlerini 'güncelleme tarihleri' isimli özel sayfaya kaydeder ve biçimlendirir.
 */
function recordLastUpdate(ss, unitName, formattedDate) {
  var sheet = getSheetWithFallback(ss, "güncelleme tarihleri");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["birim adı", "güncelleme tarih saat"]);
    sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
    sheet.setColumnWidth(1, 320);
    sheet.setColumnWidth(2, 220);
  }
  
  var lastRow = sheet.getLastRow();
  var entireRange = sheet.getRange(1, 1, lastRow, 2);
  var values = entireRange.getValues();
  var updated = false;
  
  var targetUnitNameClean = unitName.trim().toLowerCase();
  
  // Satır bazlı arama yapalım (Başlık satırından sonrakiler)
  for (var r = 1; r < values.length; r++) {
    var rowUnitName = String(values[r][0]).trim();
    var rowUnitNameClean = rowUnitName.toLowerCase();
    
    // Satırda "Birim" yazıyorsa veya tam eşleşme varsa veya benzer isimli ise üzerine yaz
    if (rowUnitNameClean === targetUnitNameClean || 
        rowUnitNameClean === "birim" ||
        (targetUnitNameClean.indexOf("bell") !== -1 && rowUnitNameClean.indexOf("bell") !== -1 && targetUnitNameClean.indexOf("yaz") !== -1 && rowUnitNameClean.indexOf("yaz") !== -1 && getMonthFromTitle(targetUnitNameClean) === getMonthFromTitle(rowUnitNameClean)) ||
        (targetUnitNameClean.indexOf("t-70") !== -1 && rowUnitNameClean.indexOf("t-70") !== -1 && targetUnitNameClean.indexOf("yaz") !== -1 && rowUnitNameClean.indexOf("yaz") !== -1 && getMonthFromTitle(targetUnitNameClean) === getMonthFromTitle(rowUnitNameClean)) ||
        (targetUnitNameClean.indexOf("at-802") !== -1 && rowUnitNameClean.indexOf("at-802") !== -1 && targetUnitNameClean.indexOf("yaz") !== -1 && rowUnitNameClean.indexOf("yaz") !== -1 && getMonthFromTitle(targetUnitNameClean) === getMonthFromTitle(rowUnitNameClean))) {
      
      sheet.getRange(r + 1, 1).setValue(unitName); // "Birim" ise ismini de düzeltir
      sheet.getRange(r + 1, 2).setValue(formattedDate);
      updated = true;
      break;
    }
  }
  
  if (!updated) {
    sheet.appendRow([unitName, formattedDate]);
  }
  
  // Tablo çizgisini ve hizalamasını mükemmelleştirme
  var newLastRow = sheet.getLastRow();
  var formattedRange = sheet.getRange(1, 1, newLastRow, 2);
  formattedRange.setFontFamily("Calibri").setFontSize(11).setVerticalAlignment("middle");
  
  // İlk satır başlıklarını küçük harf kalsın diye dokunmuyoruz, sadece biçim veriyoruz
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
  sheet.getRange(2, 1, newLastRow - 1, 2).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
}

function getMonthFromTitle(title) {
  if (title.indexOf("(") !== -1) {
    return title.split("(")[1].replace(")", "").trim();
  }
  return "";
}

/**
 * Sayfayı adına göre bulur. Bulamazsa temizce oluşturur.
 */
function getSheetWithFallback(ss, name) {
  var ssId = ss.getId();
  var ssName = ss.getName().toUpperCase();
  var isTaskline = (ssId === "1L05588TdYZmH401Lvn4_yr4zwiw2pW4EJ8dIyl-UTVQ" || ssName.indexOf("TASKLINE") !== -1);

  if (isTaskline) {
    var forbiddenWords = ["personel", "görev", "gorev", "tarih", "tüm", "tum", "gün", "gun"];
    var lowerName = name.toLowerCase();
    var isForbidden = forbiddenWords.some(function(word) {
      return lowerName.indexOf(word) !== -1;
    });
    if (isForbidden || lowerName !== "sayfa1") {
      return ss.getSheetByName("Sayfa1") || ss.getSheets()[0];
    }
  }

  var sheets = ss.getSheets();
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  
  var sanitize = function(str) {
    return str.toLowerCase()
      .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u')
      .replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c')
      .replace(/[^a-z0-9]/g, '');
  };
  
  var targetSanitized = sanitize(name);
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName();
    if (sanitize(sName) === targetSanitized) return sheets[i];
  }
  
  // Özel eşleşme köprüleri
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().toLowerCase();
    var target = name.toLowerCase();
    if (sName.indexOf("bell_429") !== -1 && target.indexOf("bell_429") !== -1) return sheets[i];
    if (sName.indexOf("t_70") !== -1 && target.indexOf("t_70") !== -1) return sheets[i];
    if (sName.indexOf("at_802") !== -1 && target.indexOf("at_802") !== -1) return sheets[i];
    if (sName.indexOf("bakim") !== -1 && target.indexOf("bakim") !== -1) return sheets[i];
    if (sName.indexOf("1-") === 0 && target.indexOf("1-") === 0) return sheets[i];
  }
  
  return ss.insertSheet(name);
}

/**
 * Tablodaki gerçek son satırı (ilk 6 sütunda herhangi bir veri olan son satır) döndürür.
 * sheet.getLastRow() bazen boş/biçimlendirilmiş satırları da sayabildiği için bu yöntemi kullanıyoruz.
 */
function getActualLastRow(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) return 0;
  
  var checkRows = Math.min(lastRow, 2000);
  var values = sheet.getRange(1, 1, checkRows, 6).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    for (var j = 0; j < 6; j++) {
      if (String(values[i][j] || "").trim() !== "") {
        return i + 1;
      }
    }
  }
  return 0;
}

function getUnitTitleByPrefix(prefix) {
  var pl = prefix.toLowerCase();
  if (pl.indexOf("1-gorevlendirme") !== -1) return "1. GÖREVLENDİRME ÇİZELGELERİ";
  if (pl.indexOf("3-bakim_yetki") !== -1) return "3. BAKIM YETKİ ÇİZELGELERİ";
  if (pl.indexOf("5-personel_bilgi") !== -1) return "5. PERSONEL BİLGİ ÇİZELGELERİ";
  if (pl.indexOf("6-personel_ucus_hizmet") !== -1) return "6. PERSONEL UÇUŞ-HİZMET YILLARI";
  return prefix;
}

function getUnitTitleById(id, month) {
  var title = "Birim";
  if (id === 1 || id === "1") title = "1. GÖREVLENDİRME ÇİZELGELERİ";
  else if (id === 21 || id === "21") title = "2. YAZ DÖNEMİ PLANLAMASI - BELL 429";
  else if (id === 22 || id === "22") title = "2. YAZ DÖNEMİ PLANLAMASI - T-70";
  else if (id === 23 || id === "23") title = "2. YAZ DÖNEMİ PLANLAMASI - AT-802";
  else if (id === 3 || id === "3") title = "3. BAKIM YETKİ ÇİZELGELERİ";
  else if (id === 5 || id === "5") title = "5. PERSONEL BİLGİ ÇİZELGELERİ";
  else if (id === 6 || id === "6") title = "6. PERSONEL UÇUŞ-HİZMET YILLARI";
  
  if (month && month !== "Genel Plan") {
    title += " (" + month + ")";
  }
  return title;
}

/**
 * Excel formatında indirme modal penceresini açar.
 */
function exportAndDownloadExcel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var downloadUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx";
  
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', Tahoma, Geneva, Verdana, sans-serif; padding: 25px; color: #1e293b; background-color: #f8fafc; border-radius: 16px; border: 1px solid #e2e8f0; text-align: center;">' +
      '<div style="display: inline-block; width: 48px; height: 48px; background-color: #f0fdf4; border-radius: 12px; margin-bottom: 15px; color: #166534; line-height: 48px; font-size: 24px;">🚁</div>' +
      '<h3 style="color: #0b3d1d; margin-top: 0; font-weight: 800; font-size: 15px; text-transform: uppercase; tracking: 0.5px;">EXCEL DOSYASI HAZIRLANDI</h3>' +
      '<p style="font-size: 11px; color: #64748b; line-height: 1.5; margin: 10px 0 25px 0; font-weight: 500;">Canlı entegre e-tablo verileriniz indirilebilir durumdadır. Çevrimdışı HA Portalınız üzerinde hemen kullanabilirsiniz.</p>' +
      '<div style="margin-top: 20px;">' +
        '<a href="' + downloadUrl + '" target="_blank" style="background-color: #0b3d1d; hover:background-color: #166534; color: white; padding: 12px 28px; text-decoration: none; font-size: 12px; font-weight: 800; border-radius: 8px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);" onclick="google.script.host.close()">📥 EXCEL DOSYASINI İNDİR</a>' +
      '</div>' +
    '</div>'
  ).setWidth(420).setHeight(220).setTitle('HA Bakım - Excel İndirme Sihirbazı');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Excel Dosyasını Paketle');
}

/**
 * Entegrasyon kılavuzu modal penceresini açar.
 */
function showIntegrationGuide() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: \'Segoe UI\', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; color: #1e293b;">' +
      '<h2 style="color: #0b3d1d; font-size: 15px; font-weight: 800; margin-bottom: 12px; border-bottom: 2px solid #0b3d1d; padding-bottom: 8px; uppercase">💡 ENTEGRASYON BİLGİSİ</h2>' +
      '<p style="font-size: 12px; line-height: 1.6; color: #475569; font-weight: 500;">' +
        'Bu senkronizasyon motoru aracılığıyla, Hava Araçları Bakım Bilgi Portali üzerinden yapılan tüm e-tablo anlık güncellemeleri, ' +
        'canlı olarak bu tabloya işlenmekte ve güncellenme tarihleri otomatik kaydedilmektedir.' +
      '</p>' +
      '<div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-top: 15px;">' +
        '<strong style="font-size: 11px; color: #0b3d1d; display:block; margin-bottom: 4px;">📌 PDF VE KLASÖR KURALLARI:</strong>' +
        '<span style="font-size: 10px; color: #64748b; line-height: 1.4; display:block;">' +
          '• Yaz dönemine ait PDF planlamaları otomatik olarak ortak Google Drive klasöründen okunur.<br/>' +
          '• PDF dosyalarınızı her ay için belirlenen formatta Drive klasörüne yüklemeniz yeterlidir.' +
        '</span>' +
      '</div>' +
      '<div style="text-align: center; margin-top: 25px;">' +
        '<button onclick="google.script.host.close()" style="background-color: #64748b; color: white; border: none; padding: 8px 18px; border-radius: 6px; font-size: 11px; font-weight: bold; cursor: pointer; transition: background 0.2s;">Anlaşıldı, Kapat</button>' +
      '</div>' +
    '</div>'
  ).setWidth(420).setHeight(260).setTitle('HA Bakım Entegrasyon Bilgi Kılavuzu');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Entegrasyon Kılavuzu');
}

/**
 * Görevlendirme çizelgelerini özel tasarımla biçimlendirir.
 */
function formatGorevlendirmeSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return;
  
  var cols = Math.max(lastCol, 5);
  var range = sheet.getRange(1, 1, lastRow, cols);
  try { range.breakApart(); } catch(e) {}
  
  range.setFontFamily("Calibri").setFontSize(11).setVerticalAlignment("middle").setHorizontalAlignment("left");
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 220);
  
  range.setBorder(true, true, true, true, true, true, "#3b82f6", SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Genel sayfaları kurumsal renkler (Yeşil/Beyaz) ve otomatik sütun genişliği ile biçimlendirir.
 */
function formatGeneralSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return;
  
  var range = sheet.getRange(1, 1, lastRow, lastCol);
  range.setFontFamily("Calibri").setFontSize(11).setVerticalAlignment("middle");
  
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff").setHorizontalAlignment("center");
  range.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  
  for (var c = 1; c <= lastCol; c++) {
    sheet.autoResizeColumn(c);
    var width = sheet.getColumnWidth(c);
    if (width < 120) sheet.setColumnWidth(c, 120);
    else if (width > 320) sheet.setColumnWidth(c, 320);
  }
}

/**
 * Teçhizat gelecek kontrol gün takibini inceler ve 90, 60, 30 gün kala otomatik toplu mail atar.
 */
function processEquipmentReminders(ss, formattedDate) {
  try {
    var techSheet = getSheetWithFallback(ss, "TÜM TECHİZAT");
    var gunTakipSheet = ensureGunTakipSheet(ss);
    
    var lastRowTech = techSheet.getLastRow();
    if (lastRowTech <= 1) return;
    
    // Read GÜN TAKİP contacts
    var gunTakipLastRow = gunTakipSheet.getLastRow();
    var gunTakipData = gunTakipSheet.getRange(1, 1, gunTakipLastRow, 6).getValues();
    var emailMap = {}; // key: birimAdı (BÜYÜK HARF) -> { rowIndex: number, originalUnitName: string, unitKey: string, name: string, email: string }
    
    for (var i = 1; i < gunTakipData.length; i++) {
      var unit = String(gunTakipData[i][0]).trim();
      var unitUpper = unit.toUpperCase();
      var name = String(gunTakipData[i][1]).trim();
      var email = String(gunTakipData[i][2]).trim();
      if (unit) {
        emailMap[unitUpper] = { 
          rowIndex: i + 1, 
          originalUnitName: unit,
          unitKey: unitUpper,
          name: name, 
          email: email 
        };
      }
    }
    
    // Ensure all 18 columns exist in TÜM TECHİZAT
    var lastColTech = techSheet.getLastColumn();
    if (lastColTech < 18) {
      techSheet.getRange(1, 16).setValue("90 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      techSheet.getRange(1, 17).setValue("60 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      techSheet.getRange(1, 18).setValue("30 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      lastColTech = 18;
    }
    
    var techRange = techSheet.getRange(2, 1, lastRowTech - 1, 18);
    var techValues = techRange.getValues();
    var modified = false;
    
    var pendingAlerts = {};
    
    for (var r = 0; r < techValues.length; r++) {
      var row = techValues[r];
      var unitName = String(row[0]).trim().toUpperCase();
      var siraNo = String(row[1]).trim();
      var isKara = (unitName === "KARA ARAÇLARI" || unitName.indexOf("KARA") !== -1);
      
      var techName = isKara ? String(row[3]).trim() : String(row[2]).trim();
      if (techName === "-") techName = "";
      
      var partNo = isKara ? String(row[5]).trim() : String(row[4]).trim();
      if (partNo === "-") partNo = "";
      
      var seriNo = isKara ? "" : String(row[6]).trim();
      if (seriNo === "-") seriNo = "";
      
      var location = String(row[9]).trim();
      var status = String(row[10]).trim();
      var gelecekBakimStr = String(row[12]).trim();
      var sonKontrolYapan = String(row[13]).trim();
      var aciklama = String(row[14]).trim();
      
      var mail90 = String(row[15] || "").trim();
      var mail60 = String(row[16] || "").trim();
      var mail30 = String(row[17] || "").trim();
      
      if (!gelecekBakimStr) continue;
      
      var daysDiff = getDaysDiff(gelecekBakimStr);
      if (daysDiff === null) continue;
      
      // Reset if future control date is updated (extended)
      if (daysDiff > 90) {
        if (mail90 || mail60 || mail30) {
          techValues[r][15] = "";
          techValues[r][16] = "";
          techValues[r][17] = "";
          modified = true;
        }
        continue;
      }
      
      var warningLevel = "";
      var colToUpdate = -1; // index in row (15 for 90 days, 16 for 60 days, 17 for 30 days)
      
      if (daysDiff <= 30) {
        if (!mail30) {
          warningLevel = "30 GÜN UYARISI";
          colToUpdate = 17;
        }
      } else if (daysDiff <= 60) {
        if (!mail60) {
          warningLevel = "60 GÜN UYARISI";
          colToUpdate = 16;
        }
      } else if (daysDiff <= 90) {
        if (!mail90) {
          warningLevel = "90 GÜN UYARISI";
          colToUpdate = 15;
        }
      }
      
      if (warningLevel !== "" && colToUpdate !== -1) {
        var contact = null;
        var normUnitName = normalizeUnitName(unitName);
        for (var uKey in emailMap) {
          var normKey = normalizeUnitName(uKey);
          if (normUnitName.indexOf(normKey) !== -1 || normKey.indexOf(normUnitName) !== -1) {
            contact = emailMap[uKey];
            break;
          }
        }
        
        var targetUnitKey = contact ? contact.unitKey : "VARSAYILAN";
        var contactName = contact ? contact.name : "Sorumlu Personel";
        var contactEmail = contact ? contact.email : "romanhavacilik.bakimsube@gmail.com";
        var gunTakipRowIndex = contact ? contact.rowIndex : -1;
        
        var groupKey = targetUnitKey + "||" + warningLevel;
        if (!pendingAlerts[groupKey]) {
          pendingAlerts[groupKey] = {
            contactName: contactName,
            contactEmail: contactEmail,
            gunTakipRowIndex: gunTakipRowIndex,
            warningLevel: warningLevel,
            unitName: contact ? contact.originalUnitName : unitName,
            colToUpdate: colToUpdate,
            triggeringIndexes: [],
            items: []
          };
        }
        
        pendingAlerts[groupKey].triggeringIndexes.push(r);
      }
    }
    
    // Grouping all equipment items under 90 days
    for (var gKey in pendingAlerts) {
      var group = pendingAlerts[gKey];
      var groupUnitNorm = normalizeUnitName(group.unitName);
      
      for (var r = 0; r < techValues.length; r++) {
        var rRow = techValues[r];
        var rUnitName = String(rRow[0]).trim().toUpperCase();
        var rUnitNorm = normalizeUnitName(rUnitName);
        
        if (groupUnitNorm.indexOf(rUnitNorm) !== -1 || rUnitNorm.indexOf(groupUnitNorm) !== -1) {
          var rGelecekBakimStr = String(rRow[12]).trim();
          if (!rGelecekBakimStr) continue;
          
          var rDaysDiff = getDaysDiff(rGelecekBakimStr);
          if (rDaysDiff !== null && rDaysDiff <= 90) {
            var rIsKara = (rUnitName === "KARA ARAÇLARI" || rUnitName.indexOf("KARA") !== -1);
            var rTechName = rIsKara ? String(rRow[3]).trim() : String(rRow[2]).trim();
            if (rTechName === "-") rTechName = "";
            
            var rPartNo = rIsKara ? String(rRow[5]).trim() : String(rRow[4]).trim();
            if (rPartNo === "-") rPartNo = "";
            
            var rSeriNo = rIsKara ? "" : String(rRow[6]).trim();
            if (rSeriNo === "-") rSeriNo = "";
            
            group.items.push({
              techValuesIndex: r,
              techName: rTechName,
              partNo: rPartNo,
              seriNo: rSeriNo,
              location: String(rRow[9]).trim(),
              status: String(rRow[10]).trim(),
              gelecekBakimStr: rGelecekBakimStr,
              sonKontrolYapan: String(rRow[13]).trim(),
              aciklama: String(rRow[14]).trim(),
              daysDiff: rDaysDiff
            });
          }
        }
      }
      
      group.items.sort(function(a, b) {
        return a.daysDiff - b.daysDiff;
      });
    }
    
    // Send email reminders
    for (var gKey in pendingAlerts) {
      var group = pendingAlerts[gKey];
      if (group.items.length > 0) {
        sendConsolidatedReminderEmail(group.contactEmail, group.contactName, group.warningLevel, group.unitName, group.items);
        
        for (var k = 0; k < group.triggeringIndexes.length; k++) {
          var rIdx = group.triggeringIndexes[k];
          techValues[rIdx][group.colToUpdate] = formattedDate;
          
          // Cascading updates
          if (group.colToUpdate === 17) { // 30 GÜN
            if (!techValues[rIdx][16]) {
              techValues[rIdx][16] = formattedDate;
            }
            if (!techValues[rIdx][15]) {
              techValues[rIdx][15] = formattedDate;
            }
          } else if (group.colToUpdate === 16) { // 60 GÜN
            if (!techValues[rIdx][15]) {
              techValues[rIdx][15] = formattedDate;
            }
          }
        }
        modified = true;
        
        // Update contact warning log on GÜN TAKİP
        if (group.gunTakipRowIndex !== -1) {
          var gtCol = -1;
          if (group.warningLevel === "90 GÜN UYARISI") gtCol = 4;
          else if (group.warningLevel === "60 GÜN UYARISI") gtCol = 5;
          else if (group.warningLevel === "30 GÜN UYARISI") gtCol = 6;
          
          if (gtCol !== -1) {
            gunTakipSheet.getRange(group.gunTakipRowIndex, gtCol).setValue(formattedDate);
          }
        }
      }
    }
    
    if (modified) {
      techRange.setValues(techValues);
      formatGeneralSheet(techSheet);
      formatGeneralSheet(gunTakipSheet);
    }
  } catch (err) {
    Logger.log("processEquipmentReminders hatası: " + err.toString());
  }
}

/**
 * Gelecek bakım tarihi ile bugün arasındaki gün farkını hesaplar.
 */
function getDaysDiff(dateStr) {
  if (!dateStr) return null;
  
  // Eğer zaten bir Date objesiyse doğrudan kullanalım
  if (dateStr instanceof Date) {
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var target = new Date(dateStr.getTime());
    target.setHours(0, 0, 0, 0);
    return Math.ceil((target.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  }
  
  var cleaned = String(dateStr).trim();
  var dmyRegex = /^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})$/;
  var ymdRegex = /^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})$/;
  
  var dateObj = null;
  var m = cleaned.match(dmyRegex);
  if (m) {
    var day = parseInt(m[1], 10);
    var month = parseInt(m[2], 10) - 1;
    var year = parseInt(m[3], 10);
    dateObj = new Date(year, month, day);
  } else {
    m = cleaned.match(ymdRegex);
    if (m) {
      var year = parseInt(m[1], 10);
      var month = parseInt(m[2], 10) - 1;
      var day = parseInt(m[3], 10);
      dateObj = new Date(year, month, day);
    } else {
      var timestamp = Date.parse(cleaned);
      if (!isNaN(timestamp)) {
        dateObj = new Date(timestamp);
      }
    }
  }
  
  if (!dateObj || isNaN(dateObj.getTime())) return null;
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  dateObj.setHours(0, 0, 0, 0);
  
  var diffTime = dateObj.getTime() - today.getTime();
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

/**
 * Otomatik şablonlu toplu hatırlatma e-postası gönderir.
 */
function sendConsolidatedReminderEmail(email, personName, warningLevel, unitName, items) {
  var subject = "🚨 YER DESTEK TEÇHİZATLARI YAKLAŞAN DURUMLAR EKTEDİR - [" + unitName + "] (" + warningLevel + ")";
  
  var colorHex = "#16a34a"; // Yeşil (90)
  if (warningLevel.indexOf("60") !== -1) colorHex = "#ea580c"; // Turuncu (60)
  else if (warningLevel.indexOf("30") !== -1) colorHex = "#dc2626"; // Kırmızı (30)
  
  var tableRowsHtml = "";
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    tableRowsHtml += 
      "<tr style='border-bottom: 1px solid #e2e8f0;'>" +
        "<td style='padding: 10px 8px; font-weight: bold; font-size: 11px; text-align: center; color: #64748b;'>" + (i + 1) + "</td>" +
        "<td style='padding: 10px 8px; font-weight: 700; font-size: 12px; color: #1e293b;'>" + item.techName + "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; color: #475569;'>" + (item.partNo || "-") + "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; font-family: monospace; color: #475569;'>" + (item.seriNo || "-") + "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; color: #475569;'>" + (item.location || "-") + "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; font-weight: 700; color: " + colorHex + "; text-align: center;'>" + item.gelecekBakimStr + "<br><span style='font-size: 10px; font-weight: 500;'>(" + item.daysDiff + " gün kaldı)</span></td>" +
        "<td style='padding: 10px 8px; font-size: 11px; color: #64748b;'>" + (item.aciklama || "-") + "</td>" +
      "</tr>";
  }
  
  var body = 
    "<html>" +
    "<head>" +
      "<style>" +
        "body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; color: #1e293b; padding: 20px; }" +
        ".container { max-width: 850px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; border: 1px solid #e2e8f0; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }" +
        ".header { background-color: " + colorHex + "; padding: 30px 20px; text-align: center; color: #ffffff; }" +
        ".header h2 { margin: 0; font-size: 22px; font-weight: 800; letter-spacing: -0.5px; text-transform: uppercase; }" +
        ".header p { margin: 5px 0 0 0; font-size: 13px; font-weight: 500; opacity: 0.9; }" +
        ".content { padding: 30px 24px; }" +
        ".intro { font-size: 14px; font-weight: 500; color: #475569; margin-bottom: 25px; }" +
        ".items-table { width: 100%; border-collapse: collapse; margin-bottom: 25px; text-align: left; }" +
        ".items-table th { background-color: #f1f5f9; padding: 12px 8px; font-size: 11px; font-weight: bold; color: #475569; text-transform: uppercase; border-bottom: 2px solid #cbd5e1; }" +
        ".footer { background-color: #f8fafc; padding: 20px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; }" +
      "</style>" +
    "</head>" +
    "<body>" +
      "<div class='container'>" +
        "<div class='header'>" +
          "<h2>🚁 OGM HAVACILIK BAKIM VE TEKNİK ŞUBE MÜDÜRLÜĞÜ</h2>" +
          "<p>" + warningLevel + " - TOPLU BİLGİLENDİRME</p>" +
        "</div>" +
        "<div class='content'>" +
          "<div class='intro'>Sayın <strong>" + personName + "</strong>,<br><br>" +
          "<strong>YER DESTEK TEÇHİZATLARI YAKLAŞAN DURUMLAR EKTEDİR.</strong><br><br>" +
          "<strong>" + unitName + "</strong> biriminize ait aşağıda listelenen yer destek teçhizatlarının <strong>GELECEK KONTROL / BAKIM</strong> tarihlerine <strong>" + (warningLevel.indexOf("30") !== -1 ? "30 günden az" : warningLevel.indexOf("60") !== -1 ? "60 günden az" : "90 günden az") + "</strong> süre kalmıştır. Gerekli kontrollerin ve bakımların zamanında yapılması kritik önem taşımaktadır.</div>" +
          
          "<table class='items-table'>" +
            "<thead>" +
              "<tr>" +
                "<th style='width: 5%; text-align: center;'>SIRA</th>" +
                "<th style='width: 25%;'>TEÇHİZAT ADI</th>" +
                "<th style='width: 15%;'>PARÇA NO / MODEL</th>" +
                "<th style='width: 15%;'>SERİ NO (S/N)</th>" +
                "<th style='width: 15%;'>BULUNDUĞU YER</th>" +
                "<th style='width: 13%; text-align: center;'>KONTROL TARİHİ</th>" +
                "<th style='width: 12%;'>AÇIKLAMA</th>" +
              "</tr>" +
            "</thead>" +
            "<tbody>" +
              tableRowsHtml +
            "</tbody>" +
          "</table>" +
          
          "<div style='background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 12px; padding: 15px; text-align: center; font-size: 11px; font-weight: bold; color: #991b1b;'>" +
            "⚠️ Lütfen listedeki teçhizatların bakımlarını tamamladıktan sonra portal üzerinden Excel belgesini güncelleyerek e-tabloyu yenileyiniz." +
          "</div>" +
        "</div>" +
        "<div class='footer'>" +
          "T.C. ORMAN GENEL MÜDÜRLÜĞÜ - HAVA ARAÇLARI BAKIM VE TEKNİK ŞUBE MÜDÜRLÜĞÜ<br>" +
          "Bu e-posta otomatik olarak üretilmiştir, lütfen yanıtlamayınız." +
          "<br><span style='font-size: 9px;'>Gönderim Tarihi: " + Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy HH:mm:ss") + "</span>" +
        "</div>" +
      "</div>" +
    "</body>" +
    "</html>";
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body
    });
  } catch (err) {
    Logger.log("Toplu Mail gönderim hatası (" + email + "): " + err.toString());
  }
}

/**
 * Otomatik şablonlu hatırlatma e-postası gönderir.
 */
function sendReminderEmail(email, personName, warningLevel, unitName, techName, partNo, seriNo, gelecekBakim, daysLeft, location, status, sonKontrolYapan, aciklama) {
  var subject = "🚨 " + warningLevel + " - Teçhizat Bakım Hatırlatması: [" + unitName + "] " + techName;
  
  var colorHex = "#16a34a"; // Yeşil (90)
  if (warningLevel.indexOf("60") !== -1) colorHex = "#ea580c"; // Turuncu (60)
  else if (warningLevel.indexOf("30") !== -1) colorHex = "#dc2626"; // Kırmızı (30)
  
  var body = 
    "<html>" +
    "<head>" +
      "<style>" +
        "body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; color: #1e293b; padding: 20px; }" +
        ".container { max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; border: 1px solid #e2e8f0; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }" +
        ".header { background-color: " + colorHex + "; padding: 30px 20px; text-align: center; color: #ffffff; }" +
        ".header h2 { margin: 0; font-size: 22px; font-weight: 800; letter-spacing: -0.5px; text-transform: uppercase; }" +
        ".header p { margin: 5px 0 0 0; font-size: 13px; font-weight: 500; opacity: 0.9; }" +
        ".content { padding: 30px 24px; }" +
        ".intro { font-size: 14px; font-weight: 500; color: #475569; margin-bottom: 25px; }" +
        ".detail-table { width: 100%; border-collapse: collapse; margin-bottom: 25px; }" +
        ".detail-table td { padding: 12px 10px; border-bottom: 1px solid #f1f5f9; font-size: 13px; }" +
        ".detail-table td.label { font-weight: bold; color: #64748b; width: 35%; text-transform: uppercase; font-size: 11px; }" +
        ".detail-table td.value { font-weight: 700; color: #1e293b; }" +
        ".footer { background-color: #f8fafc; padding: 20px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; }" +
      "</style>" +
    "</head>" +
    "<body>" +
      "<div class='container'>" +
        "<div class='header'>" +
          "<h2>🚁 OGM HAVACILIK BAKIM TAKİP SİSTEMİ</h2>" +
          "<p>" + warningLevel + " - OTOMATİK BİLGİLENDİRME</p>" +
        "</div>" +
        "<div class='content'>" +
          "<div class='intro'>Sayın <strong>" + personName + "</strong>,<br><br>" +
          "Aşağıda detayları belirtilen yer destek teçhizatının <strong>GELECEK KONTROL / BAKIM</strong> tarihine <strong>" + daysLeft + " gün</strong> kalmıştır. Gerekli kontrollerin ve bakımların zamanında yapılması kritik önem taşımaktadır.</div>" +
          
          "<table class='detail-table'>" +
            "<tr>" +
              "<td class='label'>AİT OLDUĞU BİRİM</td>" +
              "<td class='value'>" + unitName + "</td>" +
            "</tr>" +
            "<tr>" +
              "<td class='label'>TEÇHİZAT ADI</td>" +
              "<td class='value' style='color: " + colorHex + "'>" + techName + "</td>" +
            "</tr>" +
            "<tr>" +
              "<td class='label'>PARÇA NO (P/N) / MODEL</td>" +
              "<td class='value'>" + (partNo || "-") + "</td>" +
            "</tr>" +
            "<tr>" +
              "<td class='label'>SERİ NO (S/N)</td>" +
              "<td class='value'>" + (seriNo || "-") + "</td>" +
            "</tr>" +
            "<tr>" +
              "<td class='label'>BULUNDUĞU YER</td>" +
              "<td class='value'>" + (location || "-") + "</td>" +
            "</tr>" +
            "<tr>" +
              "<td class='label'>DURUMU</td>" +
              "<td class='value'>" + (status || "-") + "</td>" +
            "</tr>" +
            "<tr>" +
              "<td class='label'>GELECEK KONTROL / BAKIM</td>" +
              "<td class='value' style='color: " + colorHex + "; font-size: 14px;'>" + gelecekBakim + " (" + daysLeft + " gün kaldı)</td>" +
            "</tr>" +
            "<tr>" +
              "<td class='label'>SON KONTROLÜ YAPAN FİRMA</td>" +
              "<td class='value'>" + (sonKontrolYapan || "-") + "</td>" +
            "</tr>" +
            "<tr>" +
              "<td class='label'>AÇIKLAMA</td>" +
              "<td class='value'>" + (aciklama || "-") + "</td>" +
            "</tr>" +
          "</table>" +
          
          "<div style='background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 12px; padding: 15px; text-align: center; font-size: 11px; font-weight: bold; color: #991b1b;'>" +
            "⚠️ Lütfen teçhizat bakımını tamamladıktan sonra portal üzerinden Excel belgesini güncelleyerek e-tabloyu yenileyiniz." +
          "</div>" +
        "</div>" +
        "<div class='footer'>" +
          "T.C. ORMAN GENEL MÜDÜRLÜĞÜ - HAVA ARAÇLARI BAKIM VE TEKNİK ŞUBE MÜDÜRLÜĞÜ<br>" +
          "Bu e-posta otomatik olarak üretilmiştir, lütfen yanıtlamayınız." +
        "</div>" +
      "</div>" +
    "</body>" +
    "</html>";
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body
    });
  } catch (err) {
    Logger.log("Mail gönderim hatası (" + email + "): " + err.toString());
  }
}

/**
 * Gün Takip e-tablosunun varlığından ve varsayılan verilerinden emin olur.
 */
function ensureGunTakipSheet(ss) {
  var sheet = getSheetWithFallback(ss, "GÜN TAKİP");
  if (sheet.getLastRow() === 0) {
    var headers = [
      "SORUMLU BİRİM", 
      "ADI SOYADI", 
      "E-POSTA ADRESİ", 
      "90 GÜN UYARISI MAİL GÖNDERİM TARİHİ", 
      "60 GÜN UYARISI MAİL GÖNDERİM TARİHİ", 
      "30 GÜN UYARISI MAİL GÖNDERİM TARİHİ"
    ];
    sheet.appendRow(headers);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff").setHorizontalAlignment("center");
    
    var defaults = [
      ["BELL 429", "Sorumlu Personel", "ormanhavacilik.bakimsube@gmail.com", "", "", ""],
      ["AT-802F", "Sorumlu Personel", "ormanhavacilik.bakimsube@gmail.com", "", "", ""],
      ["T-70 YER DESTEK", "Sorumlu Personel", "ormanhavacilik.bakimsube@gmail.com", "", "", ""],
      ["T-70 BUMBİ BACKET", "Sorumlu Personel", "ormanhavacilik.bakimsube@gmail.com", "", "", ""],
      ["B-360", "Sorumlu Personel", "ormanhavacilik.bakimsube@gmail.com", "", "", ""],
      ["C-650", "Sorumlu Personel", "ormanhavacilik.bakimsube@gmail.com", "", "", ""],
      ["HANGAR YER DESTEK", "Sorumlu Personel", "ormanhavacilik.bakimsube@gmail.com", "", "", ""],
      ["KARA ARAÇLARI", "Sorumlu Personel", "ormanhavacilik.bakimsube@gmail.com", "", "", ""]
    ];
    sheet.getRange(2, 1, defaults.length, 6).setValues(defaults);
    formatGeneralSheet(sheet);
  } else {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 6) {
      sheet.getRange(1, 4).setValue("90 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      sheet.getRange(1, 5).setValue("60 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      sheet.getRange(1, 6).setValue("30 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      formatGeneralSheet(sheet);
    }
  }
  return sheet;
}

/**
 * Türkçe karakterleri normalize eder ve boşlukları/özel karakterleri silerek eşleşmeleri sağlamlaştırır.
 */
function normalizeUnitName(str) {
  if (!str) return "";
  return String(str).trim()
    .replace(/İ/g, "I")
    .replace(/ı/g, "i")
    .replace(/Ğ/g, "G")
    .replace(/ğ/g, "g")
    .replace(/Ü/g, "U")
    .replace(/ü/g, "u")
    .replace(/Ş/g, "S")
    .replace(/ş/g, "s")
    .replace(/Ö/g, "O")
    .replace(/ö/g, "o")
    .replace(/Ç/g, "C")
    .replace(/ç/g, "c")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

/**
 * Updates the company name in the matching "TÜM TECHİZAT" sheet rows
 */
function updateTechizatFirmaInAllTechizat(ss, partNo, seriNo, newFirma) {
  var techSheet = getSheetWithFallback(ss, "TÜM TECHİZAT");
  var lastRow = techSheet.getLastRow();
  if (lastRow > 1) {
    var range = techSheet.getRange(2, 1, lastRow - 1, 15);
    var values = range.getValues();
    for (var r = 0; r < values.length; r++) {
      var unitName = String(values[r][0] || "").trim().toUpperCase();
      var isKara = (unitName === "KARA ARAÇLARI" || unitName.indexOf("KARA") !== -1);
      
      var rowPart = isKara ? String(values[r][5] || "").trim() : String(values[r][4] || "").trim();
      var rowSeri = isKara ? String(values[r][3] || "").trim() : String(values[r][6] || "").trim();
      
      if ((rowPart === partNo && rowSeri === seriNo) || (isKara && rowSeri === partNo)) {
        // "SON KONTROLÜ YAPAN FİRMA" is column 14 (1-based index 14, i.e., column N)
        techSheet.getRange(r + 2, 14).setValue(newFirma);
      }
    }
  }
}

/**
 * Her gün otomatik çalışacak tetikleyici fonksiyonu.
 * Sistem açık veya kapalı fark etmeksizin her gün kontrol eder ve e-posta gönderir.
 */
function dailyReminderTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var pad = function(n) { return String(n).padStart(2, '0'); };
  var formattedDate = pad(now.getDate()) + "." + pad(now.getMonth() + 1) + "." + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes());
  
  processEquipmentReminders(ss, formattedDate);
}

/**
 * Zaman ayarlı günlük tetikleyiciyi kurmak için yardımcı fonksiyon.
 * Google Apps Script Editörü üzerinden bir kez çalıştırılması yeterlidir.
 */
function createDailyTrigger() {
  // Mevcut tetikleyicileri temizleyelim
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "dailyReminderTrigger") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Her gün sabah 09:00 - 10:00 arasında çalışacak şekilde yeni tetikleyici kuralım
  ScriptApp.newTrigger("dailyReminderTrigger")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  // Kullanıcıya işlemin başarılı olduğunu belirten bir bildirim gösterelim
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      "⏰ OTOMATİK HATIRLATICI ETKİNLEŞTİRİLDİ",
      "Gelecek kontrol / bakım tarihleri için günlük otomatik e-posta taraması başarıyla kuruldu!\n\n" +
      "• Sistem her sabah 09:00 - 10:00 saatleri arasında çalışacaktır.\n" +
      "• Kontrol tarihine 90, 60 veya 30 gün kalan teçhizatlar için ilgili Sorumlu Birim yetkililerine otomatik toplu hatırlatma e-postaları gönderilecektir.",
      ui.ButtonSet.OK
    );
  } catch (err) {
    Logger.log("Arayüz bildirimi gösterilemedi: " + err.toString());
  }
}


