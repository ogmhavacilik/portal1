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
        if (fileNameLower.indexOf("tech_img_") === 0 || fileNameLower.indexOf(".xlsx") !== -1 || fileNameLower.indexOf(".xls") !== -1 || fileNameLower.indexOf(".csv") !== -1 || fileNameLower.indexOf("techizat") !== -1 || monthStr === "Teçhizat Takip" || formIdNum === 6) {
          folderId = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP"; // Teçhizat Takip & Excel Klasörü
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

    } else if (action === "uploadTechPublication") {
      var fileName = String(postData.fileName || "").trim();
      var base64Data = postData.base64Data;
      var unit = postData.unit || "GENEL";
      var unitKey = postData.unitKey || "GENEL";
      var category = postData.category || "IPC";
      var title = postData.title || fileName;
      var revision = postData.revision || "Rev. 01";
      var section = postData.section || "DİĞER";
      var notes = postData.notes || "";
      var originalFileName = postData.originalFileName || fileName;
      var folderId = postData.folderId || "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0";
      
      var folder = DriveApp.getFolderById(folderId);
      
      // Clean previous identical file if overwrite requested
      var existingFiles = folder.getFilesByName(fileName);
      while (existingFiles.hasNext()) {
        var file = existingFiles.next();
        try {
          file.setTrashed(true);
        } catch (delErr) {
          Logger.log("Eski teknik yayın silinemedi: " + delErr.toString());
        }
      }
      
      var decoded = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(decoded, "application/pdf", fileName);
      var newFile = folder.createFile(blob);
      
      // Store rich metadata in file description
      var metaObj = {
        title: title,
        unit: unit,
        unitKey: unitKey,
        category: category,
        revision: revision,
        section: section,
        notes: notes,
        uploadDate: formattedDate,
        originalFileName: originalFileName
      };
      try {
        newFile.setDescription(JSON.stringify(metaObj));
      } catch (descErr) {
        Logger.log("Description eklenemedi: " + descErr.toString());
      }
      
      var viewUrl = "https://drive.google.com/file/d/" + newFile.getId() + "/preview";
      var downloadUrl = "https://drive.google.com/uc?export=download&id=" + newFile.getId();
      
      // Sync to TEKNİK YAYINLAR Google Sheet
      syncTeknikYayinToSheet(ss, {
        fileId: newFile.getId(),
        fileName: fileName,
        unit: unit,
        unitKey: unitKey,
        category: category,
        title: title,
        revision: revision,
        section: section,
        notes: notes,
        uploadDate: formattedDate,
        viewUrl: viewUrl,
        downloadUrl: downloadUrl
      });
      
      recordLastUpdate(ss, "TEKNİK YAYINLAR - " + unit, formattedDate);
      
      response.message = "Teknik yayın '" + fileName + "' (" + unit + " / " + section + ") başarıyla Drive'a yüklendi ve e-tabloya kaydedildi.";
      response.fileId = newFile.getId();
      response.viewUrl = viewUrl;
      response.downloadUrl = downloadUrl;
      response.fileName = fileName;
      response.unit = unit;
      response.unitKey = unitKey;
      response.category = category;
      response.title = title;
      response.revision = revision;
      response.section = section;
      response.notes = notes;
      response.uploadDate = formattedDate;

    } else if (action === "deleteTechPublication") {
      var fileId = postData.fileId;
      var fileName = postData.fileName;
      var deleted = false;
      
      if (fileId) {
        try {
          var targetFile = DriveApp.getFileById(fileId);
          targetFile.setTrashed(true);
          deleted = true;
        } catch (fErr) {
          Logger.log("FileId ile silme hatası: " + fErr.toString());
        }
      }
      
      if (!deleted && fileName) {
        var folderId = postData.folderId || "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0";
        var folder = DriveApp.getFolderById(folderId);
        var files = folder.getFilesByName(fileName);
        while (files.hasNext()) {
          var f = files.next();
          try {
            f.setTrashed(true);
            deleted = true;
          } catch (e) {}
        }
      }
      
      // Delete corresponding row from TEKNİK YAYINLAR Google Sheet
      deleteTeknikYayinFromSheet(ss, fileId, fileName);
      
      response.deleted = deleted;
      response.message = deleted ? "Teknik yayın Drive'dan ve e-tablodan başarıyla silindi." : "Dosya Drive'da bulunamadı ancak e-tablo temizlendi.";

    } else if (action === "updateTechPublication" || action === "editTechPublication") {
      var fileId = postData.fileId;
      var fileName = postData.fileName;
      var title = postData.title;
      var unit = postData.unit || "GENEL";
      var unitKey = postData.unitKey || "GENEL";
      var category = postData.category || "IPC";
      var revision = postData.revision || "Rev. 01";
      var section = postData.section || "Genel Teknik Döküman";
      var notes = postData.notes || "";
      var updated = false;

      if (fileId) {
        try {
          var targetFile = DriveApp.getFileById(fileId);
          var metaObj = {
            title: title,
            unit: unit,
            unitKey: unitKey,
            category: category,
            revision: revision,
            section: section,
            notes: notes,
            uploadDate: formattedDate,
            originalFileName: fileName
          };
          targetFile.setDescription(JSON.stringify(metaObj));
          updated = true;
        } catch (uErr) {
          Logger.log("Drive açıklama güncelleme hatası: " + uErr.toString());
        }
      }

      // Update in TEKNİK YAYINLAR Google Sheet
      syncTeknikYayinToSheet(ss, {
        fileId: fileId,
        fileName: fileName,
        unit: unit,
        unitKey: unitKey,
        category: category,
        title: title,
        revision: revision,
        section: section,
        notes: notes,
        uploadDate: formattedDate
      });

      response.updated = true;
      response.message = "Teknik yayın bilgileri başarıyla güncellendi ve e-tabloya işlendi.";

    } else if (action === "renameTechPublication") {
      var fileId = postData.fileId;
      var newFileName = postData.newFileName;
      var renamed = false;
      
      if (fileId && newFileName) {
        try {
          var targetFile = DriveApp.getFileById(fileId);
          targetFile.setName(newFileName);
          renamed = true;
        } catch (rErr) {
          Logger.log("Teknik yayın yeniden adlandırma hatası: " + rErr.toString());
        }
      }
      
      if (renamed) {
        syncTeknikYayinToSheet(ss, {
          fileId: fileId,
          fileName: newFileName
        });
      }
      
      response.renamed = renamed;
      response.newFileName = newFileName;
      response.message = renamed ? "Teknik yayın adı başarıyla güncellendi." : "Dosya adı güncellenemedi.";

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
    } else if (action === "readExcelFromDrive") {
      try {
        var folderId = (postData && postData.folderId) || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var fileName = postData && postData.fileName;
        var fileId = postData && postData.fileId;
        var folder = null;
        try {
          folder = DriveApp.getFolderById(folderId);
        } catch (fErr) {
          folder = null;
        }
        var targetFile = null;
        if (fileId) {
          try {
            var f = DriveApp.getFileById(fileId);
            if (f && !f.isTrashed()) {
              targetFile = f;
            }
          } catch (idErr) {}
        } else if (fileName && folder) {
          var files = folder.getFilesByName(fileName);
          while (files.hasNext()) {
            var cand = files.next();
            if (!cand.isTrashed()) {
              targetFile = cand;
              break;
            }
          }
        }
        if (targetFile && !targetFile.isTrashed() && targetFile.getSize() > 0) {
          var blob = targetFile.getBlob();
          response.base64 = Utilities.base64Encode(blob.getBytes());
          response.fileName = targetFile.getName();
          response.fileId = targetFile.getId();
          response.updated = Utilities.formatDate(targetFile.getLastUpdated(), "GMT+3", "dd.MM.yyyy HH:mm:ss");
          response.message = "Excel dosyası başarıyla okundu.";
        } else {
          response.status = "error";
          response.message = "Dosya Drive klasöründe yüklü değil veya bulunamadı: " + (fileName || fileId);
        }
      } catch (err) {
        response.status = "error";
        response.message = "Drive Excel okuma hatası: " + err.toString();
      }

    } else if (action === "listTechizatExcelsFromDrive") {
      try {
        var folderId = (postData && postData.folderId) || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var folder = null;
        try {
          folder = DriveApp.getFolderById(folderId);
        } catch (fErr) {
          folder = null;
        }
        if (!folder) {
          response.status = "error";
          response.excels = [];
          response.message = "Drive klasörü bulunamadı veya erişilemez: " + folderId;
        } else {
          var files = folder.getFiles();
          var excelList = [];
          while (files.hasNext()) {
            var file = files.next();
            if (file.isTrashed()) {
              continue;
            }
            var fname = file.getName();
            var fnameLower = fname.toLowerCase();
            var isExcel = fnameLower.indexOf(".xlsx") !== -1 || fnameLower.indexOf(".xls") !== -1 || fnameLower.indexOf(".csv") !== -1;
            if (isExcel && file.getSize() > 0) {
              excelList.push({
                id: file.getId(),
                name: fname,
                size: file.getSize(),
                lastUpdated: Utilities.formatDate(file.getLastUpdated(), "GMT+3", "dd.MM.yyyy HH:mm:ss"),
                downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId()
              });
            }
          }
          response.excels = excelList;
          if (excelList.length > 0) {
            response.message = excelList.length + " adet yüklü Excel dosyası listelendi.";
          } else {
            response.message = "Drive klasöründe yüklü herhangi bir Excel dosyası bulunamadı.";
          }
        }
      } catch (err) {
        response.status = "error";
        response.excels = [];
        response.message = "Excel listeleme hatası: " + err.toString();
      }

    } else if (action === "uploadTechizatExcel" || action === "saveTechizatExcelToDrive") {
      try {
        var folderId = (postData && postData.folderId) || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var fileName = postData && postData.fileName;
        var targetKey = postData && postData.targetKey;
        var base64Data = postData && postData.base64Data;
        var folder = DriveApp.getFolderById(folderId);

        // Normalize filename
        if (!fileName) {
          fileName = (targetKey ? targetKey : "techizat_guncel") + ".xlsx";
        }
        if (!fileName.toLowerCase().endsWith(".xlsx") && !fileName.toLowerCase().endsWith(".xls")) {
          fileName += ".xlsx";
        }

        // 1. Delete exact match existing files
        var existingFiles = folder.getFilesByName(fileName);
        while (existingFiles.hasNext()) {
          var oldF = existingFiles.next();
          try { oldF.setTrashed(true); } catch (e) {}
        }

        // 2. Identify target key / category to purge any previous old-named versions
        var normalizedTarget = (targetKey || fileName).toLowerCase()
          .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u')
          .replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c')
          .replace(/[\-_]/g, ' ')
          .replace('.xlsx', '').replace('.xls', '');

        var allFilesInFolder = folder.getFiles();
        while (allFilesInFolder.hasNext()) {
          var fileInDrive = allFilesInFolder.next();
          var driveNameNorm = fileInDrive.getName().toLowerCase()
            .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u')
            .replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c')
            .replace(/[\-_]/g, ' ')
            .replace('.xlsx', '').replace('.xls', '');

          // Check if it belongs to the exact same aircraft unit & sub-section
          var isSameTarget = false;
          var unitTokens = ["at 802", "at802", "bell 429", "bell429", "t 70", "t70", "b 360", "b360", "c 650", "c650", "hangar", "kara arac"];
          var sectionTokens = ["ozel", "sarf", "kimyasal", "yer destek", "bumbi", "helitak"];

          var currentUnit = "";
          for (var u = 0; u < unitTokens.length; u++) {
            if (normalizedTarget.indexOf(unitTokens[u]) !== -1) {
              currentUnit = unitTokens[u].replace(/\s+/g, '');
              break;
            }
          }

          var driveUnit = "";
          for (var du = 0; du < unitTokens.length; du++) {
            if (driveNameNorm.indexOf(unitTokens[du]) !== -1) {
              driveUnit = unitTokens[du].replace(/\s+/g, '');
              break;
            }
          }

          if (currentUnit && driveUnit && currentUnit === driveUnit) {
            var currentSec = "";
            for (var s = 0; s < sectionTokens.length; s++) {
              if (normalizedTarget.indexOf(sectionTokens[s]) !== -1) {
                currentSec = sectionTokens[s];
                break;
              }
            }

            var driveSec = "";
            for (var ds = 0; ds < sectionTokens.length; ds++) {
              if (driveNameNorm.indexOf(sectionTokens[ds]) !== -1) {
                driveSec = sectionTokens[ds];
                break;
              }
            }

            if (currentSec === driveSec && currentSec !== "") {
              isSameTarget = true;
            }
          }

          if (isSameTarget) {
            try { fileInDrive.setTrashed(true); } catch (e) {}
          }
        }

        var decoded = Utilities.base64Decode(base64Data);
        var blob = Utilities.newBlob(decoded, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        var newFile = folder.createFile(blob);

        response.fileId = newFile.getId();
        response.fileName = newFile.getName();
        response.downloadUrl = "https://drive.google.com/uc?export=download&id=" + newFile.getId();
        response.message = "Excel dosyası '" + fileName + "' adıyla Google Drive klasörüne başarıyla kaydedildi. Varsa önceki eski sürüm silindi.";
      } catch (err) {
        response.status = "error";
        response.message = "Drive Excel yükleme hatası: " + err.toString();
      }

    } else if (action === "deleteTechizatExcelFromDrive") {
      try {
        var folderId = (postData && postData.folderId) || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var fileId = postData && postData.fileId;
        var fileName = postData && postData.fileName;
        var folder = DriveApp.getFolderById(folderId);
        var deleted = false;

        if (fileId) {
          try {
            var targetFile = DriveApp.getFileById(fileId);
            targetFile.setTrashed(true);
            deleted = true;
          } catch (fErr) {}
        }

        if (!deleted && fileName) {
          var files = folder.getFilesByName(fileName);
          while (files.hasNext()) {
            var f = files.next();
            try {
              f.setTrashed(true);
              deleted = true;
            } catch (e) {}
          }
        }

        response.deleted = deleted;
        response.message = deleted ? "Excel dosyası Drive'dan başarıyla silindi." : "Dosya bulunamadı.";
      } catch (err) {
        response.status = "error";
        response.message = "Excel silme hatası: " + err.toString();
      }

    } else if (action === "uploadTechizatDocPdf" || action === "uploadTechizatPdf") {
      // Teçhizat ve Hangar Yer Destek PDF Belge / Evrak Yükleme
      try {
        var folderId = (postData && postData.folderId) || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var fileName = postData && postData.fileName;
        var itemKey = (postData && postData.itemKey) || "GENEL";
        var docType = (postData && postData.docType) || "Evrak";
        var firma = (postData && postData.firma) || "";
        var base64Data = postData && postData.base64Data;
        var folder = DriveApp.getFolderById(folderId);

        if (!fileName) {
          fileName = "TECH_DOC_" + itemKey.replace(/[^a-zA-Z0-9]/g, '_') + "_" + Date.now() + ".pdf";
        }
        if (!fileName.toLowerCase().endsWith(".pdf")) {
          fileName += ".pdf";
        }

        var decoded = Utilities.base64Decode(base64Data);
        var blob = Utilities.newBlob(decoded, "application/pdf", fileName);
        var newFile = folder.createFile(blob);

        var metaObj = {
          type: "TECHIZAT_PDF_DOC",
          itemKey: itemKey,
          docType: docType,
          firma: firma,
          uploadDate: formattedDate,
          originalFileName: fileName
        };
        try {
          newFile.setDescription(JSON.stringify(metaObj));
        } catch (dErr) {}

        var viewUrl = "https://drive.google.com/file/d/" + newFile.getId() + "/preview";
        var downloadUrl = "https://drive.google.com/uc?export=download&id=" + newFile.getId();

        response.status = "success";
        response.fileId = newFile.getId();
        response.fileName = newFile.getName();
        response.viewUrl = viewUrl;
        response.downloadUrl = downloadUrl;
        response.itemKey = itemKey;
        response.docType = docType;
        response.firma = firma;
        response.uploadDate = formattedDate;
        response.message = "PDF belge '" + fileName + "' adıyla Google Drive'a başarıyla yüklendi.";
      } catch (err) {
        response.status = "error";
        response.message = "Drive PDF belge yükleme hatası: " + err.toString();
      }

    } else if (action === "deleteTechizatDocPdf") {
      // Drive'dan Teçhizat PDF Belgesi Silme
      try {
        var folderId = (postData && postData.folderId) || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var fileId = postData && postData.fileId;
        var fileName = postData && postData.fileName;
        var deleted = false;

        if (fileId) {
          try {
            var targetFile = DriveApp.getFileById(fileId);
            targetFile.setTrashed(true);
            deleted = true;
          } catch (fErr) {}
        }

        if (!deleted && fileName) {
          var folder = DriveApp.getFolderById(folderId);
          var files = folder.getFilesByName(fileName);
          while (files.hasNext()) {
            var f = files.next();
            try {
              f.setTrashed(true);
              deleted = true;
            } catch (e) {}
          }
        }

        response.deleted = deleted;
        response.message = deleted ? "PDF belgesi Drive'dan başarıyla silindi." : "Silinecek PDF bulunamadı.";
      } catch (err) {
        response.status = "error";
        response.message = "PDF silme hatası: " + err.toString();
      }

    } else if (action === "saveDepoTransfers" || action === "updateDepoTransfers") {
      try {
        var transfers = postData.transfers || postData.data || [];
        var targetSpreadsheetId = (postData && postData.spreadsheetId) || "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
        var targetSs = ss;
        try {
          targetSs = SpreadsheetApp.openById(targetSpreadsheetId);
        } catch (ssOpenErr) {
          targetSs = ss;
        }

        // 1. Ana Transfer Geçmişi Tablosunu Yaz (Tüm transferler)
        writeTransfersToSpecificSheet(targetSs, "TRANSFER GEÇMİŞİ", transfers, "#0b3d1d");

        // 2. Depo / Birim Bazlı Ayrı Sayfalar Oluştur ve Yaz
        var unitCategories = {
          "TRANSFER GEÇMİŞİ - AT802": [],
          "TRANSFER GEÇMİŞİ - BELL429": [],
          "TRANSFER GEÇMİŞİ - T70": [],
          "TRANSFER GEÇMİŞİ - B360": [],
          "TRANSFER GEÇMİŞİ - C650": [],
          "TRANSFER GEÇMİŞİ - HANGAR": [],
          "TRANSFER GEÇMİŞİ - KARA ARAÇLARI": []
        };

        for (var trIdx = 0; trIdx < transfers.length; trIdx++) {
          var tItem = transfers[trIdx];
          var unitKey = getDepoTransferUnitKey(tItem);
          if (unitKey === "at802") unitCategories["TRANSFER GEÇMİŞİ - AT802"].push(tItem);
          else if (unitKey === "bell429") unitCategories["TRANSFER GEÇMİŞİ - BELL429"].push(tItem);
          else if (unitKey === "t70") unitCategories["TRANSFER GEÇMİŞİ - T70"].push(tItem);
          else if (unitKey === "b360") unitCategories["TRANSFER GEÇMİŞİ - B360"].push(tItem);
          else if (unitKey === "c650") unitCategories["TRANSFER GEÇMİŞİ - C650"].push(tItem);
          else if (unitKey === "hangar") unitCategories["TRANSFER GEÇMİŞİ - HANGAR"].push(tItem);
          else if (unitKey === "kara_araclari") unitCategories["TRANSFER GEÇMİŞİ - KARA ARAÇLARI"].push(tItem);
          else unitCategories["TRANSFER GEÇMİŞİ - AT802"].push(tItem);
        }

        var headerColors = {
          "TRANSFER GEÇMİŞİ - AT802": "#1e3a8a",
          "TRANSFER GEÇMİŞİ - BELL429": "#065f46",
          "TRANSFER GEÇMİŞİ - T70": "#831843",
          "TRANSFER GEÇMİŞİ - B360": "#78350f",
          "TRANSFER GEÇMİŞİ - C650": "#4c1d95",
          "TRANSFER GEÇMİŞİ - HANGAR": "#334155",
          "TRANSFER GEÇMİŞİ - KARA ARAÇLARI": "#7c2d12"
        };

        for (var catSheetName in unitCategories) {
          if (unitCategories[catSheetName].length > 0) {
            writeTransfersToSpecificSheet(targetSs, catSheetName, unitCategories[catSheetName], headerColors[catSheetName] || "#0b3d1d");
          }
        }

        // 3. AT-802 Transfer Geçmişini Google Drive Klasörüne (1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP) Excel (.xlsx) olarak yükle
        try {
          var targetFolderId = (postData && postData.folderId) || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
          var targetFolder = DriveApp.getFolderById(targetFolderId);
          var at802Transfers = unitCategories["TRANSFER GEÇMİŞİ - AT802"] || [];
          if (at802Transfers.length === 0 && transfers.length > 0) {
            at802Transfers = transfers;
          }

          if (postData && postData.base64Data) {
            var decoded = Utilities.base64Decode(postData.base64Data);
            var blob = Utilities.newBlob(decoded, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "transfer_gecmisi_at802.xlsx");
            
            var existingFiles = targetFolder.getFilesByName("transfer_gecmisi_at802.xlsx");
            while (existingFiles.hasNext()) {
              try { existingFiles.next().setTrashed(true); } catch(e){}
            }
            var existingAlt = targetFolder.getFilesByName("TRANSFER GEÇMİŞİ _AT-802.xlsx");
            while (existingAlt.hasNext()) {
              try { existingAlt.next().setTrashed(true); } catch(e){}
            }
            targetFolder.createFile(blob);
          } else if (at802Transfers.length > 0) {
            var tempSs = SpreadsheetApp.create("TRANSFER GEÇMİŞİ _AT-802");
            writeTransfersToSpecificSheet(tempSs, "TRANSFER GEÇMİŞİ _AT-802", at802Transfers, "#1e3a8a");
            tempSs.deleteSheet(tempSs.getSheets()[0]);
            SpreadsheetApp.flush();
            
            var tempFile = DriveApp.getFileById(tempSs.getId());
            var excelBlob = tempFile.getBlob().getAs("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            excelBlob.setName("transfer_gecmisi_at802.xlsx");

            var existingFiles = targetFolder.getFilesByName("transfer_gecmisi_at802.xlsx");
            while (existingFiles.hasNext()) {
              try { existingFiles.next().setTrashed(true); } catch(e){}
            }
            targetFolder.createFile(excelBlob);
            try { tempFile.setTrashed(true); } catch(e){}
          }
        } catch (excelErr) {
          Logger.log("AT-802 Excel Drive yazma uyarısı: " + excelErr.toString());
        }

        recordLastUpdate(targetSs, "TRANSFER VE İŞLEM GEÇMİŞİ", formattedDate);
        response.message = (transfers ? transfers.length : 0) + " adet transfer kaydı E-Tabloya ve Drive klasörüne (transfer_gecmisi_at802.xlsx) başarıyla aktarıldı.";
        response.count = transfers ? transfers.length : 0;
      } catch (trErr) {
        response.status = "error";
        response.message = "Transfer geçmişi kaydetme hatası: " + trErr.toString();
      }

    } else if (action === "readDepoTransfers" || action === "getDepoTransfers") {
      try {
        var targetSpreadsheetId = (postData && postData.spreadsheetId) || "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
        var targetSs = SpreadsheetApp.openById(targetSpreadsheetId);
        var sheet = targetSs.getSheetByName("TRANSFER GEÇMİŞİ - AT802") || targetSs.getSheetByName("TRANSFER GEÇMİŞİ");
        if (sheet && sheet.getLastRow() > 1) {
          var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
          response.transfers = rows.map(function(r) {
            return {
              id: "tx_" + String(r[0] || Math.random()),
              date: String(r[1] || ""),
              timestamp: String(r[1] || ""),
              type: String(r[2] || "TRANSFER"),
              itemName: String(r[3] || ""),
              pn: String(r[4] || ""),
              sn: String(r[5] || ""),
              quantity: Number(r[6]) || 1,
              tailNo: String(r[7] || ""),
              sourceLocation: String(r[8] || "DEPO"),
              targetLocation: String(r[9] || ""),
              location: String(r[9] || r[8] || "DEPO"),
              operator: String(r[10] || ""),
              receivedBy: String(r[11] || ""),
              notes: String(r[12] || "")
            };
          });
          response.status = "success";
        } else {
          response.transfers = [];
          response.status = "success";
        }
      } catch (rErr) {
        response.status = "error";
        response.message = rErr.toString();
      }

    } else if (action === "updateGunTakip") {
      var data = postData.data; // Array of [SORUMLU BİRİM, ADI SOYADI, E-POSTA ADRESİ]
      var targetSs = ss; // Her zaman aktif e-tabloyu kullan
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
          
          var mail90Val = (row[3] !== undefined && row[3] !== null && String(row[3]).trim() !== "") ? String(row[3]).trim() : saved.mail90;
          var mail60Val = (row[4] !== undefined && row[4] !== null && String(row[4]).trim() !== "") ? String(row[4]).trim() : saved.mail60;
          var mail30Val = (row[5] !== undefined && row[5] !== null && String(row[5]).trim() !== "") ? String(row[5]).trim() : saved.mail30;

          return [
            rowBirim,
            row[1] !== undefined && row[1] !== null ? String(row[1]).trim() : "",
            row[2] !== undefined && row[2] !== null ? String(row[2]).trim() : "",
            mail90Val,
            mail60Val,
            mail30Val
          ];
        });
        
        sheet.getRange(startRow, 1, numRows, numCols).setValues(valuesToInsert);
      }
      
      formatGeneralSheet(sheet);
      recordLastUpdate(targetSs, "GÜN TAKİP", formattedDate);
      
      response.message = "GÜN TAKİP (Sorumlu Birim Mail Ayarları) başarıyla güncellendi.";

    } else if (action === "recordUpdateTimestamp") {
      var unitName = postData.unitName || "GENEL";
      var formattedDateToRecord = postData.formattedDate || Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy HH:mm");
      recordLastUpdate(ss, unitName, formattedDateToRecord);
      response.status = "success";
      response.message = "Güncelleme tarihi başarıyla kaydedildi.";

    } else if (action === "sendCustomBirimReminder") {
      var targetBirim = postData.birim || "";
      var adSoyad = postData.adSoyad || "Sorumlu Personel";
      var eposta = postData.eposta || "orhavak.bakimsube@gmail.com";
      var reminderItems = postData.items || [];
      var reminderTimestamp = postData.timestamp || Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy HH:mm:ss");
      
      sendCustomBirimReminderEmail(eposta, adSoyad, targetBirim, reminderItems, reminderTimestamp);
      
      // GÜN TAKİP sayfasında 90 Gün kolonuna saniyeli tarihi yaz
      var gtSheet = getSheetWithFallback(ss, "GÜN TAKİP");
      var gtLastRow = gtSheet.getLastRow();
      if (gtLastRow > 1) {
        var gtValues = gtSheet.getRange(2, 1, gtLastRow - 1, 6).getValues();
        for (var r = 0; r < gtValues.length; r++) {
          if (String(gtValues[r][0]).trim().toUpperCase() === targetBirim.trim().toUpperCase()) {
            gtSheet.getRange(r + 2, 4).setValue(reminderTimestamp);
            break;
          }
        }
      }
      
      response.status = "success";
      response.message = "E-posta başarıyla gönderildi ve saat/saniye ile kayıt altına alındı.";
    } else if (action === "appendEbysTable") {
      var ebysNo = String(postData.ebysNo || "").trim();
      var tableData = postData.data; // Array of [Ait Olduğu Birim, Teçhizat Adı, Parça No, Seri No, Miktar, Son Kontrolü Yapan Firma, Açıklama]
      
      var targetSs = ss; // Her zaman aktif e-tabloyu kullan
      
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
      var targetSs = ss; // Her zaman aktif e-tabloyu kullan
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
          if (file.isTrashed() || file.getSize() === 0) {
            continue;
          }
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

    } else if (action === "listTechPublications") {
      var list = [];
      try {
        var techSheet = ensureTeknikYayinlarSheet(ss);
        var lastRow = techSheet.getLastRow();
        
        // 1. Read existing records from TEKNİK YAYINLAR Google Sheet
        var sheetMap = {};
        if (lastRow > 1) {
          var allValues = techSheet.getRange(2, 1, lastRow - 1, 13).getValues();
          for (var i = 0; i < allValues.length; i++) {
            var row = allValues[i];
            var fId = String(row[9] || "").trim();
            var fName = String(row[7] || "").trim();
            var itemKey = fId || fName;
            if (!itemKey) continue;
            
            var item = {
              id: fId ? "drive_" + fId : "sheet_" + (i + 1),
              title: String(row[4] || fName),
              unit: String(row[1] || "GENEL"),
              unitLabel: String(row[2] || "TÜM HAVA ARAÇLARI"),
              category: String(row[3] || "IPC"),
              revision: String(row[5] || "Rev. 01"),
              section: String(row[6] || "Genel Teknik Döküman"),
              fileName: fName,
              fileSize: "1.0 MB",
              uploadDate: String(row[8] || ""),
              driveFileId: fId,
              viewUrl: String(row[10] || (fId ? "https://drive.google.com/file/d/" + fId + "/preview" : "")),
              downloadUrl: String(row[11] || (fId ? "https://drive.google.com/uc?export=download&id=" + fId : "")),
              notes: String(row[12] || ""),
              lastUpdated: String(row[8] || "")
            };
            list.push(item);
            if (fId) sheetMap[fId] = true;
            if (fName) sheetMap[fName] = true;
          }
        }
        
        // 2. Cross-check with Drive folder to sync any new PDFs
        var folderId = e.parameter.folderId || "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0";
        var folder = DriveApp.getFolderById(folderId);
        var files = folder.getFiles();
        while (files.hasNext()) {
          var file = files.next();
          var name = file.getName();
          var nameLower = name.toLowerCase();
          if (nameLower.endsWith(".pdf") && !sheetMap[file.getId()] && !sheetMap[name]) {
            var desc = "";
            try {
              desc = file.getDescription() || "";
            } catch (e) {}
            
            var meta = {};
            if (desc) {
              try {
                meta = JSON.parse(desc);
              } catch (parseErr) {}
            }
            
            var sizeBytes = 0;
            try {
              sizeBytes = file.getSize();
            } catch (sErr) {}
            var sizeStr = sizeBytes > 0 ? (sizeBytes / (1024 * 1024)).toFixed(1) + " MB" : "1.0 MB";
            
            var parsedTitle = meta.title || meta.originalFileName || name.replace(/\.[^/.]+$/, "").replace(/^pub_[^_]+_/, "").replace(/_/g, " ");
            var parsedUnitKey = meta.unitKey || "GENEL";
            var parsedUnitLabel = meta.unit || "TÜM HAVA ARAÇLARI";
            if (!meta.unitKey && nameLower.indexOf("802") !== -1) {
              parsedUnitKey = "AT802";
              parsedUnitLabel = "AT-802 UÇAKLARI";
            } else if (!meta.unitKey && (nameLower.indexOf("t70") !== -1 || nameLower.indexOf("t-70") !== -1)) {
              parsedUnitKey = "T70";
              parsedUnitLabel = "T-70 HELİKOPTERLERİ";
            } else if (!meta.unitKey && nameLower.indexOf("429") !== -1) {
              parsedUnitKey = "B429";
              parsedUnitLabel = "BELL 429 HELİKOPTERLERİ";
            } else if (!meta.unitKey && (nameLower.indexOf("c650") !== -1 || nameLower.indexOf("650") !== -1)) {
              parsedUnitKey = "C650";
              parsedUnitLabel = "CITATION 650 (UÇAK)";
            } else if (!meta.unitKey && (nameLower.indexOf("b360") !== -1 || nameLower.indexOf("360") !== -1)) {
              parsedUnitKey = "B360";
              parsedUnitLabel = "BEECHCRAFT B360 (UÇAK)";
            } else if (!meta.unitKey && nameLower.indexOf("hangar") !== -1) {
              parsedUnitKey = "HANGAR";
              parsedUnitLabel = "HANGAR & YER EKİPMANLARI";
            }
            
            var parsedCategory = meta.category || "IPC";
            if (!meta.category) {
              if (nameLower.indexOf("amm") !== -1) parsedCategory = "AMM";
              else if (nameLower.indexOf("cmm") !== -1) parsedCategory = "CMM";
              else if (nameLower.indexOf("srm") !== -1) parsedCategory = "SRM";
              else if (nameLower.indexOf("wdm") !== -1) parsedCategory = "WDM";
              else if (nameLower.indexOf("ndtm") !== -1) parsedCategory = "NDTM";
            }
            
            var parsedRevision = meta.revision || "Rev. 01";
            var parsedSection = meta.section || "Genel Teknik Döküman";
            var uploadDateFormatted = meta.uploadDate || Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), "yyyy-MM-dd");
            var vUrl = "https://drive.google.com/file/d/" + file.getId() + "/preview";
            var dUrl = "https://drive.google.com/uc?export=download&id=" + file.getId();
            
            var newItem = {
              id: "drive_" + file.getId(),
              title: parsedTitle,
              unit: parsedUnitKey,
              unitLabel: parsedUnitLabel,
              category: parsedCategory,
              revision: parsedRevision,
              section: parsedSection,
              fileName: name,
              fileSize: sizeStr,
              uploadDate: uploadDateFormatted,
              driveFileId: file.getId(),
              viewUrl: vUrl,
              downloadUrl: dUrl,
              notes: meta.notes || "",
              lastUpdated: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm")
            };
            
            list.push(newItem);
            
            // Sync this discovered PDF into TEKNİK YAYINLAR Google Sheet
            syncTeknikYayinToSheet(ss, {
              fileId: file.getId(),
              fileName: name,
              unit: parsedUnitLabel,
              unitKey: parsedUnitKey,
              category: parsedCategory,
              title: parsedTitle,
              revision: parsedRevision,
              section: parsedSection,
              notes: meta.notes || "",
              uploadDate: uploadDateFormatted,
              viewUrl: vUrl,
              downloadUrl: dUrl
            });
          }
        }
      } catch (driveErr) {
        Logger.log("Teknik yayın listeleme hatası: " + driveErr.toString());
      }
      response.data = list;
      response.message = list.length + " adet teknik yayın listelendi.";

    } else if (action === "readExcelFromDrive") {
      try {
        var folderId = e.parameter.folderId || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var fileName = e.parameter.fileName;
        var fileId = e.parameter.fileId;
        var folder = null;
        try {
          folder = DriveApp.getFolderById(folderId);
        } catch (fErr) {
          folder = null;
        }
        var targetFile = null;
        if (fileId) {
          try {
            var f = DriveApp.getFileById(fileId);
            if (f && !f.isTrashed()) {
              targetFile = f;
            }
          } catch (idErr) {}
        } else if (fileName && folder) {
          var files = folder.getFilesByName(fileName);
          while (files.hasNext()) {
            var cand = files.next();
            if (!cand.isTrashed()) {
              targetFile = cand;
              break;
            }
          }
        }
        if (targetFile && !targetFile.isTrashed() && targetFile.getSize() > 0) {
          var blob = targetFile.getBlob();
          response.base64 = Utilities.base64Encode(blob.getBytes());
          response.fileName = targetFile.getName();
          response.fileId = targetFile.getId();
          response.updated = Utilities.formatDate(targetFile.getLastUpdated(), "GMT+3", "dd.MM.yyyy HH:mm:ss");
          response.message = "Excel dosyası başarıyla okundu.";
        } else {
          response.status = "error";
          response.message = "Dosya Drive klasöründe yüklü değil veya bulunamadı: " + (fileName || fileId);
        }
      } catch (err) {
        response.status = "error";
        response.message = "Drive Excel okuma hatası: " + err.toString();
      }

    } else if (action === "listTechizatExcelsFromDrive") {
      try {
        var folderId = e.parameter.folderId || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var folder = null;
        try {
          folder = DriveApp.getFolderById(folderId);
        } catch (fErr) {
          folder = null;
        }
        if (!folder) {
          response.status = "error";
          response.excels = [];
          response.message = "Drive klasörü bulunamadı veya erişilemez: " + folderId;
        } else {
          var files = folder.getFiles();
          var excelList = [];
          while (files.hasNext()) {
            var file = files.next();
            if (file.isTrashed()) {
              continue;
            }
            var fname = file.getName();
            var fnameLower = fname.toLowerCase();
            var isExcel = fnameLower.indexOf(".xlsx") !== -1 || fnameLower.indexOf(".xls") !== -1 || fnameLower.indexOf(".csv") !== -1;
            if (isExcel && file.getSize() > 0) {
              excelList.push({
                id: file.getId(),
                name: fname,
                size: file.getSize(),
                lastUpdated: Utilities.formatDate(file.getLastUpdated(), "GMT+3", "dd.MM.yyyy HH:mm:ss"),
                downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId()
              });
            }
          }
          response.excels = excelList;
          if (excelList.length > 0) {
            response.message = excelList.length + " adet yüklü Excel dosyası listelendi.";
          } else {
            response.message = "Drive klasöründe yüklü herhangi bir Excel dosyası bulunamadı.";
          }
        }
      } catch (err) {
        response.status = "error";
        response.excels = [];
        response.message = "Excel listeleme hatası: " + err.toString();
      }

    } else if (action === "listImagesFromDrive") {
      var imagesMap = {};
      try {
        var folderId = (e.parameter && e.parameter.folderId) || "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";
        var folder = DriveApp.getFolderById(folderId);
        var files = folder.getFiles();
        while (files.hasNext()) {
          var file = files.next();
          try {
            if (file.isTrashed() || file.getSize() === 0) continue;
          } catch(ignore) {}
          var name = file.getName();
          var nameLower = name.toLowerCase();
          var mime = "";
          try { mime = file.getMimeType(); } catch(mErr) {}
          
          var isImg = (mime && mime.indexOf("image/") !== -1) || 
                      nameLower.indexOf(".png") !== -1 || 
                      nameLower.indexOf(".jpg") !== -1 || 
                      nameLower.indexOf(".jpeg") !== -1 || 
                      nameLower.indexOf(".webp") !== -1;
          
          if (!isImg) continue;

          var fileId = file.getId();
          var directUrl = "https://lh3.googleusercontent.com/d/" + fileId;
          var thumbUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1200";

          // 1. Raw name
          imagesMap[name] = directUrl;
          imagesMap[name + "_id"] = fileId;
          imagesMap[name + "_thumb"] = thumbUrl;

          // 2. Clean key without extension
          var cleanKey = name;
          var dotIdx = cleanKey.lastIndexOf(".");
          if (dotIdx !== -1) {
            cleanKey = cleanKey.substring(0, dotIdx);
          }
          imagesMap[cleanKey] = directUrl;
          imagesMap[cleanKey + "_id"] = fileId;
          imagesMap[cleanKey + "_thumb"] = thumbUrl;

          // 3. If prefixed with tech_img_
          if (cleanKey.toLowerCase().indexOf("tech_img_") === 0) {
            var subKey = cleanKey.substring("tech_img_".length);
            var parts = subKey.split("_");
            if (parts.length > 1) {
              var lastPart = parts[parts.length - 1];
              if (!isNaN(lastPart) && lastPart.length >= 10) {
                parts.pop();
                subKey = parts.join("_");
              }
            }
            imagesMap[subKey] = directUrl;
            imagesMap[subKey + "_id"] = fileId;
            imagesMap[subKey + "_thumb"] = thumbUrl;
          }

          // 4. Normalize AT-802 and Hangar keys
          var keyLower = cleanKey.toLowerCase();
          if (keyLower.indexOf("at-802") !== -1 || keyLower.indexOf("at802") !== -1) {
            var k1 = cleanKey.replace(/at-802/gi, "at802");
            var k2 = cleanKey.replace(/at802/gi, "at-802");
            imagesMap[k1] = directUrl;
            imagesMap[k2] = directUrl;
            imagesMap[k1.toLowerCase()] = directUrl;
            imagesMap[k2.toLowerCase()] = directUrl;
          }
          if (keyLower.indexOf("hangar") !== -1) {
            imagesMap[cleanKey] = directUrl;
            imagesMap[cleanKey.toLowerCase()] = directUrl;
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
      var targetSs = ss; // Her zaman aktif e-tabloyu kullan
      
      // Eğer Personel Bilgi çizelgesi isteniyorsa, öncelikle Drive'daki en güncel Excel dosyasını otomatik olarak senkronize et!
      if (sheetName === "5-Personel_Bilgi" || (sheetName && sheetName.toLowerCase().indexOf("personel") !== -1)) {
        try {
          syncPersonnelExcelToGoogleSheet();
        } catch (syncErr) {
          Logger.log("Okuma esnasında Excel senkronizasyonu başarısız: " + syncErr.toString());
        }
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
    var techSheet = ss.getSheetByName("TÜM TECHİZAT") || ss.getSheetByName("TÜM TEÇHİZAT");
    if (!techSheet) return;
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
    
    // Ensure all 19 columns exist in TÜM TECHİZAT
    var lastColTech = techSheet.getLastColumn();
    if (lastColTech < 19) {
      techSheet.getRange(1, 16).setValue("90 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      techSheet.getRange(1, 17).setValue("60 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      techSheet.getRange(1, 18).setValue("30 GÜN UYARISI MAİL GÖNDERİM TARİHİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      techSheet.getRange(1, 19).setValue("BÖLÜM / KATEGORİ").setFontWeight("bold").setBackground("#0f3d1d").setFontColor("#ffffff");
      lastColTech = 19;
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
    var isNew = item.isNew === true;
    var rowBg = isNew ? "#fef08a" : (i % 2 === 0 ? "#ffffff" : "#f8fafc");
    var rowBorder = isNew ? "2px solid #ca8a04" : "1px solid #e2e8f0";

    tableRowsHtml += 
      "<tr style='background-color: " + rowBg + "; border-bottom: " + rowBorder + ";'>" +
        "<td style='padding: 10px 8px; font-weight: bold; font-size: 11px; text-align: center; color: #64748b;'>" + 
          (i + 1) + 
          (isNew ? "<br><span style='background-color: #ca8a04; color: #ffffff; padding: 2px 6px; border-radius: 4px; font-size: 9px; font-weight: 800;'>⭐ YENİ</span>" : "") + 
        "</td>" +
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
 * Birime özel 90 gün uyarı e-postası gönderir.
 * Yeni eklenen veya daha önce bildirilmemiş malzemeler SARI DOLGU ile işaretlenir.
 */
function sendCustomBirimReminderEmail(email, personName, unitName, items, timestamp) {
  var subject = "🚨 90 GÜN BAKIM / KALİBRASYON / ÖMÜR BİTİŞ UYARISI - [" + unitName + "]";
  var colorHex = "#16a34a"; // 90 Gün Yeşili

  var tableRowsHtml = "";
  var newCount = 0;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var isNew = item.isNew === true;
    if (isNew) newCount++;

    var rowBg = isNew ? "#fef08a" : (i % 2 === 0 ? "#ffffff" : "#f8fafc");
    var rowBorder = isNew ? "2px solid #ca8a04" : "1px solid #e2e8f0";

    tableRowsHtml += 
      "<tr style='background-color: " + rowBg + "; border-bottom: " + rowBorder + ";'>" +
        "<td style='padding: 10px 8px; font-weight: bold; font-size: 11px; text-align: center; color: #475569;'>" + 
          (i + 1) + 
          (isNew ? "<br><span style='background-color: #ca8a04; color: #ffffff; padding: 2px 6px; border-radius: 4px; font-size: 9px; font-weight: 800;'>⭐ YENİ</span>" : "") + 
        "</td>" +
        "<td style='padding: 10px 8px; font-weight: 700; font-size: 12px; color: #0f172a;'>" + (item.techName || "Teçhizat / Malzeme") + "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; font-family: monospace; color: #334155;'>" + (item.partNo || "-") + "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; font-family: monospace; color: #334155;'>" + (item.seriNo || "-") + "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; color: #334155;'>" + (item.location || "-") + "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; font-weight: 700; color: #dc2626; text-align: center;'>" + 
          (item.gelecekTarih || "-") + "<br><span style='font-size: 10px; font-weight: 600; color: #b45309;'>(" + (item.daysDiff !== undefined ? item.daysDiff : "") + " gün kaldı)</span>" + 
        "</td>" +
        "<td style='padding: 10px 8px; font-size: 11px; text-align: center;'>" + 
          (isNew ? "<span style='background-color: #fef08a; color: #713f12; padding: 3px 6px; border-radius: 6px; font-weight: 800; font-size: 10px; border: 1px solid #eab308;'>YENİ EKLENDİ</span>" : "<span style='color: #64748b; font-size: 10px;'>Önceden Bildirildi</span>") + 
        "</td>" +
      "</tr>";
  }

  var body = 
    "<html>" +
    "<head>" +
      "<style>" +
        "body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; color: #1e293b; padding: 20px; }" +
        ".container { max-width: 900px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; border: 1px solid #e2e8f0; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }" +
        ".header { background-color: #0b3d1d; padding: 30px 20px; text-align: center; color: #ffffff; }" +
        ".header h2 { margin: 0; font-size: 22px; font-weight: 800; letter-spacing: -0.5px; text-transform: uppercase; }" +
        ".header p { margin: 5px 0 0 0; font-size: 13px; font-weight: 500; opacity: 0.9; color: #86efac; }" +
        ".content { padding: 30px 24px; }" +
        ".intro { font-size: 14px; font-weight: 500; color: #475569; margin-bottom: 25px; line-height: 1.6; }" +
        ".alert-box { background-color: #fef3c7; border: 1px solid #fde68a; border-radius: 12px; padding: 14px; margin-bottom: 20px; font-size: 12px; color: #92400e; }" +
        ".items-table { width: 100%; border-collapse: collapse; margin-bottom: 25px; text-align: left; }" +
        ".items-table th { background-color: #0b3d1d; padding: 12px 8px; font-size: 11px; font-weight: bold; color: #ffffff; text-transform: uppercase; border-bottom: 2px solid #072612; }" +
        ".footer { background-color: #f8fafc; padding: 20px; text-align: center; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; }" +
      "</style>" +
    "</head>" +
    "<body>" +
      "<div class='container'>" +
        "<div class='header'>" +
          "<h2>🚁 OGM HAVACILIK BAKIM VE TEKNİK ŞUBE MÜDÜRLÜĞÜ</h2>" +
          "<p>90 GÜN BAKIM, KALİBRASYON VE ÖMÜR BİTİŞ UYARI BİLDİRİMİ</p>" +
        "</div>" +
        "<div class='content'>" +
          "<div class='intro'>Sayın <strong>" + personName + "</strong>,<br><br>" +
          "<strong>" + unitName + "</strong> biriminize ait bakım, kalibrasyon veya raf ömrü bitiş tarihine <strong>90 gün ve daha az</strong> süre kalmış olan malzemelerin güncel takip listesi aşağıda yer almaktadır.<br>" +
          (newCount > 0 ? "⚠️ <strong>DİKKAT:</strong> Listeye yeni eklenen ve ilk kez bildirilen <strong>" + newCount + " adet</strong> malzeme tabloda <strong>SARI DOLGU</strong> ile vurgulanmıştır." : "") +
          "</div>" +
          
          "<table class='items-table'>" +
            "<thead>" +
              "<tr>" +
                "<th style='width: 6%; text-align: center;'>SIRA</th>" +
                "<th style='width: 26%;'>MALZEME / TEÇHİZAT ADI</th>" +
                "<th style='width: 15%;'>PARÇA NO (P/N)</th>" +
                "<th style='width: 14%;'>SERİ NO (S/N)</th>" +
                "<th style='width: 14%;'>BULUNDUĞU YER</th>" +
                "<th style='width: 13%; text-align: center;'>SONRAKİ TARİH</th>" +
                "<th style='width: 12%; text-align: center;'>BİLDİRİM</th>" +
              "</tr>" +
            "</thead>" +
            "<tbody>" +
              tableRowsHtml +
            "</tbody>" +
          "</table>" +
          
          "<div style='background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 12px; padding: 15px; text-align: center; font-size: 11px; font-weight: bold; color: #166534;'>" +
            "✅ Lütfen gerekli kontrolleri ve bakımları tamamlayıp portal üzerinden güncelleyiniz." +
          "</div>" +
        "</div>" +
        "<div class='footer'>" +
          "T.C. ORMAN GENEL MÜDÜRLÜĞÜ - HAVA ARAÇLARI BAKIM VE TEKNİK ŞUBE MÜDÜRLÜĞÜ<br>" +
          "Bu e-posta sistem tarafından otomatik olarak üretilmiştir.<br>" +
          "<span style='font-size: 10px; font-weight: bold; color: #64748b;'>Gönderim Tarih ve Saati (Saniye Dahil): " + timestamp + "</span>" +
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
    Logger.log("Birim Mail gönderim hatası (" + email + "): " + err.toString());
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

/**
 * Google Drive'daki en güncel Personel Bilgi excel dosyasını (.xlsx) arar,
 * geçici olarak Google Sheets formatına dönüştürür, verilerini okur ve
 * '5-Personel_Bilgi' sayfasına senkronize eder.
 */
function syncPersonnelExcelToGoogleSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("5-Personel_Bilgi");
    if (!sheet) {
      Logger.log("5-Personel_Bilgi sayfası bulunamadı, oluşturuluyor...");
      sheet = ss.insertSheet("5-Personel_Bilgi");
    }
    
    // Google Drive'da personel bilgi dosyası ara (.xlsx veya .xls)
    var files = DriveApp.searchFiles("title contains 'personel' and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'application/vnd.ms-excel' or title contains '.xlsx' or title contains '.xls')");
    var excelFile = null;
    var latestTime = 0;
    
    while (files.hasNext()) {
      var file = files.next();
      if (file.isTrashed() || file.getSize() === 0) continue;
      var lastUpdated = file.getLastUpdated().getTime();
      if (lastUpdated > latestTime) {
        latestTime = lastUpdated;
        excelFile = file;
      }
    }
    
    // Eğer genel aramada bulunamadıysa, form kayıtları klasörünü ("1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0") tara
    if (!excelFile) {
      try {
        var folder = DriveApp.getFolderById("1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0");
        var fFiles = folder.getFiles();
        while (fFiles.hasNext()) {
          var f = fFiles.next();
          if (f.isTrashed() || f.getSize() === 0) continue;
          var name = f.getName().toLowerCase();
          if (name.indexOf("personel") !== -1 && (name.endsWith(".xlsx") || name.endsWith(".xls"))) {
            var fTime = f.getLastUpdated().getTime();
            if (fTime > latestTime) {
              latestTime = fTime;
              excelFile = f;
            }
          }
        }
      } catch (err) {
        Logger.log("Klasör taraması başarısız: " + err.toString());
      }
    }
    
    if (!excelFile) {
      Logger.log("Drive üzerinde herhangi bir Personel Bilgi Excel dosyası bulunamadı.");
      return false;
    }
    
    Logger.log("En son yüklenen Personel Bilgi excel dosyası bulundu: " + excelFile.getName());
    
    // Excel dosyasını Google E-Tabloya dönüştür
    var blob = excelFile.getBlob();
    var metadata = {
      name: "[Temp_Personnel_Import]_" + excelFile.getName().replace(/\.(xlsx|xls)$/i, ""),
      mimeType: "application/vnd.google-apps.spreadsheet"
    };
    
    var boundary = "xxxxxxxxxxxxxxxxx";
    var delimiter = "\r\n--" + boundary + "\r\n";
    var closeDelimiter = "\r\n--" + boundary + "--";
    
    var multipartRequestBody = 
      delimiter +
      'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
      JSON.stringify(metadata) +
      delimiter +
      'Content-Type: ' + blob.getMimeType() + '\r\n' +
      'Content-Transfer-Encoding: base64\r\n\r\n' +
      Utilities.base64Encode(blob.getBytes()) +
      closeDelimiter;
      
    var response = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files?uploadType=multipart", {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
        "Content-Type": "multipart/related; boundary=" + boundary
      },
      payload: multipartRequestBody,
      muteHttpExceptions: true
    });
    
    var resText = response.getContentText();
    var fileInfo = JSON.parse(resText);
    if (fileInfo.error) {
      Logger.log("Excel dönüştürme hatası: " + fileInfo.error.message);
      return false;
    }
    
    var tempSsId = fileInfo.id;
    var tempSs = SpreadsheetApp.openById(tempSsId);
    var tempSheet = tempSs.getSheets()[0]; // İlk sayfayı oku
    
    var lastRow = tempSheet.getLastRow();
    var lastCol = tempSheet.getLastColumn();
    
    if (lastRow > 0 && lastCol > 0) {
      var excelValues = tempSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      
      // '5-Personel_Bilgi' sayfasını temizle ve verileri yaz
      sheet.clear();
      sheet.getRange(1, 1, excelValues.length, lastCol).setValues(excelValues);
      
      // Biçimlendirme uygula
      formatGeneralSheet(sheet);
      
      // Güncelleme tarihlerine kaydet
      var pad = function(n) { return String(n).padStart(2, '0'); };
      var now = new Date();
      var formattedDate = pad(now.getDate()) + "." + pad(now.getMonth() + 1) + "." + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes());
      recordLastUpdate(ss, "5. PERSONEL BİLGİ ÇİZELGELERİ", formattedDate);
      
      Logger.log("Drive üzerindeki Excel dosyasından '" + excelFile.getName() + "' okunan " + excelValues.length + " satır veri e-tabloya başarıyla eşitlendi.");
      
      // Geçici dosyayı sil
      try {
        UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + tempSsId, {
          method: "DELETE",
          headers: {
            "Authorization": "Bearer " + ScriptApp.getOAuthToken()
          },
          muteHttpExceptions: true
        });
      } catch (delErr) {
        Logger.log("Geçici dosya silinemedi: " + delErr.toString());
      }
      
      return true;
    }
  } catch (err) {
    Logger.log("syncPersonnelExcelToGoogleSheet hatası: " + err.toString());
  }
  return false;
}

/**
 * Teknik Yayınlar takip e-tablo sayfasının varlığından ve başlıklarından emin olur.
 */
function ensureTeknikYayinlarSheet(ss) {
  var sheetName = "TEKNİK YAYINLAR";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  if (sheet.getLastRow() === 0) {
    var headers = [
      "SIRA NO", 
      "BİRİM KODU", 
      "AİT OLDUĞU BİRİM", 
      "KATEGORİ (IPC/AMM/CMM...)", 
      "DOKÜMAN / YAYIN BAŞLIĞI", 
      "REVİZYON", 
      "BÖLÜM / ALAN", 
      "DOSYA ADI", 
      "YÜKLEME TARİHİ", 
      "DRIVE DOSYA ID", 
      "GÖRÜNTÜLEME LİNKİ", 
      "İNDİRME LİNKİ", 
      "NOTLAR / AÇIKLAMA"
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#0b3d1d")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontFamily("Calibri")
      .setFontSize(11)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.setRowHeight(1, 32);
    sheet.setFrozenRows(1);
    
    for (var i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  return sheet;
}

/**
 * Teknik yayın kaydını TEKNİK YAYINLAR tablosuna ekler veya günceller
 */
function syncTeknikYayinToSheet(ss, item) {
  try {
    var sheet = ensureTeknikYayinlarSheet(ss);
    var lastRow = sheet.getLastRow();
    var foundRow = 0;
    
    if (lastRow > 1) {
      var idValues = sheet.getRange(2, 10, lastRow - 1, 1).getValues(); // 10. sütun: DRIVE DOSYA ID
      var nameValues = sheet.getRange(2, 8, lastRow - 1, 1).getValues(); // 8. sütun: DOSYA ADI
      for (var r = 0; r < idValues.length; r++) {
        var rowFileId = String(idValues[r][0] || "").trim();
        var rowFileName = String(nameValues[r][0] || "").trim();
        if ((item.fileId && rowFileId === item.fileId) || (item.fileName && rowFileName === item.fileName)) {
          foundRow = r + 2;
          break;
        }
      }
    }
    
    var viewUrl = item.viewUrl || (item.fileId ? ("https://drive.google.com/file/d/" + item.fileId + "/preview") : "");
    var downloadUrl = item.downloadUrl || (item.fileId ? ("https://drive.google.com/uc?export=download&id=" + item.fileId) : "");
    
    var rowValues = [
      foundRow > 0 ? sheet.getRange(foundRow, 1).getValue() : (lastRow > 0 ? lastRow : 1),
      item.unitKey || "GENEL",
      item.unit || "TÜM HAVA ARAÇLARI",
      item.category || "IPC",
      item.title || item.fileName || "",
      item.revision || "Rev. 01",
      item.section || "Genel Teknik Döküman",
      item.fileName || "",
      item.uploadDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
      item.fileId || "",
      viewUrl,
      downloadUrl,
      item.notes || ""
    ];
    
    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      var nextRow = sheet.getLastRow() + 1;
      rowValues[0] = nextRow - 1; // SIRA NO
      sheet.getRange(nextRow, 1, 1, rowValues.length).setValues([rowValues]);
      
      var newRange = sheet.getRange(nextRow, 1, 1, rowValues.length);
      newRange.setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
      if ((nextRow % 2) === 0) {
        newRange.setBackground("#f8fafc");
      }
    }
  } catch (err) {
    Logger.log("syncTeknikYayinToSheet hatası: " + err.toString());
  }
}

/**
 * TEKNİK YAYINLAR tablosundan bir kaydı siler ve sıra numaralarını yeniden düzenler
 */
function deleteTeknikYayinFromSheet(ss, fileId, fileName) {
  try {
    var sheet = ensureTeknikYayinlarSheet(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;
    
    var idValues = sheet.getRange(2, 10, lastRow - 1, 1).getValues();
    var nameValues = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
    var deleted = false;
    
    for (var r = idValues.length - 1; r >= 0; r--) {
      var rowFileId = String(idValues[r][0] || "").trim();
      var rowFileName = String(nameValues[r][0] || "").trim();
      if ((fileId && rowFileId === fileId) || (fileName && rowFileName === fileName)) {
        sheet.deleteRow(r + 2);
        deleted = true;
      }
    }
    
    // Sıra numaralarını yeniden düzenle
    var newLastRow = sheet.getLastRow();
    if (newLastRow > 1) {
      for (var k = 2; k <= newLastRow; k++) {
        sheet.getRange(k, 1).setValue(k - 1);
      }
    }
    return deleted;
  } catch (err) {
    Logger.log("deleteTeknikYayinFromSheet hatası: " + err.toString());
    return false;
  }
}

/**
 * Belirtilen sayfaya transfer listesini yazar ve biçimlendirir
 */
function writeTransfersToSpecificSheet(targetSs, sheetName, transfersList, headerBgColor) {
  var sheet = targetSs.getSheetByName(sheetName);
  if (!sheet) {
    sheet = targetSs.insertSheet(sheetName);
  }
  sheet.clear();

  var headers = [
    "SIRA NO",
    "MALZEME ADI",
    "ADET",
    "TARİH",
    "İŞLEM TÜRÜ",
    "SERİAL NUMBER (S/N)",
    "KUYRUK KODU",
    "TESLİM ALAN",
    "KABUL YAPAN",
    "DEPO YERİ / LOKASYON",
    "AÇIKLAMA / NOTLAR"
  ];

  sheet.appendRow(headers);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(headerBgColor || "#0b3d1d")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontFamily("Calibri")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 30);
  sheet.setFrozenRows(1);

  if (transfersList && transfersList.length > 0) {
    var curDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
    var rowsToInsert = transfersList.map(function(t, idx) {
      var name = t.itemName || t.name || t["MALZEME ADI"] || "-";
      var qty = t.quantity !== undefined ? t.quantity : (t.qty || t["ADET"] || "1");
      var date = t.date || (t.timestamp ? String(t.timestamp).split(' ')[0] : "") || t["TARİH"] || curDateStr;
      var type = t.type || t["İŞLEM TÜRÜ"] || "TRANSFER";
      var sn = t.sn || t["SERİAL NUMBER"] || t["SERİ NO"] || "-";
      var tail = t.tailNo || t["KUYRUK KODU"] || "-";
      var op = t.operator || t["TESLİM ALAN"] || "-";
      var rec = t.receivedBy || t["KABUL YAPAN"] || "-";
      var loc = t.location || t.targetLocation || t.sourceLocation || t["DEPO YERİ"] || "DEPO";
      var notes = t.notes || t["AÇIKLAMA"] || "";

      return [
        idx + 1,
        name,
        qty,
        date,
        type,
        sn,
        tail,
        op,
        rec,
        loc,
        notes
      ];
    });

    sheet.getRange(2, 1, rowsToInsert.length, headers.length).setValues(rowsToInsert);
    var dataRange = sheet.getRange(2, 1, rowsToInsert.length, headers.length);
    dataRange.setFontFamily("Calibri").setFontSize(10).setVerticalAlignment("middle");
    
    for (var r = 0; r < rowsToInsert.length; r++) {
      if (r % 2 === 1) {
        sheet.getRange(r + 2, 1, 1, headers.length).setBackground("#f8fafc");
      }
    }
  }

  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
    var w = sheet.getColumnWidth(c);
    if (c === 2) { if (w < 200) sheet.setColumnWidth(c, 220); }
    else if (c === 10 || c === 11) { if (w < 180) sheet.setColumnWidth(c, 200); }
    else { if (w < 90) sheet.setColumnWidth(c, 100); }
  }
}

/**
 * Transfer kaydının ait olduğu hava aracı/depo kategorisini belirler
 */
function getDepoTransferUnitKey(t) {
  var fullText = ((t.unit || "") + " " + (t.depo || "") + " " + (t.tailNo || "") + " " + (t.itemName || "") + " " + (t.location || "") + " " + (t.sourceLocation || "") + " " + (t.targetLocation || "") + " " + (t.notes || "")).toLowerCase()
    .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u').replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c');
  
  if (fullText.indexOf("at802") !== -1 || fullText.indexOf("at-802") !== -1 || fullText.indexOf("or-20") !== -1) return "at802";
  if (fullText.indexOf("bell") !== -1 || fullText.indexOf("429") !== -1 || fullText.indexOf("tc-zog") !== -1 || fullText.indexOf("tc-zoh") !== -1 || fullText.indexOf("tc-zoi") !== -1) return "bell429";
  if (fullText.indexOf("t70") !== -1 || fullText.indexOf("t-70") !== -1 || fullText.indexOf("bumbi") !== -1 || fullText.indexOf("helitak") !== -1 || fullText.indexOf("sikorsky") !== -1) return "t70";
  if (fullText.indexOf("b360") !== -1 || fullText.indexOf("b-360") !== -1 || fullText.indexOf("king air") !== -1 || fullText.indexOf("tc-ogm") !== -1) return "b360";
  if (fullText.indexOf("c650") !== -1 || fullText.indexOf("c-650") !== -1 || fullText.indexOf("citation") !== -1 || fullText.indexOf("tc-cya") !== -1 || fullText.indexOf("tc-cyb") !== -1) return "c650";
  if (fullText.indexOf("hangar") !== -1) return "hangar";
  if (fullText.indexOf("kara") !== -1) return "kara_araclari";
  return "at802";
}

/**
 * TRANSFER GEÇMİŞİ sayfasını oluşturur veya mevcut olanı döndürür
 */
function ensureTransferGecmisiSheet(ss) {
  var sheetName = "TRANSFER GEÇMİŞİ";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [
      "SIRA NO",
      "MALZEME ADI",
      "ADET",
      "TARİH",
      "İŞLEM TÜRÜ",
      "SERİAL NUMBER (S/N)",
      "KUYRUK KODU",
      "TESLİM ALAN",
      "KABUL YAPAN",
      "DEPO YERİ / LOKASYON",
      "AÇIKLAMA / NOTLAR"
    ];
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#0b3d1d")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontFamily("Calibri")
      .setFontSize(11)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.setRowHeight(1, 30);
    sheet.setFrozenRows(1);
    
    for (var c = 1; c <= headers.length; c++) {
      sheet.autoResizeColumn(c);
      var w = sheet.getColumnWidth(c);
      if (c === 2) { if (w < 200) sheet.setColumnWidth(c, 220); }
      else if (c === 10 || c === 11) { if (w < 180) sheet.setColumnWidth(c, 200); }
      else { if (w < 90) sheet.setColumnWidth(c, 100); }
    }
  }
  return sheet;
}



