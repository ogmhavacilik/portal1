/**
 * ==============================================================================
 * HAVA ARAÇLARI BAKIM TEKNİK ŞUBE MÜDÜRLÜĞÜ - E-TABLO SENKRONİZASYON MOTORU (V6)
 * ==============================================================================
 * Dosya Adı: komut.gs
 * Google Drive Hedef Klasör ID (PDF): 1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0
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
      var unitName = getUnitTitleById(formId, month);
      
      recordLastUpdate(ss, unitName, formattedDate);
      response.message = "'" + unitName + "' planlaması için güncelleme tarihi başarıyla e-tabloya kaydedildi.";
      response.uploadDate = formattedDate;

    } else if (action === "uploadPdfToDrive") {
      var fileName = postData.fileName;
      var base64Data = postData.base64Data;
      var formId = postData.formId ? Number(postData.formId) : null;
      var month = postData.month || null;
      
      var folderId = "1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0";
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
      
      var decoded = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(decoded, "application/pdf", fileName);
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
      
      var unitName = getUnitTitleById(derivedFormId, derivedMonth);
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
    
    if (action === "getSheets") {
      response.sheets = ss.getSheets().map(function(s) {
        return { name: s.getName(), id: s.getSheetId() };
      });
      response.message = "Sayfalar başarıyla listelendi.";

    } else if (action === "listPdfsFromDrive") {
      var list = [];
      try {
        var folder = DriveApp.getFolderById("1_fIGvuPVpC9N5on1irOfGG8OsD1KSXD0");
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
      var sheet = getSheetWithFallback(ss, sheetName);
      var rows = [];
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      
      if (lastRow > 0 && lastCol > 0) {
        var allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
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
