// Web uygulamasını başlatır
function doGet() {
  return HtmlService.createHtmlOutputFromFile('İsgBelgeY');
}

// Belirtilen klasördeki dosyaları listeler
function getDriveFiles(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const fileList = [];

    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        name: file.getName(),
        id: file.getId(),
        mimeType: file.getMimeType()
      });
    }
    return fileList;
  } catch (e) {
    Logger.log("getDriveFiles hatası: " + e.message);
    return []; // Hata durumunda boş liste döndür
  }
}
/**
 * Belirli bir dosyanın adını ve ID'sini döndürür.
 * @param {string} fileId Bilgisi alınacak dosyanın ID'si.
 * @returns {Object} Dosyanın adı ve ID'sini içeren obje.
 */
function getFileInfo(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return { name: file.getName(), id: file.getId() };
  } catch (e) {
    Logger.log("getFileInfo hatası: " + e.message);
    return null; // Hata durumunda null döndür
  }
}
// Seçilen dosyaları PDF olarak kaydeder (Mevcut Fonksiyon - Korundu)
function saveAsPdf(fileIds) {
  const targetFolderId = '12TqiA2MttEhq9DRXGCOa8s7zWypVoy7O'; // Hedef klasör ID (GMAİL PDF Klasörü)
  const folder = DriveApp.getFolderById(targetFolderId);
  const result = [];

  fileIds.forEach(function(fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      const mimeType = file.getMimeType();

      if ([MimeType.GOOGLE_DOCS, MimeType.GOOGLE_SHEETS, MimeType.GOOGLE_SLIDES].includes(mimeType)) {
        const pdfBlob = file.getAs('application/pdf');
        folder.createFile(pdfBlob).setName(file.getName() + ".pdf");
        result.push('PDF olarak kaydedildi: ' + file.getName());
      } else if (mimeType === MimeType.PDF) {
        // Eğer zaten PDF ise kopyalıyoruz
        folder.createFile(file.getBlob()).setName(file.getName());
        result.push('PDF kopyalandı: ' + file.getName());
      } else {
        result.push('PDF olarak kaydedilemedi: ' + file.getName() + ' (Desteklenmeyen dosya türü)');
      }
    } catch (e) {
      result.push(`PDF kaydetme hatası (${fileId}): ${e.message}`);
    }
  });

  return result.join('\n');
}

/**
 * Belirtilen dosyaları bir kaynak klasörden başka bir hedef klasöre taşır.
 * Bu fonksiyon hem EK-2 klasörü hem de SmartMove (firma klasörüne taşıma) için kullanılabilir.
 * @param {string[]} fileIds Taşınacak dosya ID'lerinin dizisi.
 * @param {string} sourceFolderId Dosyaların bulunduğu kaynak klasörün ID'si.
 * @param {string} targetFolderId Dosyaların taşınacağı hedef klasörün ID'si.
 * @returns {string} İşlem sonucu mesajı.
 */
function moveDriveFiles(fileIds, sourceFolderId, targetFolderId) {
  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  const results = [];

  fileIds.forEach(fileId => {
    try {
      const file = DriveApp.getFileById(fileId);
      // Dosyayı kaynaktan kaldır, hedefe ekle. Bu taşıma işlemini simüle eder.
      sourceFolder.removeFile(file); // Önce kaynak klasörden kaldır
      targetFolder.addFile(file);   // Sonra hedef klasöre ekle
      results.push(`Taşındı: ${file.getName()}`);
    } catch (e) {
      results.push(`Taşıma Başarısız (${fileId}): ${e.message}`);
      Logger.log(`Dosya taşıma hatası ${fileId}: ${e.message}`);
    }
  });
  return results.join('\n');
}


// Seçilen dosyaları siler
function deleteFiles(fileIds) {
  var result = [];
  fileIds.forEach(function(fileId) {
    try {
      var file = DriveApp.getFileById(fileId);
      file.setTrashed(true); // Çöp kutusuna gönderiyor
      result.push('Silindi: ' + file.getName());
    } catch (e) {
      result.push('Silinemedi: ' + fileId + ' - ' + e.message);
    }
  });
  return result.join('\n');
}

// Seçilen PDF dosyalarıyla Gmail taslağı oluşturur
function createGmailDraft(fileIds) {
  const files = fileIds.map(fileId => DriveApp.getFileById(fileId));
  const subject = "İş Sağlığı ve Güvenliği Belgeleri";
  const plainBody = "Talep etmiş olduğunuz İSG belgeleri ektedir. İyi çalışmalar dileriz."; // Yedek metin

  // FLIST sayfasını al
  const sheet = SpreadsheetApp.getActive().getSheetByName('FLIST');
  const data = sheet.getDataRange().getValues(); // Tüm verileri çek

  let toField = "";

  if (files.length > 0) {
    const firstFileName = files[0].getName();

    // FLIST içindeki başlık satırını bul
    const headerRow = data[0];
    const emailIndex = headerRow.indexOf('E_POSTA'); // 'E-Posta' sütununun indeksini dinamik olarak bul
    const unvanIndex = headerRow.indexOf('UNVANI'); // 'UNVANI' sütununun indeksini dinamik olarak bul
    
    // Eğer gerekli sütunlar bulunamazsa hata döndür
    if (emailIndex === -1 || unvanIndex === -1) {
      return '❌ Hata: "FLIST" sayfasında "E-Posta" veya "UNVANI" sütunu bulunamadı.';
    }

    // Dosya adının ilk 7 karakterini al, eğer dosya adı en az 7 karakterliyse
    const matchKey = firstFileName.length >= 7 ? firstFileName.substring(0, 7) : firstFileName;

    // FLIST içinde eşleşen satırı bul
    for (let i = 1; i < data.length; i++) {
      const unvan = data[i][unvanIndex];
      if (unvan && unvan.toString().startsWith(matchKey)) {
        const emails = data[i][emailIndex];
        if (emails) {
          toField = emails.split(',').map(email => email.trim()).join(',');
        }
        break; // İlk eşleşmeden sonra döngüyü durdur
      }
    }
  }

  const htmlBody = `
    <p style="font-size: 15px;">Sayın Yetkili,</p>
    <p>İSG hizmetleri kapsamında aşağıda belirtilen belgeleri ekte bilgilerinize sunuyoruz:</p>
    <ul style="line-height: 1.6;">
      <li><b>EK-2 Sağlık Raporları:</b> <span style="color: #e63946;">"Çalışanın Adı Soyadı"</span> bölümü doldurulup imzalanmalıdır.</li>
      <li><b>Eğitime Katılım Belgesi:</b> İmzalandıktan sonra <span style="color: #457b9d;">senaosgb1@gmail.com</span> adresine gönderilmelidir.</li>
      <li><b>Eğitim Sertifikaları:</b> <span style="color: #2a9d8f;">İşveren</span> tarafından imzalanmalıdır.</li>
      <li><b>KKD Zimmet Tutanakları:</b> Her çalışan için çıktı alınmalı, gerekli bilgiler doldurularak <span style="color: #e76f51;">hem çalışan hem de işveren</span> tarafından imzalanmalıdır.</li>
      <li><b>Eğitime Katılım Formu:</b> Tüm imzalar tamamlandıktan sonra e-posta yoluyla gönderiniz.</li>
      <li>
        <b>Geçerlilik Süresi:</b><br>
        - Çok Tehlikeli: <span style="color: #e63946;"><b>1 yıl</b></span><br>
        - Tehlikeli: <span style="color: #f4a261;"><b>2 yıl</b></span><br>
        - Az Tehlikeli: <span style="color: #2a9d8f;"><b>3 yıl</b></span>
      </li>
    </ul>
    <p style="font-size: 13px;">Bilgilerinize sunar, iyi çalışmalar dileriz.</p>
    <p style="font-size: 12px; color: gray;">SENA OSGB</p>
  `;

  try {
    GmailApp.createDraft(toField, subject, plainBody, {
      attachments: files.map(file => file.getAs(MimeType.PDF)), // Sadece PDF formatında ekler
      htmlBody: htmlBody
    });
    return '✅ Gmail taslağı oluşturuldu. Taslaklar sayfasına yönlendiriliyorsunuz.';
  } catch (e) {
    Logger.log("Gmail taslağı oluşturma hatası: " + e.message);
    return '❌ Gmail taslağı oluşturulurken bir hata oluştu: ' + e.message;
  }
}

/**
 * Belirtilen üst klasördeki tüm alt klasörleri (firmaları) listeler.
 * @param {string} parentFolderId Firma klasörlerinin bulunduğu ana klasörün ID'si.
 * @returns {Array<Object>} Firma klasörlerinin adlarını ve ID'lerini içeren dizi.
 */
function getCompanyFolders(parentFolderId) {
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const subFolders = parentFolder.getFolders();
    const folderList = [];
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      folderList.push({
        name: folder.getName(),
        id: folder.getId()
      });
    }
    // Alfabetik sıralama ekleyelim
    folderList.sort((a, b) => a.name.localeCompare(b.name));
    return folderList;
  } catch (e) {
    Logger.log("getCompanyFolders hatası: " + e.message);
    return [];
  }
}

/**
 * Yeni bir firma klasörü oluşturur.
 * @param {string} folderName Oluşturulacak yeni klasörün adı.
 * @param {string} parentFolderId Yeni klasörün oluşturulacağı ana klasörün ID'si.
 * @returns {Object} Oluşturulan klasörün adı ve ID'si.
 */
function createCompanyFolder(folderName, parentFolderId) {
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const newFolder = parentFolder.createFolder(folderName);
    return { name: newFolder.getName(), id: newFolder.getId() };
  } catch (e) {
    Logger.log("createCompanyFolder hatası: " + e.message);
    throw new Error("Klasör oluşturulamadı: " + e.message); // Hata durumunda istemciye bilgi gönder
  }
}

/**
 * Seçilen dosyaları, belirtilen firma klasörünün altına
 * "Adı Soyadı_Taşındığı Tarih" formatında alt klasörler oluşturarak taşır.
 * Dosya adından "Kişi Adı Soyadı" kısmını (ilk alt çizgi ile ikinci alt çizgi arasında kalan) çıkarır.
 * Her benzersiz kişi için tek bir alt klasör oluşturur.
 *
 * @param {string[]} fileIds Taşınacak dosya ID'lerinin dizisi.
 * @param {string} targetCompanyFolderId Dosyaların taşınacağı firma klasörünün ID'si.
 * @param {string} currentSourceFolderId Dosyaların mevcut kaynak klasörü (Gmail PDF klasörü).
 * @returns {Object} İşlem sonucu mesajı ve hedef klasör ID'si.
 */
function smartMoveFiles(fileIds, targetCompanyFolderId, currentSourceFolderId) {
  const targetCompanyFolder = DriveApp.getFolderById(targetCompanyFolderId);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy"); // Script'in zaman dilimi kullanılır
  const results = [];
  const currentSourceFolder = DriveApp.getFolderById(currentSourceFolderId);

  const personFolders = {}; // { "Kişi Adı Soyadı": FolderObject }

  // 1. Adım: Tüm dosyaları tek tek işleyip kişi adlarını çıkar ve kişi bazında grupla
  const filesToProcess = [];
  fileIds.forEach(fileId => {
    try {
      const file = DriveApp.getFileById(fileId);
      const fileName = file.getName();

      let personName = "Bilinmeyen Kişi";
      
      // Dosya adını alt çizgilere göre bölelim
      const parts = fileName.split('_');

      // Beklenen format: FIRMA_KISIASISOYADI_BELGETIPI_...
      // Kişi adı her zaman parts[1] olacaktır (0 tabanlı indeksleme)
      if (parts.length > 1) {
          personName = parts[1].trim(); // İkinci parçayı kişi adı olarak al
      }
      
      // Klasör adı için kişinin adını soyadını düzenle:
      // Türkçe karakterleri İngilizce karşılıklarına çevir ve boşlukları alt çizgiye çevir.
      const folderPersonName = personName
                                .replace(/Ğ/g, 'G').replace(/ğ/g, 'g')
                                .replace(/İ/g, 'I').replace(/ı/g, 'i')
                                .replace(/Ş/g, 'S').replace(/ş/g, 's')
                                .replace(/Ö/g, 'O').replace(/ö/g, 'o')
                                .replace(/Ü/g, 'U').replace(/ü/g, 'u')
                                .replace(/Ç/g, 'C').replace(/ç/g, 'c')
                                .replace(/ /g, '_') // Boşlukları alt çizgiye çevir
                                .replace(/[^a-zA-Z0-9_]/g, ''); // Harf, rakam ve alt çizgi dışındaki karakterleri kaldır


      const subFolderName = `${folderPersonName}_${today}`;


      filesToProcess.push({
        file: file,
        fileName: fileName,
        personName: personName, // Ham kişi adı (eşleşme için)
        subFolderName: subFolderName, // Oluşturulacak alt klasör adı (Drive için)
      });

    } catch (e) {
      results.push(`Dosya bilgisi alınamadı (${fileId}): ${e.message}`);
      Logger.log(`Dosya bilgisi alma hatası ${fileId}: ${e.message}`);
    }
  });

  // 2. Adım: Benzersiz kişiler için alt klasörleri oluştur (veya bul)
  for (const item of filesToProcess) {
    const { personName, subFolderName } = item;

    if (!personFolders[personName]) { // Eğer bu kişi için klasör henüz oluşturulmadıysa/bulunmadıysa
      let subFolder;
      try {
        // Hedef firma klasörü altında bu isimde bir alt klasör var mı kontrol et
        subFolder = targetCompanyFolder.getFoldersByName(subFolderName).next();
      } catch (e) {
        // Yoksa oluştur
        subFolder = targetCompanyFolder.createFolder(subFolderName);
        results.push(`Klasör oluşturuldu: ${targetCompanyFolder.getName()}/${subFolderName}`);
      }
      personFolders[personName] = subFolder; // Klasörü haritaya kaydet
    }
  }

  // 3. Adım: Dosyaları ilgili alt klasörlere taşı
  for (const item of filesToProcess) {
    const { file, fileName, personName } = item;
    const targetSubFolder = personFolders[personName]; // Doğru kişinin klasörünü al

    if (targetSubFolder) {
      try {
        // Dosyanın mevcut ebeveyn klasöründen kaldır
        currentSourceFolder.removeFile(file);
        
        // Dosyayı yeni alt klasöre ekle
        targetSubFolder.addFile(file);
        results.push(`Taşındı: ${fileName} -> ${targetSubFolder.getName()}`);
      } catch (e) {
        results.push(`Taşıma Başarısız (${fileName}): ${e.message}`);
        Logger.log(`SmartMove Dosya taşıma hatası ${fileName}: ${e.message}`);
      }
    } else {
      results.push(`Hata: ${fileName} için hedef alt klasör bulunamadı/oluşturulamadı.`);
    }
  }
  
  return { message: results.join('\n'), targetFolderId: targetCompanyFolderId };
}