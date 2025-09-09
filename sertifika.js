var templateDocIds = [
  // LÜTFEN BU KİMLİKLERİ KENDİ GOOGLE DOKÜMANLARINIZIN KİMLİKLERİYLE DEĞİŞTİRİN!
  '13Eiab7e-lN6bDL9fCb0RtgyPUM-kgXpHbMWmd80AwQQ', // 1: temel_yc_8_8
  '1bhGfCqR4rq_Cu_k-9EYGnVfCICq_xvnq6tnSSXN6Q94', // 2: temel8
  '1THEeZDNX7WLrOAJO5kskhTfNGxgmGy2O7rvixUR-I4Y', // 3: temel8.4
  '1Z9x556YL8U57c57EW9K_IRhphcMHs2i7uwNJ6gRrcek', // 4: yc
  '1vEr-FOfncotkkDfnAjQJomPcwUSxFBgVeK3Ku04Z6hA', // 5: katilan_listesi
  '1nOHthz0kXh-aLcnQQ5cU_m6F8BRlAJrCUC1XoED0ato', // 6: kkd
  '1Fo8cQPSxLAYIY4fOAUakYCSfz2C-Td0rC_xaqvRIoak', // 7: yüksekte_çalışma_talimatı
  '1-HOvry8s1KnWRHktD4QRGceQqiaDKM4VDWwspurvzmY', // 8: temel_8_8
  '1CXQGi5pQY3-y3vW_dQtvVXhyk1PLm1gVuf_07KprHz4', // 9: sınav_soruları
  '18fvcE7N4pm-rmw3ozBP6n7T8M_KjXX0j-_l6IIc-noA', // 10: konulu_sertifika
];

function showCertificateDialog() {
  const html = HtmlService.createHtmlOutputFromFile('index.html')
    .setWidth(700)
    .setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, 'Eğitim Sertifikası Oluştur');
}

/**
 * İmza görsellerini Google Drive'dan alıp belgeye ekler
 */
function insertImageFromUrl(body, placeholderText, imageUrl) {
  try {
    if (!imageUrl || imageUrl === '') {
      // URL yoksa placeholder'ı boş bırak
      body.replaceText(escapeRegExp(placeholderText), '');
      return;
    }
    
    // Google Drive URL'sinden dosya ID'sini çıkar
    let fileId = null;
    
    // Farklı Google Drive URL formatlarını kontrol et
    if (imageUrl.includes('drive.google.com')) {
      // Format 1: https://drive.google.com/file/d/FILE_ID/view
      const match1 = imageUrl.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
      if (match1) {
        fileId = match1[1];
      }
      // Format 2: https://drive.google.com/open?id=FILE_ID
      const match2 = imageUrl.match(/[?&]id=([a-zA-Z0-9-_]+)/);
      if (!fileId && match2) {
        fileId = match2[1];
      }
    } else if (imageUrl.match(/^[a-zA-Z0-9-_]+$/)) {
      // Direkt dosya ID'si verilmişse
      fileId = imageUrl;
    }
    
    if (!fileId) {
      Logger.log('Geçersiz imza URL\'si: ' + imageUrl);
      body.replaceText(escapeRegExp(placeholderText), '');
      return;
    }
    
    // Görseli al
    const imageBlob = DriveApp.getFileById(fileId).getBlob();
    
    // Placeholder'ı bul
    const foundElement = body.findText(escapeRegExp(placeholderText));
    
    if (foundElement) {
      const paragraph = foundElement.getElement().getParent();
      
      // Görseli ekle
      const image = paragraph.asParagraph().insertInlineImage(
        paragraph.asParagraph().getChildIndex(foundElement.getElement()), 
        imageBlob
      );
      
      // Görsel boyutunu ayarla (örnek: 100x50 piksel)
      image.setWidth(100);
      image.setHeight(50);
      
      // Placeholder metnini kaldır
      foundElement.getElement().asText().setText('');
    }
    
  } catch (error) {
    Logger.log('İmza görseli eklenirken hata: ' + error.toString());
    // Hata durumunda placeholder'ı boş bırak
    body.replaceText(escapeRegExp(placeholderText), '');
  }
}

/**
 * Girilen satır numarasını ayrıştırır.
 */
function parseRowNumbers(inputStr) {
  if (!inputStr) return [];
  const parts = inputStr.split(',').map(part => part.trim());
  const rows = new Set();

  parts.forEach(part => {
    if (part.includes('-')) {
      const [start, end] = part.split('-').map(Number);
      if (start <= end) {
        for (let i = start; i <= end; i++) {
          rows.add(i);
        }
      }
    } else {
      const num = Number(part);
      if (!isNaN(num)) {
        rows.add(num);
      }
    }
  });
  return Array.from(rows).sort((a, b) => a - b);
}

/**
 * "EGITIM" sayfasının ilk satırındaki sütun başlıklarını döndürür.
 */
function getTrainingTopicTitles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EGITIM');
  if (!sheet) {
    Logger.log("EGITIM sayfası bulunamadı.");
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    return [];
  }
  // İlk satır, başlıkları içerir
  return data[0].filter(title => title && typeof title === 'string' && title.trim() !== '');
}

/**
 * Seçilen konu başlıklarına göre içerikleri getirir.
 */
function getTrainingTopicsContent(selectedTitles) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EGITIM');
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "";
  const headers = data[0];
  const formattedContent = [];

  selectedTitles.forEach(title => {
    const colIndex = headers.indexOf(title);
    if (colIndex !== -1) {
      const topicContent = [];
      topicContent.push(title); // Başlığı ekle
      
      const columnContent = data.slice(1).map(row => row[colIndex]).filter(cell => cell);
      topicContent.push(...columnContent); // Alt maddeleri ekle
      
      if (topicContent.length > 1) {
        formattedContent.push(topicContent.join('\n'));
      }
    }
  });
  
  return formattedContent.join('\n\n');
}

/**
 * Girilen satır numaralarına göre personel listesini döndürür.
 */
function getPersonnelByRows(rowNumbers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EK-2 BİLGİLER');
  if (!sheet) {
    throw new Error("EK-2 BİLGİLER sayfası bulunamadı!");
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const personnelList = [];
  rowNumbers.forEach(rowNum => {
    if (rowNum >= 2 && rowNum <= data.length) {
      const row = data[rowNum - 1];
      const personData = {};
      headers.forEach((header, index) => {
        // TC Kimlik No ve Mesleği/Meslek Dalı'nı string olarak zorla
        if (header === 'TC Kimlik No' || header === 'Mesleği/Meslek Dalı' || header === 'Adres') {
          personData[header] = String(row[index]);
        } else {
          personData[header] = row[index];
        }
      });
      personnelList.push(personData);
    }
  });
  return personnelList;
}

/**
* Firmanın tehlike sınıfına göre yıl değerini döndürür.
* Az Tehlikeli: 3 yıl, Tehlikeli: 2 yıl, Çok Tehlikeli: 1 yıl
*/
function getHazardYearsForFirm(firmName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flistSheet = ss.getSheetByName('FLIST');
  if (!flistSheet) {
    Logger.log("FLIST sayfası bulunamadı.");
    return 1;
  }
  const flistData = flistSheet.getDataRange().getValues();
  const flistHeaders = flistData[0];
  const idxTehlike = flistHeaders.indexOf('TEHLİKE_SINIFI');
  if (idxTehlike === -1) {
    Logger.log("FLIST sayfasında 'TEHLİKE_SINIFI' sütunu bulunamadı.");
    return 1;
  }
  // Firma adını 2. sütunda (indeks 1) aramak için güncellendi
  const firmMatch = flistData.find(row => normalizeString(row[1]) === normalizeString(firmName));
  if (firmMatch) {
    const tehlikeSinifi = (firmMatch[idxTehlike] || '').toString().toLowerCase();
    if (tehlikeSinifi.includes('çok')) return 1;
    if (tehlikeSinifi.includes('az')) return 3;
    if (tehlikeSinifi.includes('tehlikeli')) return 2;
  }
  return 1;
}

function processFormWithRowNumbers(rowNumbersStr, templates, firm, date1Str, date3Str, date5Str, trainingHours1, trainingHours2, trainingHours3, selectedTopicNames) {
  function addYears(date, years) {
    if (!date) return null;
    const d = new Date(date);
    d.setFullYear(d.getFullYear() + years);
    return d;
  }
  function fmt(d){ 
    return d ? formatDateToDDMMYYYY(d) : '';
  }

  const rowNumbers = parseRowNumbers(rowNumbersStr);
  const selectedPersonnel = getPersonnelByRows(rowNumbers);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flistSheet = ss.getSheetByName('FLIST');
  const sayacSheet = ss.getSheetByName('SertifikaSayac');
  const flistData = flistSheet.getDataRange().getValues();
  const flistHeaders = flistData[0];
  const sayacCell = sayacSheet.getRange('A1');
  const targetFolder = DriveApp.getFolderById('1yrwxxbb7WA2MxQ0UoZ5LN9jJahOaD5Uq');
  const year = new Date().getFullYear();
  let sertifikaCounter = parseInt(sayacCell.getValue(), 10) || 0;

  const date1 = date1Str ? parseDateFromDDMMYYYY(date1Str) : null;
  const date3 = date3Str ? parseDateFromDDMMYYYY(date3Str) : null;
  const date5 = date5Str ? parseDateFromDDMMYYYY(date5Str) : null;

  const trainingTopicsContent = getTrainingTopicsContent(selectedTopicNames);
  const topicsBlocks = trainingTopicsContent ? trainingTopicsContent.split('\n\n') : [];
  const topicsLeft = topicsBlocks.slice(0, 3).join('\n\n');
  const topicsRight = topicsBlocks.slice(3).join('\n\n');
  
  let flistMatch = null;
  // Firma adını 2. sütunda (indeks 1) aramak için güncellendi
  for (let i = 1; i < flistData.length; i++) {
    if (flistData[i][1] === firm) {
      flistMatch = flistData[i];
      break;
    }
  }

  const idxTehlike = flistHeaders.indexOf('TEHLİKE_SINIFI');
  const yil = getHazardYearsForFirm(firm);
  
  const date2 = addYears(date1, yil);
  const date4 = addYears(date3, yil);
  const date6 = addYears(date5, yil);
  const s1 = parseInt(trainingHours1, 10) || 0;
  const s2 = parseInt(trainingHours2, 10) || 0;
  const s3 = parseInt(trainingHours3, 10) || 0;
  const toplamSaatStr = s3 > 0 ? `${s1 + s2}+${s3} Saat` : `${s1 + s2} Saat`;
  const failedCreations = [];
  let totalDocsCreated = 0;

  // ---- Katılımcı Listesi (şablon 5) ----
  const KATILAN_LISTESI_INDEX = 5;
  if (templates.includes(KATILAN_LISTESI_INDEX)) {
    try {
      const templateDocId = templateDocIds[KATILAN_LISTESI_INDEX - 1];
      const tarihlerStr = [fmt(date1), fmt(date3), fmt(date5)].filter(Boolean).join(' ');
      const newFileName = `${firm.substring(0,15)} - Katılımcı Listesi${tarihlerStr ? ' - ' + tarihlerStr : ''}`;
      const newFile = DriveApp.getFileById(templateDocId).makeCopy(newFileName, targetFolder);
      const doc = DocumentApp.openById(newFile.getId());
      const body = doc.getBody();
      
      body.replaceText('{{FIRMA}}', firm || '');
      body.replaceText('{{TARIH1}}', fmt(date1));
      body.replaceText('{{TARIH2}}', fmt(date2));
      body.replaceText('{{TARIH3}}', fmt(date3));
      body.replaceText('{{TARIH4}}', fmt(date4));
      body.replaceText('{{TARIH5}}', fmt(date5));
      body.replaceText('{{TARIH6}}', fmt(date6));
      body.replaceText('{{SAAT1}}', trainingHours1 || '');
      body.replaceText('{{SAAT2}}', trainingHours2 || '');
      body.replaceText('{{SAAT3}}', trainingHours3 || '');
      body.replaceText('{{TOPLAM_SAAT}}', toplamSaatStr);
      
      // FLIST verilerini yerleştir (İmzalar hariç)
      flistHeaders.forEach((h, idx) => {
        if (h !== 'IGU_IMZA' && h !== 'IH_IMZA') {
          body.replaceText(escapeRegExp(`{{${h}}}`), flistMatch[idx] || '');
        }
      });
      
      // İmza görsellerini ekle
      const iguImzaIndex = flistHeaders.indexOf('IGU_IMZA');
      const ihImzaIndex = flistHeaders.indexOf('IH_IMZA');
      
      if (iguImzaIndex !== -1 && flistMatch[iguImzaIndex]) {
        insertImageFromUrl(body, '{{IGU_IMZA}}', flistMatch[iguImzaIndex]);
      } else {
        body.replaceText('{{IGU_IMZA}}', '');
      }
      
      if (ihImzaIndex !== -1 && flistMatch[ihImzaIndex]) {
        insertImageFromUrl(body, '{{IH_IMZA}}', flistMatch[ihImzaIndex]);
      } else {
        body.replaceText('{{IH_IMZA}}', '');
      }
      
      body.replaceText('{{EGITIM_KONULARI_SOL}}', topicsLeft || '');
      body.replaceText('{{EGITIM_KONULARI_SAG}}', topicsRight || '');
      body.replaceText('{{EGITIM_KONULARI}}', '');
      
      // Tabloyu ekleyeceğimiz yer tutucuyu bul
      const tablePlaceholder = body.findText('{{KATILIMCI_LISTESI}}');
      if (tablePlaceholder) {
        const tableParagraph = tablePlaceholder.getElement().getParent();
        const table = body.insertTable(body.getChildIndex(tableParagraph) + 1);
        table.setBorderWidth(1);
        
        const headerRow = table.appendTableRow();
        ['Sıra No','TC Kimlik No','Adı Soyadı','Mesleği/Meslek Dalı','İmza'].forEach(h=>{
          headerRow.appendTableCell(h);
        });
        // Sütun genişliklerini ayarla
        table.setColumnWidth(0, 50);
        table.setColumnWidth(1, 90);
        table.setColumnWidth(2, 150);
        table.setColumnWidth(3, 150);
        table.setColumnWidth(4, 100);

        let sira = 1;
        selectedPersonnel.forEach(p => {
          const r = table.appendTableRow();
          r.appendTableCell(String(sira++));
          r.appendTableCell(p['TC Kimlik No'] || '');
          r.appendTableCell(`${p['Adı'] || ''} ${p['Soyadı'] || ''}`.trim());
          r.appendTableCell(p['Mesleği/Meslek Dalı'] || '');
          r.appendTableCell('');
          r.setMinimumHeight(28);
        });

        // Yer tutucuyu ve paragrafı kaldır
        tableParagraph.removeFromParent();
      }

      doc.saveAndClose();
      totalDocsCreated++;
    } catch (e) {
      failedCreations.push({ template: 'Katılımcı Listesi', error: e.message });
    }
  }
  
  // ---- Diğer Şablonlar ----
  const otherTemplates = templates.filter(t => t !== KATILAN_LISTESI_INDEX);
  selectedPersonnel.forEach(person => {
    const data = { ...person };
    data['Tarih1'] = fmt(date1);
    data['Tarih2'] = fmt(date2);
    data['Tarih3'] = fmt(date3);
    data['Tarih4'] = fmt(date4);
    data['Tarih5'] = fmt(date5);
    data['Tarih6'] = fmt(date6);

    sertifikaCounter++;
    const sertifikaNo = `${year}-${Utilities.formatString("%05d", sertifikaCounter)}`;
    data['SertifikaNo'] = sertifikaNo;

    otherTemplates.forEach(templateIndex => {
      try {
        const templateDocId = templateDocIds[templateIndex - 1];
        const templateName = getTemplateName(templateDocId);
        const fileName = `${firm.substring(0,15)}_${data['Adı'] || ''} ${data['Soyadı'] || ''}_${templateName}_${sertifikaNo}`;
        const newDoc = DriveApp.getFileById(templateDocId).makeCopy(fileName, targetFolder);
        const doc = DocumentApp.openById(newDoc.getId());
        const body = doc.getBody();

        // Personel verilerini yerleştir
        for (let k in data) {
          body.replaceText(escapeRegExp(`{{${k}}}`), data[k] || '');
        }
        
        body.replaceText('{{EğitimSaati1}}', trainingHours1 || '');
        body.replaceText('{{EğitimSaati2}}', trainingHours2 || '');
        body.replaceText('{{EğitimSaati3}}', trainingHours3 || '');
        body.replaceText('{{TOPLAM_SAAT}}', toplamSaatStr);
        
        // FLIST verilerini yerleştir (İmzalar hariç)
        flistHeaders.forEach((h, idx) => {
          if (h !== 'IGU_IMZA' && h !== 'IH_IMZA') {
            body.replaceText(escapeRegExp(`{{${h}}}`), flistMatch[idx] || '');
          }
        });
        
        // İmza görsellerini ekle
        const iguImzaIndex = flistHeaders.indexOf('IGU_IMZA');
        const ihImzaIndex = flistHeaders.indexOf('IH_IMZA');
        
        if (iguImzaIndex !== -1 && flistMatch[iguImzaIndex]) {
          insertImageFromUrl(body, '{{IGU_IMZA}}', flistMatch[iguImzaIndex]);
        } else {
          body.replaceText('{{IGU_IMZA}}', '');
        }
        
        if (ihImzaIndex !== -1 && flistMatch[ihImzaIndex]) {
          insertImageFromUrl(body, '{{IH_IMZA}}', flistMatch[ihImzaIndex]);
        } else {
          body.replaceText('{{IH_IMZA}}', '');
        }
        
        body.replaceText('{{EGITIM_KONULARI_SOL}}', topicsLeft || '');
        body.replaceText('{{EGITIM_KONULARI_SAG}}', topicsRight || '');
        body.replaceText('{{EGITIM_KONULARI}}', trainingTopicsContent || '');

        doc.saveAndClose();
        totalDocsCreated++;
      } catch (e) {
        failedCreations.push({ person: `${person['Adı']} ${person['Soyadı']}`, template: templateIndex, error: e.message });
      }
    });
  });

  sayacCell.setValue(sertifikaCounter);

  return { success: failedCreations.length === 0, failedCreations, totalDocsCreated };
}

function getTemplateName(templateDocId) {
  const index = templateDocIds.indexOf(templateDocId);
  const names = [
    "temel_yc_8_8", "temel8", "temel8.4", "yc", 
    "katilan_listesi", "kkd", "yüksekte_çalışma_talimatı", "temel_8_8", "sınav_soruları", "konulu_sertifika"
  ];
  return index >= 0 ? names[index] : "unknown_template";
}

function getFirmList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FLIST');
  if (!sheet) {
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }
  // Firma isimleri 2. sütunda (indeks 1) olduğu için güncellendi
  return data.slice(1).map(row => row[1]);
}

function getTemplateData() {
  const templateNames = [
    "temel_yc_8_8", "temel8", "temel8.4", "yc",
    "katilan_listesi", "kkd", "yüksekte_çalışma_talimatı", "temel_8_8", "sınav_soruları", "konulu_sertifika"
  ];
  return templateDocIds.map((id, index) => ({
    id: index + 1, // Şablon dizinini 1'den başlat
    name: templateNames[index]
  }));
}

function formatDateToDDMMYYYY(date) {
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

function parseDateFromDDMMYYYY(dateString) {
  if (!dateString) return null;
  const parts = dateString.split('/');
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const year = parseInt(parts[2], 10);
    return new Date(year, month, day);
  }
  return null;
}

/**
 * Dizeyi küçük harfe dönüştürür ve özel karakterleri normalleştirir.
 */
function normalizeString(str) {
  return str.toLowerCase().replace(/ç/g, 'c').replace(/ş/g, 's').replace(/ğ/g, 'g').replace(/ı/g, 'i').replace(/ö/g, 'o').replace(/ü/g, 'u').trim();
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}