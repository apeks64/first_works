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
  const content = [];

  selectedTitles.forEach(title => {
    const colIndex = headers.indexOf(title);
    if (colIndex !== -1) {
      // Başlık ve altındaki tüm satırları al
      const columnContent = data.slice(1).map(row => row[colIndex]).filter(cell => cell);
      if (columnContent.length > 0) {
        content.push(title, ...columnContent); // Başlığı ve içeriği ekle
      }
    }
  });
  return content.join('\n'); // Yeni satırla birleştir
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
        // TC Kimlik No'yu string olarak zorla
        if (header === 'TC Kimlik No') {
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
 * Formdan gelen verilere göre sertifika oluşturma işlemini başlatır.
 */
function processFormWithRowNumbers(rowNumbersStr, templates, firm, date1Str, date3Str, date5Str, trainingHours1, trainingHours2, trainingHours3, selectedTopicNames) {
  const rowNumbers = parseRowNumbers(rowNumbersStr);
  const selectedPersonnel = getPersonnelByRows(rowNumbers);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const flistSheet = spreadsheet.getSheetByName('FLIST');
  const certificateCounterSheet = spreadsheet.getSheetByName('SertifikaSayac');
  if (!flistSheet || !certificateCounterSheet)
    throw new Error("FLIST veya SertifikaSayac sayfası bulunamadı!");

  const flistData = flistSheet.getDataRange().getValues();
  const certificateCounterCell = certificateCounterSheet.getRange('A1');

  const flistHeaders = flistData[0];
  const targetFolderId = '1yrwxxbb7WA2MxQ0UoZ5LN9jJahOaD5Uq';
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  const year = new Date().getFullYear();
  let sertifikaCounter = parseInt(certificateCounterCell.getValue(), 10);
  
  const date1 = date1Str ? parseDateFromDDMMYYYY(date1Str) : null;
  const date3 = date3Str ? parseDateFromDDMMYYYY(date3Str) : null;
  const date5 = date5Str ? parseDateFromDDMMYYYY(date5Str) : null;

  const trainingTopicsContent = getTrainingTopicsContent(selectedTopicNames);
  let flistMatch = null;
  for (let k = 1; k < flistData.length; k++) {
    if (flistData[k][0] === firm) {
      flistMatch = flistData[k];
      break;
    }
  }

  const failedCreations = []; 
  let totalDocsCreated = 0; 

  const KATILAN_LISTESI_INDEX = 5;
  const isKatilanListesiSelected = templates.length === 1 && templates[0] === KATILAN_LISTESI_INDEX;

  if (isKatilanListesiSelected) {
    try {
      const templateDocId = templateDocIds[KATILAN_LISTESI_INDEX - 1];
      if (!templateDocId) {
          throw new Error("Katılan Listesi şablon kimliği bulunamadı.");
      }
      const newFileName = `${firm} - Katılımcı Listesi`;
      const newFile = DriveApp.getFileById(templateDocId).makeCopy(newFileName, targetFolder);
      const doc = DocumentApp.openById(newFile.getId());
      const body = doc.getBody();

      const docPlaceholders = {
        '{{FIRMA}}': firm || '',
        '{{TARIH1}}': date1 ? formatDateToDDMMYYYY(date1) : '',
        '{{TARIH3}}': date3 ? formatDateToDDMMYYYY(date3) : '',
        '{{TARIH5}}': date5 ? formatDateToDDMMYYYY(date5) : '',
        '{{SAAT1}}': trainingHours1 || '',
        '{{SAAT2}}': trainingHours2 || '',
        '{{SAAT3}}': trainingHours3 || '',
        '{{EGITIM_KONULARI}}': trainingTopicsContent || '',
      };
      
      // Diğer firma bilgileri
      if (flistMatch) {
          docPlaceholders['{{ADRESI}}'] = flistMatch[flistHeaders.indexOf('ADRESI')] || '';
          docPlaceholders['{{ISGUZMANI}}'] = flistMatch[flistHeaders.indexOf('ISGUZMANI')] || '';
          docPlaceholders['{{IGU_BELGE_NO}}'] = flistMatch[flistHeaders.indexOf('IGU_BELGE_NO')] || '';
      }

      for (const placeholder in docPlaceholders) {
        body.replaceText(placeholder, docPlaceholders[placeholder]);
      }

      // **YENİ DEĞİŞİKLİK: Tabloyu bul ve değiştir**
      let placeholderParagraph = null;
      for (let i = 0; i < body.getNumChildren(); i++) {
        const child = body.getChild(i);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          if (child.asParagraph().getText().includes('Bu metnin hemen altına katılımcı tablosu eklenecektir.')) {
            placeholderParagraph = child.asParagraph();
            break;
          }
        }
      }

      if (placeholderParagraph) {
        // Tabloyu bulanan paragrafın yerine ekle
        const table = placeholderParagraph.getParent().insertTable(placeholderParagraph.getParent().getChildIndex(placeholderParagraph));
        
        // Placeholder paragrafını kaldır
        placeholderParagraph.removeFromParent();

        // Tablonun içini temizle (varsa) ve başlık satırını ekle
        table.clear();
        table.setBorderWidth(1);
        const headerRow = table.appendTableRow();
        headerRow.appendTableCell('Sıra No'); // Yeni sütun
        headerRow.appendTableCell('TC Kimlik No');
        headerRow.appendTableCell('Adı Soyadı');
        headerRow.appendTableCell('Unvanı/Mesleği'); // Başlık düzeltildi
        headerRow.appendTableCell('İmza'); // Yeni sütun

        let sıraNo = 1;
        selectedPersonnel.forEach(p => {
          const dataRow = table.appendTableRow();
          dataRow.appendTableCell(sıraNo++); // Sıra No eklendi
          dataRow.appendTableCell(p['TC Kimlik No'] || '');
          dataRow.appendTableCell(`${p['Adı'] || ''} ${p['Soyadı'] || ''}`);
          dataRow.appendTableCell(p['İşi'] || '');
          dataRow.appendTableCell(''); // İmza için boş hücre
        });
      } else {
        // Eğer yer tutucu metin bulunamazsa, tabloyu belgenin en sonuna ekle
        const table = body.appendTable();
        table.clear();
        table.setBorderWidth(1);
        const headerRow = table.appendTableRow();
        headerRow.appendTableCell('Sıra No'); // Yeni sütun
        headerRow.appendTableCell('TC Kimlik No');
        headerRow.appendTableCell('Adı Soyadı');
        headerRow.appendTableCell('Unvanı/Mesleği'); // Başlık düzeltildi
        headerRow.appendTableCell('İmza'); // Yeni sütun

        let sıraNo = 1;
        selectedPersonnel.forEach(p => {
          const dataRow = table.appendTableRow();
          dataRow.appendTableCell(sıraNo++); // Sıra No eklendi
          dataRow.appendTableCell(p['TC Kimlik No'] || '');
          dataRow.appendTableCell(`${p['Adı'] || ''} ${p['Soyadı'] || ''}`);
          dataRow.appendTableCell(p['İşi'] || '');
          dataRow.appendTableCell(''); // İmza için boş hücre
        });
      }

      doc.saveAndClose();
      totalDocsCreated++;

    } catch (e) {
      Logger.log(`Hata: Katılımcı Listesi belgesi oluşturulamadı. Hata: ${e.message}`);
      failedCreations.push({
        template: 'Katılımcı Listesi',
        error: e.message
      });
    }

  } else {
    // Diğer tüm şablonlar için mevcut döngü
    selectedPersonnel.forEach(person => {
      const sertifikaVerisi = {};
      for (const key in person) {
        sertifikaVerisi[key] = person[key];
      }
      
      sertifikaVerisi['Tarih1'] = date1 ? formatDateToDDMMYYYY(date1) : '';
      sertifikaVerisi['Tarih3'] = date3 ? formatDateToDDMMYYYY(date3) : '';
      sertifikaVerisi['Tarih5'] = date5 ? formatDateToDDMMYYYY(date5) : '';

      const tarih2 = date1 ? new Date(date1) : null;
      if (tarih2) tarih2.setFullYear(tarih2.getFullYear() + 1);
      
      const tarih4 = date3 ? new Date(date3) : null;
      if (tarih4) tarih4.setFullYear(tarih4.getFullYear() + 1);
      const tarih6 = date5 ? new Date(date5) : null;
      if (tarih6) tarih6.setFullYear(tarih6.getFullYear() + 1);

      sertifikaVerisi['Tarih2'] = tarih2 ? formatDateToDDMMYYYY(tarih2) : '';
      sertifikaVerisi['Tarih4'] = tarih4 ? formatDateToDDMMYYYY(tarih4) : '';
      sertifikaVerisi['Tarih6'] = tarih6 ? formatDateToDDMMYYYY(tarih6) : '';
      sertifikaCounter++;
      const sertifikaNo = `${year}-${Utilities.formatString("%05d", sertifikaCounter)}`;
      sertifikaVerisi['SertifikaNo'] = sertifikaNo;

      templates.forEach(templateIndex => {
        const templateDocId = templateDocIds[templateIndex - 1]; 
        if (!templateDocId) {
            Logger.log(`Belirtilen şablon indeksi (${templateIndex}) için Google Doküman Kimliği bulunamadı.`);
            return;
        }
        try {
          const templateName = getTemplateName(templateDocId);
          const işYerininAdıSubstring = firm ? firm.substring(0, Math.min(firm.length, 13)) : "BilinmeyenFirma";
        
          const fileName = `${işYerininAdıSubstring}_${sertifikaVerisi['Adı']} ${sertifikaVerisi['Soyadı']}_${templateName}_Belgesi_${sertifikaNo}`;
          const newDoc = DriveApp.getFileById(templateDocId).makeCopy(fileName, targetFolder);

          const doc = DocumentApp.openById(newDoc.getId());
          const body = doc.getBody();

          for (let key in sertifikaVerisi) {
            const placeholder = `{{${key}}}`;
            body.replaceText(escapeRegExp(placeholder), sertifikaVerisi[key] || '');
          }
 
          body.replaceText('{{EğitimSaati1}}', trainingHours1 || '');
          body.replaceText('{{EğitimSaati2}}', trainingHours2 || '');
          body.replaceText('{{EğitimSaati3}}', trainingHours3 || '');
          if (flistMatch) {
            for (let m = 0; m < flistHeaders.length; m++) {
              const flistPlaceholder = `{{${flistHeaders[m]}}}`;
              body.replaceText(escapeRegExp(flistPlaceholder), flistMatch[m] || '');
            }
          }
          
          body.replaceText('{{EGITIM_KONULARI}}', trainingTopicsContent);
          doc.saveAndClose();
          totalDocsCreated++;
        } catch (e) {
          Logger.log(`Şablon ${templateIndex} (${getTemplateName(templateDocId)}) işlenirken hata: ${e.message}`);
          failedCreations.push({
            person: `${person['Adı']} ${person['Soyadı']}`,
            error: e.message,
            template: getTemplateName(templateDocId)
          });
        }
      });
    });
  }
  
  certificateCounterCell.setValue(sertifikaCounter);
  
  return {
    success: failedCreations.length === 0,
    failedCreations: failedCreations,
    totalDocsCreated: totalDocsCreated
  };
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
  return data.slice(1).map(row => row[0]);
}

function formatDateToDDMMYYYY(date) {
    const d = new Date(date);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}/${month}/${year}`;
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}