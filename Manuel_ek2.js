const TEMPLATE_FILE_ID = '1jnytmZfHhHi9bkgBCX5uMRLNFTffmGKqTurMfwKzShE';
const DESTINATION_FOLDER_ID = '1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ';

function openFilterAndGenerateDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('FilterAndGenerateDialog')
      .setWidth(600)
      .setHeight(999);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Personel Seç ve Belge Oluştur');
}

/**
 * Filtrelenmiş verilere göre benzersiz İş Yeri ve Adı Soyadı değerlerini alır.
 * @param {Object} filters - Filtreleri içeren nesne.
 * @param {string} filters.bFilter - İş Yerinin Adı için arama metni.
 * @param {string} filters.cdFilter - Adı Soyadı için arama metni.
 * @param {string[]} filters.selectedWorkplaces - Seçili İş Yerleri.
 * @returns {Object} Güncellenmiş benzersiz listeleri içeren nesne.
 */
function getFilteredData(filters) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EK-2 BİLGİLER');
  if (!sheet) return { error: 'EK-2 BİLGİLER sayfası bulunamadı.' };
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  
  const bColIndex = headers.indexOf('İş Yerinin Adı');
  const cColIndex = headers.indexOf('Adı');
  const dColIndex = headers.indexOf('Soyadı');

  if (bColIndex === -1 || cColIndex === -1 || dColIndex === -1) {
    return { error: 'Gerekli sütun başlıkları (İş Yerinin Adı, Adı, Soyadı) bulunamadı.' };
  }
  
  const bFilter = filters.bFilter ? filters.bFilter.trim().toLowerCase() : '';
  const cdFilter = filters.cdFilter ? filters.cdFilter.trim().toLowerCase() : '';
  const selectedWorkplaces = filters.selectedWorkplaces || [];

  const filteredRows = values.slice(1).filter(row => {
    const bValue = String(row[bColIndex]).trim();
    const bValueLower = bValue.toLowerCase();
    const cdValue = `${String(row[cColIndex]).trim()} ${String(row[dColIndex]).trim()}`.toLowerCase();

    const bMatch = !bFilter || bValueLower.includes(bFilter);
    const cdMatch = !cdFilter || cdValue.includes(cdFilter);
    const workplaceSelectionMatch = selectedWorkplaces.length === 0 || selectedWorkplaces.includes(bValue);

    return bMatch && cdMatch && workplaceSelectionMatch;
  });
  
  const filteredWorkplaces = [...new Set(filteredRows.map(row => String(row[bColIndex]).trim()))].filter(v => v).sort();
  const filteredCombinedNames = [...new Set(filteredRows.map(row => `${String(row[cColIndex]).trim()} ${String(row[dColIndex]).trim()}`))].filter(v => v.trim()).sort();
  
  return { 
    'İş Yerinin Adı': filteredWorkplaces,
    'Adı Soyadı': filteredCombinedNames 
  };
}

function generateDocumentsFromFilteredRows(selections) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EK-2 BİLGİLER');
  if (!sheet) {
    return { success: false, message: 'Tablo bulunamadı.' };
  }

  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const bColIndex = headers.indexOf('İş Yerinin Adı');
  const cColIndex = headers.indexOf('Adı');
  const dColIndex = headers.indexOf('Soyadı');

  if (bColIndex === -1 || cColIndex === -1 || dColIndex === -1) {
    return { success: false, message: 'Gerekli sütun başlıkları (İş Yerinin Adı, Adı, Soyadı) bulunamadı.' };
  }

  const rowsToGenerate = allData.map((row, index) => {
    if (index === 0) return -1; // Başlık satırını atla

    const combinedCDValue = `${String(row[cColIndex]).trim()} ${String(row[dColIndex]).trim()}`.trim();

    // Bu kısım UI'dan gelen seçili isimlere göre çalışır, arama kutusuyla değil
    if (selections.selectedNames.includes(combinedCDValue)) {
      return index + 1;
    }
    return -1;
  }).filter(rowNum => rowNum !== -1);
  if (rowsToGenerate.length === 0) {
    return { success: false, message: 'Seçili kriterlere uyan satır bulunamadı.' };
  }

  let createdCount = 0;
  let failedCount = 0;
  rowsToGenerate.forEach((rowNumber) => {
    try {
      generateDocument(rowNumber, true); // Sessiz modda çalıştır
      createdCount++;
    } catch (e) {
      Logger.log(`Satır ${rowNumber} için belge oluşturulurken hata: ` + e.message);
      failedCount++;
    }
  });
  return { 
    success: true, 
    createdCount: createdCount, 
    failedCount: failedCount,
    totalCount: rowsToGenerate.length
  };
}

// --- Belge oluşturma ana fonksiyonu ---
function generateDocument(rowNumber, silentMode = false) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EK-2 BİLGİLER');
    if (!sheet) throw new Error('Tablo bulunamadı.');

    if (!rowNumber) rowNumber = sheet.getLastRow();

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    const responseData = headers.reduce((obj, header, i) => ({ ...obj, [header]: data[i] }), {});
    Logger.log('responseData:', responseData);

    const workplaceName = responseData['İş Yerinin Adı'] ? responseData['İş Yerinin Adı'].substring(0, 13) : 'Bilinmiyor';
    const firstName = responseData['Adı'] || 'Bilinmiyor';
    const lastName = responseData['Soyadı'] || 'Bilinmiyor';

    const fileName = `${workplaceName} ${firstName} ${lastName} - Muayene Formu`;
    const destinationFolder = DriveApp.getFolderById('1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ');
    const template = DriveApp.getFileById('1jnytmZfHhHi9bkgBCX5uMRLNFTffmGKqTurMfwKzShE');
    const newDoc = template.makeCopy(fileName, destinationFolder);
    const doc = DocumentApp.openById(newDoc.getId());
    const body = doc.getBody();

    fillPlaceholders(body, responseData);
    addFLISTData(responseData, body); // Bu fonksiyonun kodu eksik olduğu için yorum satırı yaptım.
    addSignatureImage(responseData, body); // Bu fonksiyonun kodu eksik olduğu için yorum satırı yaptım.

    doc.saveAndClose();
    Logger.log('Belge başarıyla oluşturuldu!');
  } catch (e) {
    Logger.log('Hata: ' + e.message);
    if (!silentMode) {
      SpreadsheetApp.getUi().alert('Hata oluştu: ' + e.message);
    }
    throw e; // Hatanın yayılmasını sağlıyoruz
  }
}

// --- Tek bir belge oluştur (seçili satır) ---
function createSingleDocument() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EK-2 BİLGİLER');
  const activeCell = sheet.getActiveCell();
  const rowNumber = activeCell.getRow();
  generateDocument(rowNumber, true); // Sessiz modda çalıştır
  SpreadsheetApp.getUi().alert('✅ Belge başarıyla oluşturuldu!');
}

function convertToTurkishUpperCase(str) {
  var turkishMap = {
    'ı': 'I', 'i': 'İ', 'ç': 'Ç', 'ş': 'Ş', 'ğ': 'Ğ', 'ü': 'Ü', 'ö': 'Ö',
  };
  return str.replace(/[ıiçşğüö]/g, match => turkishMap[match]).toUpperCase();
}

function convertAllToUpperCase() {
  const sheetNames = ['EK-2 BİLGİLER', 'EK-2_DB'];
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    const range = sheet.getDataRange();
    const values = range.getValues();

    for (let i = 1; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        if (typeof values[i][j] === 'string' && !values[i][j].includes('@')) {
          values[i][j] = convertToTurkishUpperCase(values[i][j]);
        }
      }
    }
    range.setValues(values);
  });
}

function archiveNewFormResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetA = ss.getSheetByName('EK-2 BİLGİLER'); 
  const sheetB = ss.getSheetByName('EK-2_DB'); 

  if (!sheetA || !sheetB) {
    Logger.log("❌ Bir veya daha fazla sayfa mevcut değil.");
    return;
  }

  const tcColumnIndex = 5; 
  const dataB = sheetB.getDataRange().getValues();
  const existingTCNumbers = new Set(dataB.slice(1).map(row => row[tcColumnIndex - 1]));
  const dataA = sheetA.getDataRange().getValues();
  Logger.log("✅ EK-2 BİLGİLER'den alınan satır sayısı: " + dataA.length);

  const newUniqueRows = [];
  for (let i = 1; i < dataA.length; i++) {
    const row = dataA[i];
    const tcNumber = row[tcColumnIndex - 1];

    if (!existingTCNumbers.has(tcNumber)) {
      newUniqueRows.push(row);
      existingTCNumbers.add(tcNumber);
    }
  }

  Logger.log("✅ EK-2_DB'ye eklenecek yeni satır sayısı: " + newUniqueRows.length);
  Logger.log(newUniqueRows);
  if (newUniqueRows.length > 0) {
    const startRow = sheetB.getLastRow() + 1;
    sheetB.getRange(startRow, 1, newUniqueRows.length, newUniqueRows[0].length).setValues(newUniqueRows);
    Logger.log("✅ Yeni satırlar başarıyla eklendi.");
  } else {
    Logger.log("ℹ️ Yeni eklenen satır yok.");
  }
}

function fillPlaceholders(body, data) {
  const yesNoPlaceholders = [
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Balgamlı öksürük]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Nefes darlığı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Göğüs ağrısı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Çarpıntı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Sırt ağrısı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [İshal veya kabızlık]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Eklemlerde ağrı]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Bayılma]',
    'Son bir yıl içinde ve sürekli olarak aşağıdaki yakınmalardan herhangi birini  yaşadınız mı? [Allerji]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Kalp hastalığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Şeker hastalığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Böbrek rahatsızlığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Sarılık]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Mide veya on iki parmak ülseri]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [İşitme kaybı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Görme bozukluğu]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Sinir sistemi hastalığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Deri hastalığı]',
    'Aşağıdaki hastalıklardan herhangi birini geçirdiniz mi? [Besin zehirlenmesi]',
    'Hastanede yattınız mı?',
    'Ameliyat Oldunuz mu?',
    'İş kazası geçirdiniz mi?',
    'Sigara içiyor musunuz?',
    'Alkol alıyor musunuz?',
  ];
  yesNoPlaceholders.forEach((placeholder) => {
    const yesKey = `{{${placeholder} - EVET}}`;
    const noKey = `{{${placeholder} - HAYIR}}`;
    const quitKey = `{{${placeholder} - BIRAKMIŞ}}`;

    if (data[placeholder] === 'EVET') {
      body.replaceText(escapeRegExp(yesKey), '✓');
      body.replaceText(escapeRegExp(noKey), '');
      body.replaceText(escapeRegExp(quitKey), '');
    } else if (data[placeholder] === 'HAYIR') {
      body.replaceText(escapeRegExp(yesKey), '');
      body.replaceText(escapeRegExp(noKey), '✓');
      body.replaceText(escapeRegExp(quitKey), '');
    } else if (data[placeholder] === 'BIRAKMIŞ') {
      body.replaceText(escapeRegExp(yesKey), '');
      body.replaceText(escapeRegExp(noKey), '');
      body.replaceText(escapeRegExp(quitKey), '✓');
    } else {
      body.replaceText(escapeRegExp(yesKey), '');
      body.replaceText(escapeRegExp(noKey), '');
      body.replaceText(escapeRegExp(quitKey), '');
    }
  });

  // Tüm yer imleri için döngü
  for (const key in data) {
    let placeholder = `{{${key}}}`;
    let value = data[key];

    if (value instanceof Date) {
      value = formatDateToDDMMYYYY(value);
    }
    body.replaceText(escapeRegExp(placeholder), String(value || ''));
  }
}
function formatDateToDDMMYYYY() { /* Fonksiyon kodunuz eksik */ }
function escapeRegExp() { /* Fonksiyon kodunuz eksik */ }
function addFLISTData() { /* Fonksiyon kodunuz eksik */ }
function addSignatureImage() { /* Fonksiyon kodunuz eksik */ }