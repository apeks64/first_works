// @ts-nocheck
const TEMPLATE_FILE_ID = '1jnytmZfHhHi9bkgBCX5uMRLNFTffmGKqTurMfwKzShE'; // Şablon dosyanızın ID'si
const DESTINATION_FOLDER_ID = '1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ'; // Hedef klasör ID'si

// Tek bir fonksiyonla aktif sayfadan belge oluşturma
function createDocument() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Aktif sayfayı al
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Aktif sayfa bulunamadı.');
    return;
  }

  const selectedRangeList = sheet.getActiveRangeList();  // Çoklu aralık desteği

  if (!selectedRangeList) {
    SpreadsheetApp.getUi().alert('Lütfen belge oluşturmak için satırları seçin.');
    return;
  }

  const ranges = selectedRangeList.getRanges();  // Seçilen tüm aralıkları al

  ranges.forEach(range => {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    
    for (let i = 0; i < numRows; i++) {
      generateDocument(startRow + i, true); // Sessiz modda çalıştır
      Utilities.sleep(1000);  // Her belge arasında 1 saniye bekletme
    }
  });
  SpreadsheetApp.getUi().alert('Aktif sayfadan belge oluşturma işlemi tamamlandı.');
}

function convertToTurkishUpperCase(str) {
  var turkishMap = {
    'ı': 'I', 'i': 'İ', 'ç': 'Ç', 'ş': 'Ş', 'ğ': 'Ğ', 'ü': 'Ü', 'ö': 'Ö', 'ç': 'Ç',
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
  Logger.log(newUniqueRows); // Tüm yeni satırları yazdır

  if (newUniqueRows.length > 0) {
    const startRow = sheetB.getLastRow() + 1;
    sheetB.getRange(startRow, 1, newUniqueRows.length, newUniqueRows[0].length).setValues(newUniqueRows);
    Logger.log("✅ Yeni satırlar başarıyla eklendi.");
  } else {
    Logger.log("ℹ️ Yeni eklenen satır yok.");
  }
}


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
    const fileName = `${workplaceName.slice(0, 13)} ${firstName} ${lastName} - Muayene Formu`;
    const destinationFolder = DriveApp.getFolderById('1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ');

    const template = DriveApp.getFileById(TEMPLATE_FILE_ID);
    const newDoc = template.makeCopy(fileName, destinationFolder); // Yeni dosya
    const doc = DocumentApp.openById(newDoc.getId());
    const body = doc.getBody();
    fillPlaceholders(body, responseData);
    addFLISTData(responseData, body);
    addSignatureImage(responseData, body); // İmza ekleme fonksiyonunu çağır

    doc.saveAndClose();
    Logger.log('Belge başarıyla oluşturuldu!');
  } catch (e) {
    Logger.log('Hata: ' + e.message);
    if (!silentMode) {
      SpreadsheetApp.getUi().alert('Hata oluştu: ' + e.message);
    }
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
  // formatDateToDDMMYYYY fonksiyonu OrtakFonksiyonlar.gs dosyasında olmalı
  value = formatDateToDDMMYYYY(value);
}

    body.replaceText(escapeRegExp(placeholder), String(value || ''));
  }
}