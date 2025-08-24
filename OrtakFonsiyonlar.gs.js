/**
 * Düzenli ifadelerde özel anlam taşıyan karakterleri "kaçış" (escape) karakteri ekleyerek güvenli hale getirir.
 * @param {string} str Kaçış karakteri eklenmek istenen metin.
 * @return {string} Güvenli hale getirilmiş metin.
 */
function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Belgeye FLIST sayfasındaki verileri ekler.
 * Bu fonksiyon, iş yeri adının ilk 5 karakterini eşleştirerek çalışır.
 * @param {object} data Ana veri kaynağından gelen obje.
 * @param {GoogleAppsScript.Document.Body} body Google Doküman gövdesi.
 */
function addFLISTData(data, body) {
  const flistSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FLIST');
  if (!flistSheet) return;

  const flistData = flistSheet.getDataRange().getValues();
  const headers = flistData[0];
  const rows = flistData.slice(1);
  
  // İş yerinin adını daha esnek kontrol etmek için
  const matchingRow = rows.find(row => 
    row[headers.indexOf('UNVANI')].slice(0, 5).toLowerCase() === 
    data['İş Yerinin Adı'].slice(0, 5).toLowerCase()
  );

  if (matchingRow) {
    headers.forEach((header, index) => {
      const placeholder = `{{${header}}}`;
      const value = matchingRow[index] || 'Veri bulunamadı'; // Boş hücre için yedek metin
      body.replaceText(escapeRegExp(placeholder), String(value));
    });
  } else {
    console.log('Eşleşme bulunamadı.');
  }
}

/**
 * Belgeye IH_IMZA verisine göre imza görseli ekler.
 * @param {object} responseData Veri objesi.
 * @param {GoogleAppsScript.Document.Body} body Belge gövdesi.
 */
function addSignatureImage(responseData, body) {
  const imzaUrl = responseData.IH_IMZA;
  if (!imzaUrl) return;

  const fileId = getFileIdFromUrl(imzaUrl);
  if (!fileId) return;

  const file = DriveApp.getFileById(fileId);
  const imageBlob = file.getBlob();

  const searchResult = body.findText('{{IH_IMZA}}');
  if (!searchResult) return;

  const foundElement = searchResult.getElement();
  const parent = foundElement.getParent();

  foundElement.asText().setText('');

  const insertedImage = body.insertImage(body.getChildIndex(parent) + 1, imageBlob);
  insertedImage.setWidth(120);
  insertedImage.setHeight(40);
}

// URL'den dosya ID'sini çıkaran yardımcı fonksiyon
function getFileIdFromUrl(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return match[1];
  }
  return null;
}
/**
 * Date nesnesini 'dd/MM/yyyy' formatında bir string'e dönüştürür.
 * @param {Date} date Formatlanacak Date nesnesi.
 * @return {string} 'dd/MM/yyyy' formatında tarih string'i.
 */
function formatDateToDDMMYYYY(date) {
    if (!date) return '';
    const d = new Date(date);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}/${month}/${year}`;
}


/**
 * 'dd/MM/yyyy' formatında bir string'i Date nesnesine dönüştürür.
 * @param {string} dateString Ayrıştırılacak tarih string'i.
 * @return {Date|null} Oluşturulan Date nesnesi veya geçersiz format için null.
 */
function parseDateFromDDMMYYYY(dateString) {
    if (!dateString) return null;
    const parts = dateString.split('/');
    // Yıl, ay (0-indexed), gün
    return new Date(parts[2], parts[1] - 1, parts[0]);
}