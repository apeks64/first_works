/**
 * Form yanıtından sonra otomatik olarak tetiklenir (tetikleyici ayarlanmalı).
 * @param {Object} e - Form yanıtı olayı nesnesi.
 */
function onFormSubmit(e) {
  try {
    const row = e.range.getRow();
    // Bu fonksiyonlar, aynı dosyada tanımlı olmalıdır.
    convertAllToUpperCase();       // Tüm verileri büyük harfe çevirme
    archiveNewFormResponses();     // Yeni yanıtları arşivleme
    generateDocument(row, true);   // Belge oluşturma (sessiz modda)
    Logger.log("Form yanıtı işlemi başarıyla tamamlandı.");
  } catch (error) {
    Logger.log("Hata oluştu: " + error.message);
  }
}