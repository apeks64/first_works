/**
 * Düzenli ifadelerde özel anlam taşıyan karakterleri "kaçış" (escape) karakteri ekleyerek güvenli hale getirir.
 * @param {string} str Kaçış karakteri eklenmek istenen metin.
 * @return {string} Güvenli hale getirilmiş metin.
 */
function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}



function convertToTurkishUpperCase(str) {
  var turkishMap = {
    'ı': 'I', 'i': 'İ', 'ç': 'Ç', 'ş': 'Ş', 'ğ': 'Ğ', 'ü': 'Ü', 'ö': 'Ö',
  };
  return str.replace(/[ıiçşğüö]/g, match => turkishMap[match]).toUpperCase();
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