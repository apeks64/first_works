/**
 * Yeni bir Google Dokümanı oluşturur ve içine Katılımcı Listesi şablonunu
 * görseldeki gibi dikey yönde ve yer tutucuları bir tablo içinde olacak şekilde ekler.
 *
 * NOT: Bu fonksiyonu çalıştırdıktan sonra oluşan belgenin ID'sini alıp sertifika.txt dosyasındaki
 * 'katilan_listesi' şablon ID'si ile değiştirmelisiniz.
 */
function createKatilanListesiTemplate() {
  const docTitle = 'Katılanlar Listesi Şablonu (Yeni-Tablolu)';
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  // Şablon içeriğini ekle
  body.appendParagraph('                                [FİRMA LOGOSU İÇİN BOŞLUK]')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('                                    KATILIMCILAR LİSTESİ')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('');

  // Tüm yer tutucuları içeren ana tabloyu oluştur
  const mainTable = body.appendTable();
  mainTable.setBorderWidth(0); // Çerçevesiz bir tablo

  // Her bir tablo satırını ve hücrelerini ayrı ayrı oluştur
  
  // FİRMA UNVANI satırı
  let row1 = mainTable.appendTableRow();
  row1.appendTableCell('FİRMA UNVANI:');
  row1.appendTableCell('{{FIRMA}}');
  
  // EĞİTİMİN YERİ satırı
  let row2 = mainTable.appendTableRow();
  row2.appendTableCell('EĞİTİMİN YERİ:');
  row2.appendTableCell('{{ADRESI}}');
  
  // EĞİTİMİN TARİHİ satırı
  let row3 = mainTable.appendTableRow();
  row3.appendTableCell('EĞİTİMİN TARİHİ:');
  row3.appendTableCell('{{TARIH1}} {{TARIH3}} {{TARIH5}}');
  
  // EĞİTİMİN SAATİ satırı
  let row4 = mainTable.appendTableRow();
  row4.appendTableCell('EĞİTİMİN SAATİ:');
  row4.appendTableCell('{{SAAT1}} {{SAAT2}} {{SAAT3}}');
  
  // EĞİTİMİ VEREN satırı
  let row5 = mainTable.appendTableRow();
  row5.appendTableCell('EĞİTİMİ VEREN:');
  row5.appendTableCell('{{ISGUZMANI}}');
  
  // EĞİTİM VEREN BELGE NO satırı
  let row6 = mainTable.appendTableRow();
  row6.appendTableCell('EĞİTİM VEREN BELGE NO:');
  row6.appendTableCell('{{IGU_BELGE_NO}}');

  // Tablodan sonra Eğitim Konuları bölümünü ekle
  body.appendParagraph(''); // Boş satır
  body.appendParagraph('EĞİTİMİN KONULARI:');
  body.appendParagraph('{{EGITIM_KONULARI}}');

  body.appendParagraph(''); // Boş satır
  body.appendParagraph('Bu metnin hemen altına katılımcı tablosu eklenecektir.');

  doc.saveAndClose();

  const docId = doc.getId();
  const docUrl = doc.getUrl();

  Logger.log('Yeni şablon başarıyla oluşturuldu.');
  Logger.log('Belge Adı: ' + docTitle);
  Logger.log('Belge ID: ' + docId);
  Logger.log('Belge URL: ' + docUrl);

  SpreadsheetApp.getUi().alert(`Yeni şablonunuz başarıyla oluşturuldu!\n\nBelgenin ID'sini kopyalayıp sertifika.txt dosyasındaki "katilan_listesi" şablon ID'si ile değiştirin.\n\nBelge ID: ${docId}\n\nBelgeye gitmek için URL: ${docUrl}`);
}