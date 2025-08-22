/**
 * Bu fonksiyon, bir Google Sheets dosyasını her açtığınızda otomatik olarak çalışır.
 * Özel menüleri oluşturur ve eğer "EK-2 BİLGİLER" sayfası aktifse, otomatik olarak
 * eski kayıtları silme arayüzünü gösterir.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Menüleri oluşturma
  const ek2Menu = ui.createMenu('📄 EK-2 İşlemleri');
  ek2Menu.addItem('Seçili Satır(lar)dan Belge Oluştur', 'createSingleDocument');
  ek2Menu.addItem('Formdan Belge Oluştur (Filtreleme ile)', 'openFilterAndGenerateDialog');
  ek2Menu.addItem('📁 EK-2 Klasörü Aç', 'klasoruAc');
  ek2Menu.addToUi();

  const egitimMenu = ui.createMenu('🎓 Eğitim İşlemleri');
  egitimMenu.addItem('📜 Eğitim Sertifikası Oluştur', 'showCertificateDialog');
  egitimMenu.addItem('🧾 Eğitime Katılan Tablo Al', 'showEgitimListesiForm');
  egitimMenu.addItem('📥 Katılan Listesi Oluştur', 'showKatilanListesi');
  egitimMenu.addItem('📊 Eğitim Konuları ŞAblonu', 'showEgitimPano');
  egitimMenu.addToUi();

  const meditekMenu = ui.createMenu('👷 Personel İşlemleri');
  meditekMenu.addItem('🏢 Personel Sayfasına Firma Seç', 'showSelectionDialog');
  meditekMenu.addItem('⬇️ Personel Sayfasını Excel Olarak İndir', 'exportToExcel');
  meditekMenu.addItem('👷 Personel Bul', 'showSearchModal');
  meditekMenu.addToUi();

  const flistMenu = ui.createMenu('📂 Firma İşlemleri');
  flistMenu.addItem("📋 Saha Tablosu Al", "showSahaTablosu");
  flistMenu.addItem('🔄 Atamalar Sayfasını Güncelle', 'importLatestExcelToAtamalar');
  flistMenu.addItem('🔧 Veri Aktarımını Çalıştır', 'tumIslemleriCalistir');
  flistMenu.addItem('📍 İlçe Bilgisini Güncelle (Manuel)', 'logUnmatchedDistricts');
  flistMenu.addItem('📆 Yıllık Süreleri Hesapla (Manuel)', 'hesaplaVeYaz');
  flistMenu.addToUi();

  const KISAYOLMenu = ui.createMenu('📂 SayfaKısayolları');
  ui.createMenu("📌 Kısayollar")
    .addItem("LİNK Sayfasına Git", "goToLinkPage")
    .addToUi();
    
  // Kullanıcıya kontrol sağlayan menü öğesi.
  ui.createMenu('🧹 Veri Temizleme')
    .addItem('7 Günden Eski Kayıtları Temizle', 'showFilterDialog')
    .addToUi();
    
  // Dosya açıldığında otomatik olarak çalışacak kısım
  showFilterDialog();
}

function klasoruAc() {
  var folderId = '1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ'; // EK-2 klasörünüzün ID'si
  var url = 'https://drive.google.com/drive/folders/' + folderId;

  var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");google.script.host.close();</script>')
      .setWidth(100)
      .setHeight(10);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Klasör Açılıyor...');
}
