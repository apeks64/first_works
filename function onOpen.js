function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const ek2Menu = ui.createMenu('👨‍⚕️ EK-2 İşlemleri');
  ek2Menu.addItem('📜Seçili Satır(lar)dan EK-2 Oluştur', 'createSingleDocument');
  ek2Menu.addItem('📋Menüden EK-2 Oluştur (Filtreleme ile)', 'openFilterAndGenerateDialog');
  ek2Menu.addItem('🧹Eski Kayıtları Sil', 'FilterDialog');
  ek2Menu.addItem('📁 EK-2 Klasörünü Aç', 'klasoruAc');
  ek2Menu.addToUi();

  const egitimMenu = ui.createMenu('🎓 Eğitim İşlemleri');
  egitimMenu.addItem('📜 Eğitim Sertifikası Oluştur', 'showCertificateDialog');
  egitimMenu.addItem('🧾 Eğitime Katılan Tablo Al', 'showEgitimListesiForm');
  egitimMenu.addToUi();

  const meditekMenu = ui.createMenu('👷 Personel İşlemleri');
  meditekMenu.addItem('🏢 Personel Sayfasına Firma Seç', 'showSelectionDialog');
  meditekMenu.addItem('⬇️ Personel Sayfasını Excel Olarak İndir', 'exportToExcel');
  meditekMenu.addItem('👷 Personel Bul', 'showSearchModal');
  meditekMenu.addToUi();

  const flistMenu = ui.createMenu('🏭 Firma İşlemleri');
  flistMenu.addItem("📋 Saha Tablosu Al", "showSahaTablosu");
  flistMenu.addItem('🔄 Atamalar Sayfasını Güncelle', 'importLatestExcelToAtamalar');
  flistMenu.addItem('🔧 Veri Aktarımını Çalıştır', 'tumIslemleriCalistir');
  flistMenu.addItem('📍 İlçe Bilgisini Güncelle (Manuel)', 'logUnmatchedDistricts');
  flistMenu.addItem('📆 Yıllık Süreleri Hesapla (Manuel)', 'hesaplaVeYaz');
  flistMenu.addToUi();

  const KISAYOLMenu = ui.createMenu('📂 SayfaKısayolları');
  ui.createMenu("📌 Kısayollar")
    .addItem("🪶LİNK Sayfasına Git", "goToLinkPage")
    .addToUi();

  // Uyarı mesajı - sadece onOpen'da
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const targetSheetName = 'EK-2 BİLGİLER';

  if (activeSheet.getName() === targetSheetName) {
    SpreadsheetApp.getUi().alert('Lütfen Ek-2 İşlemleri altında bulunan Eski Kayıtları Sil menüsünden işi biten firmaları siliniz.');
  }

  // Global değişken tanımla
  globalThis.ek2WarningShown = true;
}

// Sayfa değişikliklerini takip et
function onSelectionChange(e) {
  const activeSheet = e.source.getActiveSheet();
  const targetSheetName = 'EK-2 BİLGİLER';
  
  // Eğer EK-2 BİLGİLER sayfasına geçildiyse ve daha önce uyarı gösterilmediyse
  if (activeSheet.getName() === targetSheetName && !globalThis.ek2WarningShown) {
    SpreadsheetApp.getUi().alert('Lütfen Ek-2 İşlemleri altında bulunan Eski Kayıtları Sil menüsünden işi biten firmaları siliniz.');
    globalThis.ek2WarningShown = true;
  }
}

/**
 * Belirtilen klasör ID'sine sahip Google Drive klasörünü açar.
 */
function klasoruAc() {
  var folderId = '1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ'; // EK-2 klasörünüzün ID'si
  var url = 'https://drive.google.com/drive/folders/' + folderId;

  var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");google.script.host.close();</script>')
      .setWidth(100)
      .setHeight(10);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Klasör Açılıyor...');
}