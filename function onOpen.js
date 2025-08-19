function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // === EK-2 İşlemleri Menüsü ===
  const ek2Menu = ui.createMenu('📄 EK-2 İşlemleri');
  ek2Menu.addItem('Seçili Satır(lar)dan Belge Oluştur', 'createDocument');
  ek2Menu.addItem('📁 EK-2 Klasörü Aç', 'openDestinationFolder');
  ek2Menu.addItem('🗑️ Tablodan Satır Silme', 'deleteRows');
  ek2Menu.addToUi();

  // === Eğitim İşlemleri Menüsü ===
  const egitimMenu = ui.createMenu('🎓 Eğitim İşlemleri');
  egitimMenu.addItem('📜 Eğitim Sertifikası Oluştur', 'showCertificateDialog');
  egitimMenu.addItem('🧾 Eğitime Katılan Tablo Al', 'showEgitimListesiForm');
  egitimMenu.addItem('📥 Katılan Listesi Oluştur', 'showKatilanListesi');
  egitimMenu.addItem('📊 Eğitim Konuları ŞAblonu', 'showEgitimPano');
  egitimMenu.addToUi();

  // === Meditek Personel İşlemleri Menüsü ===
  const meditekMenu = ui.createMenu('👷Personel İşlemleri');
  meditekMenu.addItem('🏢 Personel Sayfasına Firma Seç', 'showSelectionDialog');
  meditekMenu.addItem('⬇️ Personel Sayfasını Excel Olarak İndir', 'exportToExcel');
  meditekMenu.addItem('👷 Personel Bul', 'showSearchModal');
  meditekMenu.addToUi();

  // === FLIST İşlemler Menüsü ===
  const flistMenu = ui.createMenu('📂 Firma İşlemleri');
  flistMenu.addItem("📋 Saha Tablosu Al", "showSahaTablosu")
  flistMenu.addItem('🔄 Atamalar Sayfasını Güncelle', 'importLatestExcelToAtamalar');
  flistMenu.addItem('🔧 Veri Aktarımını Çalıştır', 'tumIslemleriCalistir');
  flistMenu.addItem('📍 İlçe Bilgisini Güncelle (Manuel)', 'logUnmatchedDistricts');
  flistMenu.addItem('📆 Yıllık Süreleri Hesapla (Manuel)', 'hesaplaVeYaz');
  flistMenu.addToUi();

  // === KISAYOL

  const KISAYOLMenu = ui.createMenu('📂 SayfaKısayolları');
  ui.createMenu("📌 Kısayollar")
    .addItem("LİNK Sayfasına Git", "goToLinkPage")
    .addToUi();
}

function goToLinkPage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LİNK"); // Sayfa adı tam olarak "LİNK" olmalı
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('❌ "LİNK" sayfası bulunamadı!');
  }
}

function klasoruAc() {
  var folderId = '1TVjW_R8SSZsQ3qNcyA6X6ghlJNyN0niJ'; // EK-2 klasörünüzün ID'si
  var url = 'https://drive.google.com/drive/folders/' + folderId;

  var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");google.script.host.close();</script>')
      .setWidth(100)
      .setHeight(10);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Klasör Açılıyor...');
}
