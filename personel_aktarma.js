// Kullanıcıya seçim yapabilmesi için HTML dialog göster
function showSelectionDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ek2Sheet = ss.getSheetByName("EK-2 BİLGİLER");
  var atamalarSheet = ss.getSheetByName("ATAMALAR");

  if (!ek2Sheet || !atamalarSheet) {
    Logger.log("Sayfalardan biri bulunamadı!");
    return;
  }

  // EK-2 BİLGİLER sayfasından B sütunundaki benzersiz verileri alalım
  var columnB = ek2Sheet.getRange(2, 2, ek2Sheet.getLastRow() - 1, 1).getValues();
  var uniqueValues = [...new Set(columnB.flat())]; // Benzersiz değerler

  // ATAMALAR sayfasındaki R sütunundan benzersiz verileri alalım
  var columnR = atamalarSheet.getRange(2, 18, atamalarSheet.getLastRow() - 1, 1).getValues();
  var uniqueAtamalarValues = [...new Set(columnR.flat())]; // Benzersiz değerler

  // HTML sayfasını oluşturup kullanıcıya göstermek
  var htmlOutput = HtmlService.createHtmlOutputFromFile('selectionDialog')
      .setWidth(500)
      .setHeight(600);
  htmlOutput.append('<script>var uniqueValues = ' + JSON.stringify(uniqueValues) + ';</script>');
  htmlOutput.append('<script>var uniqueAtamalarValues = ' + JSON.stringify(uniqueAtamalarValues) + ';</script>');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Veri Seçimi');
}

// Seçilen verileri PERSONEL sayfasına aktar
function transferDataToPersonel(selectedValues, atamalarSelectedValues) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ek2Sheet = ss.getSheetByName("EK-2 BİLGİLER");
  var atamalarSheet = ss.getSheetByName("ATAMALAR");
  var personelSheet = ss.getSheetByName("PERSONEL");

  if (!ek2Sheet || !atamalarSheet || !personelSheet) {
    Logger.log("Sayfalardan biri bulunamadı!");
    return;
  }

  // Personel sayfasını temizle (başlık hariç)
  var lastRowPersonel = personelSheet.getLastRow();
  if (lastRowPersonel > 1) {
    personelSheet.getRange(2, 1, lastRowPersonel - 1, personelSheet.getLastColumn()).clear();
  }

  // EK-2 BİLGİLER sayfasından veri al
  var lastRowEk2 = ek2Sheet.getLastRow();
  var dataPart1 = ek2Sheet.getRange(2, 2, lastRowEk2 - 1, 9).getValues(); // B-J (Adı - Medeni Durumu)
  var dataPart2 = ek2Sheet.getRange(2, 12, lastRowEk2 - 1, 9).getValues(); // L-T (Adresi - Kan Grubu)

  var newData = [];

  // EK-2 BİLGİLER sayfasındaki seçilen verilerle eşleşme ve aktarım
  for (var i = 0; i < dataPart1.length; i++) {
    if (selectedValues.includes(dataPart1[i][0])) {
      var row = new Array(21).fill(""); // 21 sütunluk boş bir dizi
      row.splice(2, 8, ...dataPart1[i]); // C-K sütunları (B-I verisi)
      row.splice(11, 9, ...dataPart2[i]); // L-T sütunları (L-T verisi)
      newData.push(row);
    }
  }

  if (newData.length > 0) {
    personelSheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
  }

  // ATAMALAR sayfasındaki R sütunundan verileri al ve Personel sayfasına aktar
  var atamalarData = atamalarSheet.getRange(2, 1, atamalarSheet.getLastRow() - 1, 19).getValues();
  var personelNewData = personelSheet.getRange(2, 3, personelSheet.getLastRow() - 1, 1).getValues(); // C sütunundaki veriler

  for (var i = 0; i < personelNewData.length; i++) {
    var personelValue = personelNewData[i][0]?.toString().substring(0, 5);
    var matchedRow = atamalarData.find(row => row[17]?.toString().substring(0, 5) === personelValue);

    if (matchedRow) {
      personelSheet.getRange(i + 2, 1).setValue(matchedRow[17]); // "R" sütununa aktar (Hizmet Alan İşyeri Unvanı)
      personelSheet.getRange(i + 2, 2).setValue(matchedRow[18]); // "S" sütununa aktar (SGK/DETSİS No)
    } else {
      personelSheet.getRange(i + 2, 1, 1, 2).clear();
    }
  }

  Logger.log("Veriler başarıyla PERSONEL sayfasına aktarıldı!");
}

// Excel olarak dışa aktarma fonksiyonu
function exportToExcel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PERSONEL");
  if (!sheet) {
    Logger.log("PERSONEL sayfası bulunamadı!");
    return;
  }

  var spreadsheetId = ss.getId();
  var sheetId = sheet.getSheetId();
  var fileName = "Personel_Verileri.xlsx";
  var url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx&gid=${sheetId}`;

  try {
    var htmlOutput = HtmlService.createHtmlOutput(
      `<a href="${url}" target="_blank" download="${fileName}">Excel Dosyasını İndir</a>`
    ).setWidth(300).setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Excel Olarak İndir");
  } catch (e) {
    Logger.log("UI işlemi bu bağlamda çalıştırılamaz: " + e.message);
  }
}
