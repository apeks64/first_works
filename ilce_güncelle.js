function updateDistricts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FLIST");
  var data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues();
  
  var districtMap = {
    "03502": "BAYINDIR", "03504": "BORNOVA", "03505": "ÇEŞME", "03507": "FOÇA",
    "03509": "KARŞIYAKA", "03510": "KEMALPAŞA", "03512": "KİRAZ", "03513": "MENEMEN",
    "03514": "ÖDEMİŞ", "03518": "TORBALI", "03519": "NARLIDERE", "03521": "BUCA",
    "03522": "MENDERES", "03523": "KONAK", "03524": "BEYDAĞ", "03525": "BALÇOVA",
    "03526": "ÇİĞLİ", "03528": "GÜZELBAHÇE", "03529": "GAZİEMİR", "03530": "BAYRAKLI", "03531": "KARABAĞLAR"
  };
  
  var output = [];
  for (var i = 0; i < data.length; i++) {
    var sgkNo = data[i][0];
    if (sgkNo && sgkNo.length >= 21) {
      var districtCode = sgkNo.substring(16, 21); // 16-21 karakterler arası
      output.push([districtMap[districtCode] || "BİLİNMİYOR"]);
    } else {
      output.push([""]);
    }
  }
  
  sheet.getRange(2, 9, output.length, 1).setValues(output); // I sütununa yazdırma
}
