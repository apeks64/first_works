function naceTehlikeSinifiGuncelle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const naceSheet = ss.getSheetByName('NACE');
  const atamalarSheet = ss.getSheetByName('ATAMALAR');

  const naceData = naceSheet.getRange('E2:E' + naceSheet.getLastRow()).getValues(); // NACE E sütunu
  const atamalarNace = atamalarSheet.getRange('W2:W' + atamalarSheet.getLastRow()).getValues(); // ATAMALAR W sütunu
  const atamalarTehlike = atamalarSheet.getRange('V2:V' + atamalarSheet.getLastRow()).getValues(); // ATAMALAR V sütunu

  const tehlikeSiniflari = [];

  for (let i = 0; i < naceData.length; i++) {
    const naceTanim = naceData[i][0];
    let tehlikeSinifi = "";

    for (let j = 0; j < atamalarNace.length; j++) {
      if (naceTanim && naceTanim === atamalarNace[j][0]) {
        tehlikeSinifi = atamalarTehlike[j][0];
        break; // eşleşme bulunduysa diğerlerini kontrol etmeye gerek yok
      }
    }

    tehlikeSiniflari.push([tehlikeSinifi]); // G sütunu için değer
  }

  // G sütununa (7. sütun) tehlike sınıfını yaz
  naceSheet.getRange(2, 7, tehlikeSiniflari.length, 1).setValues(tehlikeSiniflari);
}
