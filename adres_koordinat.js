function getCoordinates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("F2:F" + sheet.getLastRow()); // F sütunundaki adresleri al
  var values = range.getValues();
  var geoValues = [];
  var midPoint = Math.floor(values.length / 2); // Adresleri ikiye ayır

  // İlk yarıyı işleme
  for (var i = 0; i < midPoint; i++) {
    var address = values[i][0];
    if (address) {
      try {
        var geoData = Maps.newGeocoder().geocode(address);
        if (geoData.status == 'OK') {
          var lat = geoData.results[0].geometry.location.lat;
          var lng = geoData.results[0].geometry.location.lng;
          geoValues.push([lat + ", " + lng]);
        } else {
          geoValues.push(["Hata"]);
        }
      } catch (e) {
        geoValues.push(["Hata: " + e.message]);
      }
    } else {
      geoValues.push([""]);
    }
  }

  sheet.getRange(2, 7, geoValues.length, 1).setValues(geoValues); // Sonuçları I sütununa yaz

  Utilities.sleep(3000); // 3 saniye bekle

  // İkinci yarıyı işleme
  geoValues = [];
  for (var i = midPoint; i < values.length; i++) {
    var address = values[i][0];
    if (address) {
      try {
        var geoData = Maps.newGeocoder().geocode(address);
        if (geoData.status == 'OK') {
          var lat = geoData.results[0].geometry.location.lat;
          var lng = geoData.results[0].geometry.location.lng;
          geoValues.push([lat + ", " + lng]);
        } else {
          geoValues.push(["Hata"]);
        }
      } catch (e) {
        geoValues.push(["Hata: " + e.message]);
      }
    } else {
      geoValues.push([""]);
    }
  }

  sheet.getRange(2 + midPoint, 7, geoValues.length, 1).setValues(geoValues); // Sonuçları I sütununa yaz
}
