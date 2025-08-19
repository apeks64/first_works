function kontrolEtVeMailGonder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FLIST');
  var data = sheet.getDataRange().getValues();

  var firmaColumnIndex = 0; // A
  var sgknoColumnIndex = 2; // C
  var ilColumnIndex = 9;    // J
  var uzmanColumnIndex = 14; // O
  var tarihColumnIndex = 24; // Y
  var renkColumnIndex = 0; // A sütunu baz alınarak renk kontrolü

  var today = new Date();

  // Sadece Pazartesi çalış
  if (today.getDay() !== 1) {
    Logger.log("Bugün Pazartesi değil.");
    return;
  }

  var subject = "Lütfen mail içeriğindeki bilgileri kontrol ediniz.";
  var recipientEmails = ["senaosgb1@gmail.com", "senaosgb@gmail.com"];

  var adpTarihiGecmisFirmalar = [];
  var uzmanOlmayanFirmalar = [];
  var turuncuDolguluSatirlar = [];

  for (var i = 1; i < data.length; i++) {
    var firma = data[i][firmaColumnIndex];
    var sgkno = data[i][sgknoColumnIndex];
    var il = data[i][ilColumnIndex];
    var uzman = data[i][uzmanColumnIndex];
    var sonTarih = data[i][tarihColumnIndex];
    var renk = sheet.getRange(i + 1, renkColumnIndex + 1).getBackground();

    // 1. ADP ve RD Tarihleri Geçmiş Firmalar
    if (sonTarih && sonTarih instanceof Date && sonTarih < today) {
      adpTarihiGecmisFirmalar.push(`- ${firma} (Tarih: ${sonTarih.toLocaleDateString()})`);
    }

    // 2. Uzmanı Olmayan Firmalar (İL 'MANİSA' değilse)
    var ilStr = (il || "").toString().toUpperCase();
    if ((!uzman || uzman === "") && !ilStr.includes("MANİSA")) {
      uzmanOlmayanFirmalar.push(`- ${firma} (SGK No: ${sgkno}, İl: ${ilStr || "Belirtilmemiş"})`);
    }

    // 3. Turuncu Renkli Satırlar
    if (renk === '#f4c7c3') {
      turuncuDolguluSatirlar.push(`- ${firma}`);
    }
  }

  // Mail içeriği
  var body = "Merhaba,\n\nLütfen aşağıdaki firmalarla ilgili kontrolü sağlayınız:\n\n";

  body += "1. ADP ve RD Tarihleri Geçmiş Firmalar:\n";
  body += adpTarihiGecmisFirmalar.length > 0 ? adpTarihiGecmisFirmalar.join("\n") : "Yok";
  body += "\n\n";

  body += "2. Uzmanı Olmayan Firmalar (İl bilgisi 'MANİSA' olmayanlar):\n";
  body += uzmanOlmayanFirmalar.length > 0 ? uzmanOlmayanFirmalar.join("\n") : "Yok";
  body += "\n\n";

  body += "3. İSG KATİP Excel'inde Bulunmayan (Ataması Olmayan) Firmalar:\n";
  body += turuncuDolguluSatirlar.length > 0 ? turuncuDolguluSatirlar.join("\n") : "Yok";
  body += "\n\nGereğini rica ederiz.";

  if (adpTarihiGecmisFirmalar.length > 0 || uzmanOlmayanFirmalar.length > 0 || turuncuDolguluSatirlar.length > 0) {
    GmailApp.sendEmail(recipientEmails.join(","), subject, body);
    Logger.log("E-posta gönderildi:\n" + body);
  } else {
    Logger.log("Gönderilecek bir durum bulunamadı.");
  }
}
