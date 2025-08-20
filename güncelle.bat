@echo off
set "DIZIN=C:\1-Node_Yedek"

echo ----------------------------------------------------
echo Apps Script projesini guncelle ve GitHub'a gonder
echo ----------------------------------------------------
echo.

cd /d "%DIZIN%"

echo.
echo 1. Adim: Apps Script'ten son degisiklikler cekiliyor...
call clasp pull

echo.
echo 2. Adim: Git degisiklikleri ekleniyor...
git add .

echo.
echo 3. Adim: Degisiklikler yerel olarak kaydediliyor...
git commit -m "clasp pull ile son guncellemeler eklendi"

echo.
echo 4. Adim: Degisiklikler GitHub'a gonderiliyor...
git push -u origin master

echo.
echo ----------------------------------------------------
echo Islem tamamlandi.
echo ----------------------------------------------------
pause