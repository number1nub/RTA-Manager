echo off
cls
echo.
echo.
set /p summary=Enter file summary: 
echo.
echo.
echo Loading to googlecode.....
echo.
echo.
python googlecode_upload.py -s "%summary%" -p rta-manager -u number1nub@gmail.com -w ee5wH9vJ7pa9 "C:\Dropbox\SVN\Halliburton RTA Manager\build\Halliburton RTA Manager-SetupFiles\Halliburton RTA Manager.msi"
