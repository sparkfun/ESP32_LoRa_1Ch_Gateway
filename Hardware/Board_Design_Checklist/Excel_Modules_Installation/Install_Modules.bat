IF NOT EXIST C:\Python27 GOTO error

md C:\Python27\xlrd-0.9.3
xcopy .\xlrd-0.9.3 C:\Python27\xlrd-0.9.3 /s /e /y

md C:\Python27\xlwt-0.7.5
xcopy .\xlwt-0.7.5 C:\Python27\xlwt-0.7.5 /s /e /y

md C:\Python27\xlutils-1.7.1
xcopy .\xlutils-1.7.1 C:\Python27\xlutils-1.7.1 /s /e /y

cd /D C:\Python27\xlrd-0.9.3
C:\Python27\python.exe setup.py install

cd /D C:\Python27\xlwt-0.7.5
C:\Python27\python.exe setup.py install

cd /D C:\Python27\xlutils-1.7.1
C:\Python27\python.exe setup.py install

@echo off
echo.
echo Install complete!
pause
exit /B

:error
cls
@echo off
echo Error: C:\Python27 folder does not exist, halting.
pause
