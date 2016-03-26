pyinstaller -F -i icon.ico WifiPPT.py
mkdir .\dist\template
mkdir .\dist\static
xcopy .\template\*.* /s .\dist\template
xcopy .\static\*.* /s .\dist\static

pause