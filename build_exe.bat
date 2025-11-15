@echo off
echo ======================================
echo Building GujaratiPDFTool EXE (Direct Path)
echo ======================================

"C:\Users\ibrahim\AppData\Roaming\Python\Python313\Scripts\pyinstaller.exe" --onefile --noconsole src\GujaratiAllInOneGUI_v2.py --name GujaratiPDFTool

echo ======================================
echo BUILD FINISHED! Check the /dist folder.
echo ======================================
pause