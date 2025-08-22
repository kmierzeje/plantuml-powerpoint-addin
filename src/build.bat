@echo off
set TARGET=%~f1
set EXT=%TARGET:~-4%
set COMPRESS_CURDIR=powershell Compress-Archive -Update * '%TARGET%.zip'

del "%TARGET%.zip" "%TARGET%"
pushd PlantUml\common
%COMPRESS_CURDIR%
popd
pushd PlantUml\%EXT%
%COMPRESS_CURDIR%
popd
rename "%TARGET%.zip" "%~nx1"
"c:\Program Files (x86)\VBA Sync Tool\VBASync.exe" -p -a -r -f "%TARGET%" -d vba
