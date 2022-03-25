set TARGET=%1
set EXT=%TARGET:~-4%

del %TARGET%
pushd PlantUml\common
zip -r ..\..\%TARGET% .
popd
pushd PlantUml\%EXT%
zip -r ..\..\%TARGET% .
popd
"c:\Program Files (x86)\VBA Sync Tool\VBASync.exe" -p -a -r -f %TARGET% -d vba
