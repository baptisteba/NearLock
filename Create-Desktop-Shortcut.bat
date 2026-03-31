@echo off
powershell.exe -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
   $sc = $ws.CreateShortcut([System.IO.Path]::Combine($env:USERPROFILE, 'Desktop', 'NearLock.lnk')); ^
   $sc.TargetPath  = '%~dp0NearLock.bat'; ^
   $sc.WorkingDirectory = '%~dp0'; ^
   $sc.IconLocation = '%~dp0NearLock.ico, 0'; ^
   $sc.Description = 'NearLock - Bluetooth auto-lock'; ^
   $sc.WindowStyle = 7; ^
   $sc.Save()"
echo Raccourci cree sur le Bureau.
pause
