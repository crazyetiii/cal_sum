@ECHO OFF&(PUSHD "%~DP0")&(REG QUERY "HKU\S-1-5-19">NUL 2>&1)||(
powershell -Command "Start-Process '%~sdpnx0' -Verb RunAs"&&EXIT)

:MENU
ECHO.&ECHO  1�������Դ�������Ҽ��˵���
ECHO.&ECHO  2���Ƴ���Դ�������Ҽ��˵���
CHOICE /C 123 /N >NUL 2>NUL
IF "%ERRORLEVEL%"=="2" GOTO RemoveMenu
IF "%ERRORLEVEL%"=="1" GOTO AddMenu

:AddMenu
reg add "HKCU\Software\Scooter Software\Beyond Compare" /f /v "ExePath" /d "\"%~dp0BCompare.exe\"" >NUL
reg add "HKCU\SOFTWARE\Scooter Software\Beyond Compare 5" /f /v "ExePath" /d "\"%~dp0BCompare.exe\"" >NUL
reg add "HKCU\SOFTWARE\Scooter Software\Beyond Compare 5\BcShellEx" /f /v "SavedLeft" /d "\"%~dp0BCompare.exe\"" >NUL
reg add "HKLM\SOFTWARE\WOW6432Node\Scooter Software\Beyond Compare" /f /v "ExePath" /d "\"%~dp0BCompare.exe\"" >NUL
reg add "HKLM\SOFTWARE\WOW6432Node\Scooter Software\Beyond Compare 5" /f /v "ExePath" /d "\"%~dp0BCompare.exe\"" >NUL
reg add "HKLM\SOFTWARE\WOW6432Node\Scooter Software\Beyond Compare 5\BcShellEx" /f /v "SavedLeft" /d "\"%~dp0BCompare.exe\"" >NUL
reg add "HKLM\SOFTWARE\Classes\CLSID\{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}\InProcServer32" /f /ve /d "\"%~dp0BCShellEx64.dll\"" >NUL
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\BCompare.exe" /f /ve /d "\"%~dp0BCompare.exe\"" >NUL
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\BCompare.exe" /f /v "UseURL" /t REG_DWORD /d "1" >NUL
reg add "HKLM\SOFTWARE\Classes\.bcss" /f /ve /d "BeyondCompare.Snapshot" >NUL
reg add "HKLM\SOFTWARE\Classes\BeyondCompare.Snapshot" /f /ve /d "Beyond Compare Snapshot" >NUL
reg add "HKLM\SOFTWARE\Classes\BeyondCompare.Snapshot\DefaultIcon" /f /ve /d "\"%~dp0BCompare.exe,0\"" >NUL
reg add "HKLM\SOFTWARE\Classes\BeyondCompare.Snapshot\shell\open\command" /f /ve /d "\"%~dp0BCompare.exe\" \"%%1\"" >NUL
reg add "HKLM\SOFTWARE\Classes\.bcpkg" /f /ve /d "BeyondCompare.SettingsPackage" >NUL
reg add "HKLM\SOFTWARE\Classes\BeyondCompare.SettingsPackage" /f /ve /d "Beyond Compare Settings Package" >NUL
reg add "HKLM\SOFTWARE\Classes\BeyondCompare.SettingsPackage" /f /v "EditFlags" /t REG_DWORD /d "0x00100000" >NUL
reg add "HKLM\SOFTWARE\Classes\BeyondCompare.SettingsPackage\DefaultIcon" /f /ve /d "\"%~dp0BCompare.exe,0\"" >NUL
reg add "HKLM\SOFTWARE\Classes\BeyondCompare.SettingsPackage\shell\open\command" /f /ve /d "\"%~dp0BCompare.exe\" \"%%1\"" >NUL
reg add "HKLM\SOFTWARE\Classes\CLSID\{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}" /f /ve /d "CirrusShellEx" >NUL
reg add "HKLM\SOFTWARE\Classes\CLSID\{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}\InProcServer32" /f /ve /d "\"%~dp0BCShellEx64.dll\"" >NUL
reg add "HKLM\SOFTWARE\Classes\CLSID\{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}\InProcServer32" /f /v "ThreadingModel" /d "Apartment" >NUL
reg add "HKLM\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\CirrusShellEx" /f /ve /d "{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}" >NUL
reg add "HKLM\SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\CirrusShellEx" /f /ve /d "{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}" >NUL
reg add "HKLM\SOFTWARE\Classes\lnkfile\shellex\ContextMenuHandlers\CirrusShellEx" /f /ve /d "{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}" >NUL
reg add "HKLM\SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\CirrusShellEx" /f /ve /d "{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}" >NUL
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" /f /v "{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}" /d "Beyond Compare 5 Shell Extension" >NUL
reg add "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\Beyond Compare 5" /f /v "TypesSupported" /t REG_DWORD /d "7" >NUL 
reg add "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\Beyond Compare 5" /f /v "EventMessageFile" /d "\"%~dp0BCompare.exe\"" >NUL
ECHO.&ECHO ������ &TIMEOUT /t 3 >NUL&CLS&GOTO MENU

:RemoveMenu
reg delete "HKLM\SOFTWARE\Classes\.bcss" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\Classes\.bcpkg" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\Classes\BeyondCompare.Snapshot" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\Classes\BeyondCompare.SettingsPackage" /F>NUL 2>NUL
reg delete "HKCU\SOFTWARE\Scooter Software\Beyond Compare" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\Scooter Software\Beyond Compare" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\Scooter Software\Beyond Compare 5" /F>NUL 2>NUL
reg delete "HKCU\SOFTWARE\Scooter Software\Beyond Compare 5" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\WOW6432Node\Scooter Software\Beyond Compare" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\WOW6432Node\Scooter Software\Beyond Compare 5" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\BCompare.exe" /F >NUL 2>NUL
reg delete "HKLM\SOFTWARE\Classes\CLSID\{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}" /F >NUL 2>NUL
reg delete "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\Beyond Compare 5" /F>NUL 2>NUL
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved" /F /v "{812BC6B5-83CF-4AD9-97C1-6C60C8D025C5}">NUL 2>NUL

ECHO.&ECHO �Ƴ����
ECHO.&ECHO ghxi.com
TIMEOUT /t 5 >NUL&CLS&GOTO MENU