reg add "HKCU\Control Panel\International" /v sShortDate /t REG_SZ /d dd/MM/yyyy /f
rundll32.exe user32.dll,LockWorkStation