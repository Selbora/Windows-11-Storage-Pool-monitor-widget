Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"" -Sta -NoProfile -ExecutionPolicy Bypass -File ""C:\Users\alex\Documents\StoragePoolTray\StoragePoolTray.ps1""", 0
Set WshShell = Nothing
