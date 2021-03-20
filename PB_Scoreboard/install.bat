:: Запуск приложений в зависимости от разрядности из разных папок с именами x64 и x86
Set xOS=x64
If "%PROCESSOR_ARCHITECTURE%"=="x86" If Not Defined PROCESSOR_ARCHITEW6432 Set xOS=x86
"drivers\CP210x_Windows_Drivers\CP210xVCPInstaller_%xOS%.exe" /S

If %xOS%==x86 (copy drivers\dx8vb.dll C:\Windows\System32\) Else (copy drivers\dx8vb.dll C:\Windows\SysWOW64\)

If %xOS%==x86 (regsvr32 C:\Windows\System32\dx8vb.dll) Else (regsvr32 C:\Windows\SysWOW64\dx8vb.dll)

drivers\install_trial_scomm32x.exe