rem older reg modification
rem REG ADD "HKCU\Software\Microsoft\Command Processor" /V DisableUNCCheck /T REG_DWORD /F /D 1


rem do copy
copy .\eunzip\bin\Release\eUnzip.exe %APPDATA%\Microsoft\Windows\SendTo\
Pause