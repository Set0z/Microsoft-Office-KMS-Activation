@echo off
echo Activation...

:: Selecting the correct folder if exists
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16" (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
) else if exist "%ProgramFiles%\Microsoft Office\Office16" (
    cd /d "%ProgramFiles%\Microsoft Office\Office16"
) else (
    echo Could not find folder Microsoft Office\Office16.
    pause
    exit /b
)

:: Installing a license
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do (
    cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
)

:: Setting up KMS
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6F7TH >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:e8.us.to
cscript ospp.vbs /act

reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v UpdatesEnabled /t REG_SZ /d False /f

cls
echo Done!!!

pause
