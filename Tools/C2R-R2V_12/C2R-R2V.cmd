@setlocal DisableDelayedExpansion
@set uivr=v12 AIO
@echo off
:: Licenses used for converting Office 365 ProPlus:
:: set _O365asO2019=0 -> use Office 2016 Mondo (if you want some Office 365 features)
:: set _O365asO2019=1 -> use Office 2019 ProPlus (only for Windows 7 and 8.1)
set _O365asO2019=0

:: set to 1 to enable debug mode
set _Debug=0

:: set to 0 to enable debug mode without cleaning or converting
set _Cnvrt=1

:: ##################################################################

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
set "_ln============================================================="
set "_err===== ERROR ===="
set "_psc=powershell -noprofile -exec bypass -c"
set "xBit=x64"
if /i %PROCESSOR_ARCHITECTURE%==x86 (if not defined PROCESSOR_ARCHITEW6432 (
  set "xBit=x86"
  )
)

if %_Cnvrt% NEQ 1 set _Debug=1

set "_temp=%SystemRoot%\Temp"
reg query HKU\S-1-5-19 1>nul 2>nul || (
set "msg=ERROR: right click on the script and 'Run as administrator'"
goto :TheEnd
)

:Passed
set "_batf=%~f0"
set "_batp=%_batf:'=''%"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%SystemDrive%\Users\Public\Desktop\desktop.ini" set "_dsk=%SystemDrive%\Users\Public\Desktop"
setlocal EnableDelayedExpansion

if %_Debug% EQU 0 (
  set "_Nul1=1>nul"
  set "_Nul2=2>nul"
  set "_Nul6=2^>nul"
  set "_Nul3=1>nul 2>nul"
  call :Begin
) else (
  set "_Nul1="
  set "_Nul2="
  set "_Nul6="
  set "_Nul3="
  copy /y nul "!_work!\#.rw" 1>nul 2>nul && (if exist "!_work!\#.rw" del /f /q "!_work!\#.rw") || (set "_log=!_dsk!\%~n0")
  echo.
  echo Running in Debug Mode...
  echo The window will be closed when finished
  echo.
  echo writing debug log to:
  echo "!_log!_Debug.log"
  @echo on
  @prompt $G
  @call :Begin >"!_log!_tmp.log" 2>&1 &cmd /u /c type "!_log!_tmp.log">"!_log!_Debug.log"&del "!_log!_tmp.log"
)
@color 07
@title %ComSpec%
@echo off
@exit /b

:Begin
color 1F
title Office Click-to-Run Retail-to-Volume %uivr%
for /f "tokens=6 delims=[]. " %%# in ('ver') do set winbuild=%%#
set "_SLMGR=%SysPath%\slmgr.vbs"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
if %_Debug% EQU 0 (
set "_cscript=cscript //Nologo //B"
) else (
set "_cscript=cscript //Nologo"
)

echo %_ln%
echo Running C2R-R2V %uivr%
echo %_ln%

if %winbuild% LSS 7601 (
set "msg=Windows 7 SP1 is the minimum supported OS..."
goto :TheEnd
)
sc query ClickToRunSvc %_Nul3%
set error1=%errorlevel%
sc query OfficeSvc %_Nul3%
set error2=%errorlevel%
if %error1% EQU 1060 if %error2% EQU 1060 (
set "msg=Could not detect Office ClickToRun service..."
goto :TheEnd
)

set _Office16=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
set _Office15=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
if %_Office16% EQU 0 if %_Office15% EQU 0 (
set "msg=No installed Office 2013/2016/2019 product detected..."
goto :TheEnd
)

:Reg16istry
if %_Office16% EQU 0 goto :Reg15istry
set "_InstallRoot="
set "_ProductIds="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if not errorlevel 1 (set "_InstallRoot=%%b\root")
if not "%_InstallRoot%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v PackageGUID" %_Nul6%') do if not errorlevel 1 (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do if not errorlevel 1 (set "_ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
  set "_PRIDs=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"
)
set "_LicensesPath=%_InstallRoot%\Licenses16"
set "_Integrator=%_InstallRoot%\integration\integrator.exe"
for /f "skip=2 tokens=2*" %%a in ('"reg query %_PRIDs% /v ActiveConfiguration" %_Nul6%') do set "_PRIDs=%_PRIDs%\%%b"
if "%_ProductIds%"=="" (
set "msg=Could not detect Office 2016/2019 ProductIDs..."
if %_Office15% EQU 0 (goto :TheEnd) else (goto :Reg15istry)
)
if not exist "%_LicensesPath%\ProPlus*.xrm-ms" (
set "msg=Could not detect Office 2016/2019 Licenses files..."
if %_Office15% EQU 0 (goto :TheEnd) else (goto :Reg15istry)
)
if not exist "%_Integrator%" (
set "msg=Could not detect Office 2016/2019 Licenses Integrator..."
if %_Office15% EQU 0 (goto :TheEnd) else (goto :Reg15istry)
)
if %winbuild% GEQ 10240 set _O365asO2019=0
if exist "%_LicensesPath%\Word2019VL_KMS_Client_AE*.xrm-ms" (set "_tag=2019"&set "_ons= 2019") else (set "_tag="&set "_ons= 2016")
if %_Office15% EQU 0 goto :CheckC2R

:Reg15istry
set "_Install15Root="
set "_Product15Ids="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if not errorlevel 1 (set "_Install15Root=%%b\root")
if not "%_Install15Root%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v PackageGUID" %_Nul6%') do if not errorlevel 1 (set "_GU15ID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do if not errorlevel 1 (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
)
set "_OSPP15ReadT=REG_SZ"
if "%_Product15Ids%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do if not errorlevel 1 (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
)
set "_Licenses15Path=%_Install15Root%\Licenses"
set "_Integ15rator=%_Install15Root%\integration\integrator.exe"
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set "_OSPP15VBS=%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS"
)
if "%_Product15Ids%"=="" (
set "msg=Could not detect Office 2013 ProductIDs..."
if %_Office16% EQU 0 (goto :TheEnd) else (goto :CheckC2R)
)
if not exist "%_Licenses15Path%\ProPlus*.xrm-ms" (
set "msg=Could not detect Office 2013 Licenses files..."
if %_Office16% EQU 0 (goto :TheEnd) else (goto :CheckC2R)
)
if %winbuild% LSS 9200 if not exist "%_OSPP15VBS%" (
set "msg=Could not detect Office 2013 Licensing tool {OSPP.vbs}..."
if %_Office16% EQU 0 (goto :TheEnd) else (goto :CheckC2R)
)

:CheckC2R
set _OMSI=0
if %_Office16% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
)
if %_Office15% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _OMSI=1
)
if %winbuild% GEQ 9200 (
set _spp=SoftwareLicensingProduct
set _sps=SoftwareLicensingService
set "_vbsi=%_SLMGR% /ilc "
) else (
set _spp=OfficeSoftwareProtectionProduct
set _sps=OfficeSoftwareProtectionService
set _vbsi="!_OSPP15VBS!" /inslic:
)
set "_wmi="
for /f "tokens=2 delims==" %%# in ('"wmic path %_sps% get version /value" %_Nul6%') do if not errorlevel 1 set "_wmi=%%#"
if not defined _wmi (
set "msg=Could not execute %_sps% WMI..."
goto :TheEnd
)
echo.
echo %_ln%
echo Checking Office Licenses...
echo %_ln%
wmic path %_spp% where "ApplicationID='%_oApp%' AND Description like '%%KMSCLIENT%%'" get LicenseFamily %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _KMS=1) || (set _KMS=0)
wmic path %_spp% where "ApplicationID='%_oApp%' AND Description like '%%TIMEBASED%%'" get LicenseFamily %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Time=1) || (set _Time=0)
wmic path %_spp% where "ApplicationID='%_oApp%' AND Description like '%%Trial%%'" get LicenseFamily %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Time=1)
wmic path %_spp% where "ApplicationID='%_oApp%' AND Description like '%%Grace%%'" get LicenseFamily %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Grace=1) || (set _Grace=0)
if %_Office16% EQU 1 wmic path %_spp% where "ApplicationID='%_oApp%'" get LicenseFamily %_Nul2% | find /i "Office16MondoVL_KMS_Client" %_Nul1% && (
wmic path %_spp% where "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%'" get LicenseFamily %_Nul2% | find /i "O365" %_Nul1% || (set _Grace=1)
)
if %_Office15% EQU 1 wmic path %_spp% where "ApplicationID='%_oApp%'" get LicenseFamily %_Nul2% | find /i "OfficeMondoVL_KMS_Client" %_Nul1% && (
wmic path %_spp% where "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%'" get LicenseFamily %_Nul2% | find /i "O365" %_Nul1% || (set _Grace=1)
)
if %_Time% EQU 0 if %_Grace% EQU 0 if %_KMS% EQU 1 (
set "msg=No Conversion or Cleanup Required..."
goto :TheEnd
)

set _Retail=0
wmic path %_spp% where "ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey<>NULL" get Description %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
find /i "RETAIL channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "RETAIL(MAK) channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "TIMEBASED_SUB channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
set "_copp="
if exist "%SysPath%\msvcr100.dll" (
set _copp=1
) else if exist "!_InstallRoot!\vfs\System\msvcr100.dll" (
set _copp="!_InstallRoot!\vfs\System"
) else if exist "!_Install15Root!\vfs\System\msvcr100.dll" (
set _copp="!_Install15Root!\vfs\System"
) else if exist "%SystemRoot%\SysWOW64\msvcr100.dll" (
set _copp=1
set xBit=x86
) else if exist "!_InstallRoot!\vfs\SystemX86\msvcr100.dll" (
set _copp="!_InstallRoot!\vfs\SystemX86"
set xBit=x86
) else if exist "!_Install15Root!\vfs\SystemX86\msvcr100.dll" (
set _copp="!_Install15Root!\vfs\SystemX86"
set xBit=x86
)
if %_Cnvrt% EQU 1 if %_Retail% EQU 0 if %_OMSI% EQU 0 if defined _copp (
echo.
echo %_ln%
echo Cleaning Current Office Licenses...
echo %_ln%
if not exist "!_work!\bin\cleanospp%xBit%.exe" (
echo.
echo ERROR: could not find cleanospp%xBit%.exe
goto :PassCopp
)
if "!_copp!"=="1" (
%_Nul3% "!_work!\bin\cleanospp%xBit%.exe" -Licenses
goto :PassCopp
)
pushd %_copp%
%_Nul3% copy /y "!_work!\bin\cleanospp%xBit%.exe" cleanospp.exe
%_Nul3% cleanospp.exe -Licenses
%_Nul3% del /f /q cleanospp.exe
popd
)
:PassCopp
echo.
echo %_ln%
echo Installing Office Volume Licenses...
echo %_ln%
echo.
set _O16O365=0
if %_Retail% EQU 1 wmic path %_spp% where "ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey<>NULL" get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
wmic path %_spp% where "ApplicationID='%_oApp%'" get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvVolume.txt" 2>&1

if %_Office16% EQU 0 goto :R15V

set _O19Ids=ProPlus2019,ProjectPro2019,VisioPro2019,Standard2019,ProjectStd2019,VisioStd2019,Access2019,SkypeforBusiness2019
set _O16Ids=ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness
set _A19Ids=Excel2019,Outlook2019,PowerPoint2019,Publisher2019,Word2019
set _A16Ids=Excel,Outlook,PowerPoint,Publisher,Word
set _V19Ids=%_O19Ids%,%_A19Ids%
set _V16Ids=Mondo,%_O16Ids%,%_A16Ids%,OneNote
set _R16Ids=%_V16Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set _RetIds=%_V19Ids%,Professional2019,HomeBusiness2019,HomeStudent2019,%_R16Ids%

echo %_ProductIds%>"!_temp!\crvProductIds.txt"
for %%a in (%_RetIds%,ProPlus) do (
set _%%a=0
)
for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V19Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office19%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
for %%a in (%_V16Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aE5R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aEDUR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aXC2RVL_MAKC2R" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  )
)
if %_Retail% EQU 1 reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
)

if %_Cnvrt% NEQ 1 (if %_Office15% EQU 1 (goto :R15V) else (set "msg=Finished"&goto :TheEnd))

set _C16Msg=0
for %%a in (%_RetIds%,ProPlus) do if !_%%a! EQU 1 (
set _C16Msg=1
)
if %_C16Msg% EQU 0 (if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R))

if !_Mondo! EQU 1 (
call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 (
  if !_O365asO2019! EQU 1 (
  if !_Mondo! EQU 0 (
  echo O365ProPlus 2016 Suite -^> ProPlus%_ons% Licenses
  echo.
  call :InsLic ProPlus%_tag%
  )
  ) else (
  echo O365ProPlus 2016 Suite ^<-^> Mondo 2016 Licenses
  echo.
  call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
  if !_Mondo! EQU 0 call :InsLic Mondo
  )
)
if !_O365Business! EQU 1 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365Business 2016 Suite ^<-^> Mondo 2016 Licenses
echo.
call :InsLic O365Business NCHRJ-3VPGW-X73DM-6B36K-3RQ6B
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2016 Suite ^<-^> Mondo 2016 Licenses
echo.
call :InsLic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365HomePrem! EQU 1 if !_O365SmallBusPrem! EQU 0 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365HomePrem 2016 Suite ^<-^> Mondo 2016 Licenses
echo.
call :InsLic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365EduCloud! EQU 1 if !_O365HomePrem! EQU 0 if !_O365SmallBusPrem! EQU 0 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365EduCloud 2016 Suite ^<-^> Mondo 2016 Licenses
echo.
call :InsLic O365EduCloud 8843N-BCXXD-Q84H8-R4Q37-T3CPT
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 set _O16O365=1
if !_Mondo! EQU 1 if !_O365ProPlus! EQU 0 (
echo Mondo 2016 Suite
echo.
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R)
)
if !_ProPlus2019! EQU 1 if !_O365ProPlus! EQU 0 (
echo ProPlus 2019 Suite
echo.
call :InsLic ProPlus%_tag%
)
if !_ProPlus! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 (
echo ProPlus 2016 Suite -^> ProPlus%_ons% Licenses
echo.
call :InsLic ProPlus%_tag%
)
if !_Professional2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 (
echo Professional 2019 Suite -^> ProPlus%_ons% Licenses
echo.
call :InsLic ProPlus%_tag%
)
if !_Professional! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 (
echo Professional 2016 Suite -^> ProPlus%_ons% Licenses
echo.
call :InsLic ProPlus%_tag%
)
if !_Standard2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 (
echo Standard 2019 Suite
echo.
call :InsLic Standard2019
)
if !_Standard! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 (
echo Standard 2016 Suite -^> Standard%_ons% Licenses
echo.
call :InsLic Standard%_tag%
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a2019! EQU 1 (
echo %%a 2019 SKU
echo.
if defined _tag (call :InsLic %%a2019) else (call :InsLic %%a)
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! EQU 1 (
if !_%%a2019! EQU 0 (
  echo %%a 2016 SKU -^> %%a%_ons% Licenses
  echo.
  call :InsLic %%a%_tag%
  )
)
for %%a in (HomeBusiness2019,HomeStudent2019) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 (
  set _Standard2019=1
  echo %%a Suite -^> Standard 2019 Licenses
  echo.
  call :InsLic Standard2019
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2019! EQU 0 (
  set _Standard2019=1
  echo %%a 2016 Suite -^> Standard%_ons% Licenses
  echo.
  call :InsLic Standard%_tag%
  )
)
for %%a in (%_A19Ids%,OneNote) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 (
  echo %%a App
  echo.
  call :InsLic %%a
  )
)
for %%a in (%_A16Ids%) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App
  echo.
  call :InsLic %%a%_tag%
  )
)
for %%a in (Access2019) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 (
  echo %%a App
  echo.
  call :InsLic %%a
  )
)
for %%a in (Access) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App
  echo.
  call :InsLic %%a%_tag%
  )
)
for %%a in (SkypeforBusiness2019) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 (
  echo %%a App
  echo.
  call :InsLic %%a
  )
)
for %%a in (SkypeforBusiness) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App
  echo.
  call :InsLic %%a%_tag%
  )
)
if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R)

:R15V
if %_Cnvrt% EQU 1 (
for %%# in ("!_Licenses15Path!\client-issuance-*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
%_cscript% %_vbsi%"!_Licenses15Path!\pkeyconfig-office.xrm-ms"
)

set _O15Ids=Standard,ProjectPro,VisioPro,ProjectStd,VisioStd,Access,Lync
set _A15Ids=Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set _R15Ids=SPD,Mondo,%_O15Ids%,%_A15Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem
set _V15Ids=Mondo,%_O15Ids%,%_A15Ids%

echo %_Product15Ids%>"!_temp!\crvProduct15s.txt"
for %%a in (%_R15Ids%,ProPlus) do (
set _%%a=0
)
for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V15Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PR15IDs%\Active\ProPlusVolume\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  )
)
if %_Retail% EQU 1 reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
)

if %_Cnvrt% NEQ 1 (set "msg=Finished"&goto :TheEnd)

set _C15Msg=0
for %%a in (%_R15Ids%,ProPlus) do if !_%%a! EQU 1 (
set _C15Msg=1
)
if %_C15Msg% EQU 0 goto :GVLKC2R

if !_Mondo! EQU 1 (
call :Ins15Lic Mondo
)
if !_O365ProPlus! EQU 1 if !_O16O365! EQU 0 (
echo O365ProPlus 2013 Suite ^<-^> Mondo 2013 Licenses
echo.
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2013 Suite ^<-^> Mondo 2013 Licenses
echo.
call :Ins15Lic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365HomePrem! EQU 1 if !_O365SmallBusPrem! EQU 0 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365HomePrem 2013 Suite ^<-^> Mondo 2013 Licenses
echo.
call :Ins15Lic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365Business! EQU 1 if !_O365HomePrem! EQU 0 if !_O365SmallBusPrem! EQU 0 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365Business 2013 Suite ^<-^> Mondo 2013 Licenses
echo.
call :Ins15Lic O365Business MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_Mondo! EQU 1 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
echo Mondo 2013 Suite
echo.
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :GVLKC2R
)
if !_SPD! EQU 1 if !_Mondo! EQU 0 if !_O365ProPlus! EQU 0 (
echo SharePoint Designer 2013 App -^> Mondo 2013 Licenses
echo.
call :Ins15Lic Mondo
goto :GVLKC2R
)
if !_ProPlus! EQU 1 if !_O365ProPlus! EQU 0 (
echo ProPlus 2013 Suite
echo.
call :Ins15Lic ProPlus
)
if !_Professional! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 (
echo Professional 2013 Suite -^> ProPlus 2013 Licenses
echo.
call :Ins15Lic ProPlus
)
if !_Standard! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 (
echo Standard 2013 Suite
echo.
call :Ins15Lic Standard
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! EQU 1 (
echo %%a 2013 SKU
echo.
call :Ins15Lic %%a
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 if !_Standard! EQU 0 (
  set _Standard=1
  echo %%a 2013 Suite -^> Standard 2013 Licenses
  echo.
  call :Ins15Lic Standard
  )
)
for %%a in (%_A15Ids%) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 if !_Standard! EQU 0 (
  echo %%a 2013 App
  echo.
  call :Ins15Lic %%a
  )
)
for %%a in (Access) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 (
  echo %%a 2013 App
  echo.
  call :Ins15Lic %%a
  )
)
for %%a in (Lync) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 (
  echo SkypeforBusiness 2015 App
  echo.
  call :Ins15Lic %%a
  )
)
goto :GVLKC2R

:InsLic
set "_ID=%1Volume"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_pkey=PidKey=%2"
)
reg delete %_Config% /f /v %_ID%.OSPPReady %_Nul3%
"!_Integrator!" /I /License PRIDName=%_ID%.16 %_pkey% PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!" %_Nul1%
reg add %_Config% /f /v %_ID%.OSPPReady /t REG_SZ /d 1 %_Nul1%
reg query %_Config% /v ProductReleaseIds | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config% /v ProductReleaseIds') do reg add %_Config% /v ProductReleaseIds /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:Ins15Lic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=%2"
)
reg delete %_OSPP15Ready% /f /v %_ID%.OSPPReady %_Nul3%
for %%# in ("!_Licenses15Path!\%_patt%*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
if defined _pkey wmic path %_sps% where version='%_wmi%' call InstallProductKey ProductKey="%_pkey%" %_Nul3%
reg add %_OSPP15Ready% /f /v %_ID%.OSPPReady /t %_OSPP15ReadT% /d 1 %_Nul1%
reg query %_Con15fig% | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Con15fig%') do reg add %_Con15fig% /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:GVLKC2R
echo %_ln%
echo Installing Missing KMS Client Keys...
echo %_ln%
echo.
if %winbuild% GEQ 9200 wmic path %_sps% where version='%_wmi%' call RefreshLicenseStatus %_Nul3%
for %%# in (15,16,19) do call :C2RLoc %%#
for %%# in (15,16,19) do if !_Office%%#! EQU 0 call :C2Runi %%#
for %%# in (15,16,19) do if !_Office%%#! EQU 1 call :C2Rins %%#
if %winbuild% GEQ 9200 wmic path %_sps% where version='%_wmi%' call RefreshLicenseStatus %_Nul3%
if exist "%SysPath%\spp\store_test\2.0\tokens.dat" (
echo.
echo %_ln%
echo Refreshing Windows Insider Preview Licenses...
echo %_ln%
echo.
%_cscript% %_SLMGR% /rilc
)
set "msg=Finished"
goto :TheEnd

:C2Runi
for /f "tokens=2 delims==" %%# in ('wmic path %_spp% where "Name like 'Office %~1%%' AND PartialProductKey<>NULL" get ID /value %_Nul6%') do (set "aID=%%#"&call :UniKey)
exit /b

:C2Rins
for /f "tokens=2 delims==" %%# in ('"wmic path %_spp% where (Description like 'Office %1, VOLUME_KMSCLIENT%%' AND PartialProductKey=NULL) get ID /value" %_Nul6%') do (set "aID=%%#"&call :InsKey)
exit /b

:C2RLoc
set _Office%1=0
if %1 EQU 19 (
if defined _ProductIds reg query %_Config% /v ProductReleaseIds %_Nul2% | findstr 2019 %_Nul1% && set _Office%1=1
exit /b
)

for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _Office%1=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _Office%1=1

if %1 EQU 16 if defined _ProductIds (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config% /v ProductReleaseIds') do echo %%b>"!_temp!\crvO16.txt"
for %%a in (%_R16Ids%) do (
  findstr /I /C:"%%aRetail" "!_temp!\crvO16.txt" %_Nul1% && set _Office%1=1
  )
for %%a in (%_V16Ids%,ProjectProX,ProjectStdX,VisioProX,VisioStdX) do (
  findstr /I /C:"%%aVolume" "!_temp!\crvO16.txt" %_Nul1% && set _Office%1=1
  )
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && set _Office%1=1
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && set _Office%1=1
exit /b
)

if %1 EQU 15 if defined _Product15Ids (
set _Office%1=1
exit /b
)

if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" set _Office%1=1
if exist "%ProgramW6432%\Microsoft Office\Office%1\OSPP.VBS" set _Office%1=1
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" set _Office%1=1
exit /b

:UniKey
wmic path %_spp% where ID='%aID%' call UninstallProductKey %_Nul3%
exit /b

:InsKey
if /i '%aID%' EQU '1dc00701-03af-4680-b2af-007ffc758a1f' exit /b
if /i '%aID%' EQU 'e914ea6e-a5fa-4439-a394-a9bb3293ca09' exit /b
if /i '%aID%' EQU '0bc88885-718c-491d-921f-6f214349e79c' exit /b
if /i '%aID%' EQU 'fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9' exit /b
if /i '%aID%' EQU '500f6619-ef93-4b75-bcb4-82819998a3ca' exit /b
set "_key="
for /f "tokens=2 delims==" %%# in ('"wmic path %_spp% where ID='%aID%' get LicenseFamily /value"') do echo %%#
call :keys %aID%
if "%_key%"=="" (echo Could not find matching key&echo.&exit /b)
wmic path %_sps% where version='%_wmi%' call InstallProductKey ProductKey="%_key%" %_Nul3%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 (
cmd /c exit /b %ERRORCODE%
echo Failed: 0x!=ExitCode!
)
echo.
exit /b

:keys
if "%~1"=="" exit /b
goto :%1 %_Nul2%

:: Office 2019
:85dd8b5f-eaa4-4af3-a628-cce9e77c9a03
set "_key=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP" &:: Professional Plus
exit /b

:6912a74b-a5fb-401a-bfdb-2e3ab46f4b02
set "_key=6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK" &:: Standard
exit /b

:2ca2bf3f-949e-446a-82c7-e25a15ec78c4
set "_key=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B" &:: Project Professional
exit /b

:1777f0e3-7392-4198-97ea-8ae4de6f6381
set "_key=C4F7P-NCP8C-6CQPT-MQHV9-JXD2M" &:: Project Standard
exit /b

:5b5cf08f-b81a-431d-b080-3450d8620565
set "_key=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB" &:: Visio Professional
exit /b

:e06d7df3-aad0-419d-8dfb-0ac37e2bdf39
set "_key=7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2" &:: Visio Standard
exit /b

:9e9bceeb-e736-4f26-88de-763f87dcc485
set "_key=9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT" &:: Access
exit /b

:237854e9-79fc-4497-a0c1-a70969691c6b
set "_key=TMJWT-YYNMB-3BKTF-644FC-RVXBD" &:: Excel
exit /b

:c8f8a301-19f5-4132-96ce-2de9d4adbd33
set "_key=7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK" &:: Outlook
exit /b

:3131fd61-5e4f-4308-8d6d-62be1987c92c
set "_key=RRNCX-C64HY-W2MM7-MCH9G-TJHMQ" &:: PowerPoint
exit /b

:9d3e4cca-e172-46f1-a2f4-1d2107051444
set "_key=G2KWX-3NW6P-PY93R-JXK2T-C9Y9V" &:: Publisher
exit /b

:734c6c6e-b0ba-4298-a891-671772b2bd1b
set "_key=NCJ33-JHBBY-HTK98-MYCV8-HMKHJ" &:: Skype for Business
exit /b

:059834fe-a8ea-4bff-b67b-4d006b5447d3
set "_key=PBX3G-NWMT6-Q7XBW-PYJGG-WXD33" &:: Word
exit /b

:: Office 2016
:829b8110-0e6f-4349-bca4-42803577788d
set "_key=WGT24-HCNMF-FQ7XH-6M8K7-DRTW9" &:: Project Professional C2R-P
exit /b

:cbbaca45-556a-4416-ad03-bda598eaa7c8
set "_key=D8NRQ-JTYM3-7J2DX-646CT-6836M" &:: Project Standard C2R-P
exit /b

:b234abe3-0857-4f9c-b05a-4dc314f85557
set "_key=69WXN-MBYV6-22PQG-3WGHK-RM6XC" &:: Visio Professional C2R-P
exit /b

:361fe620-64f4-41b5-ba77-84f8e079b1f7
set "_key=NY48V-PPYYH-3F4PX-XJRKJ-W4423" &:: Visio Standard C2R-P
exit /b

:e914ea6e-a5fa-4439-a394-a9bb3293ca09
set "_key=DMTCJ-KNRKX-26982-JYCKT-P7KB6" &:: MondoR
exit /b

:9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
set "_key=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2" &:: Mondo
exit /b

:d450596f-894d-49e0-966a-fd39ed4c4c64
set "_key=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99" &:: Professional Plus
exit /b

:dedfa23d-6ed1-45a6-85dc-63cae0546de6
set "_key=JNRGM-WHDWX-FJJG3-K47QV-DRTFM" &:: Standard
exit /b

:4f414197-0fc2-4c01-b68a-86cbb9ac254c
set "_key=YG9NW-3K39V-2T3HJ-93F3Q-G83KT" &:: Project Professional
exit /b

:da7ddabc-3fbe-4447-9e01-6ab7440b4cd4
set "_key=GNFHQ-F6YQM-KQDGJ-327XX-KQBVC" &:: Project Standard
exit /b

:6bf301c1-b94a-43e9-ba31-d494598c47fb
set "_key=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK" &:: Visio Professional
exit /b

:aa2a7821-1827-4c2c-8f1d-4513a34dda97
set "_key=7WHWN-4T7MP-G96JF-G33KR-W8GF4" &:: Visio Standard
exit /b

:67c0fc0c-deba-401b-bf8b-9c8ad8395804
set "_key=GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW" &:: Access
exit /b

:c3e65d36-141f-4d2f-a303-a842ee756a29
set "_key=9C2PK-NWTVB-JMPW8-BFT28-7FTBF" &:: Excel
exit /b

:d8cace59-33d2-4ac7-9b1b-9b72339c51c8
set "_key=DR92N-9HTF2-97XKM-XW2WJ-XW3J6" &:: OneNote
exit /b

:ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
set "_key=R69KK-NTPKF-7M3Q4-QYBHW-6MT9B" &:: Outlook
exit /b

:d70b1bba-b893-4544-96e2-b7a318091c33
set "_key=J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6" &:: Powerpoint
exit /b

:041a06cb-c5b8-4772-809f-416d03d16654
set "_key=F47MM-N3XJP-TQXJ9-BP99D-8K837" &:: Publisher
exit /b

:83e04ee1-fa8d-436d-8994-d31a862cab77
set "_key=869NQ-FJ69K-466HW-QYCP2-DDBV6" &:: Skype for Business
exit /b

:bb11badf-d8aa-470e-9311-20eaf80fe5cc
set "_key=WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6" &:: Word
exit /b

:: Office 2013
:dc981c6b-fc8e-420f-aa43-f8f33e5c0923
set "_key=42QTK-RN8M7-J3C4G-BBGYM-88CYV" &:: Mondo
exit /b

:b322da9c-a2e2-4058-9e4e-f59a6970bd69
set "_key=YC7DK-G2NP3-2QQC3-J6H88-GVGXT" &:: Professional Plus
exit /b

:b13afb38-cd79-4ae5-9f7f-eed058d750ca
set "_key=KBKQT-2NMXY-JJWGP-M62JB-92CD4" &:: Standard
exit /b

:4a5d124a-e620-44ba-b6ff-658961b33b9a
set "_key=FN8TT-7WMH6-2D4X9-M337T-2342K" &:: Project Professional
exit /b

:427a28d1-d17c-4abf-b717-32c780ba6f07
set "_key=6NTH3-CW976-3G3Y2-JK3TX-8QHTT" &:: Project Standard
exit /b

:e13ac10e-75d0-4aff-a0cd-764982cf541c
set "_key=C2FG9-N6J68-H8BTJ-BW3QX-RM3B3" &:: Visio Professional
exit /b

:ac4efaf0-f81f-4f61-bdf7-ea32b02ab117
set "_key=J484Y-4NKBF-W2HMG-DBMJC-PGWR7" &:: Visio Standard
exit /b

:6ee7622c-18d8-4005-9fb7-92db644a279b
set "_key=NG2JY-H4JBT-HQXYP-78QH9-4JM2D" &:: Access
exit /b

:f7461d52-7c2b-43b2-8744-ea958e0bd09a
set "_key=VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB" &:: Excel
exit /b

:fb4875ec-0c6b-450f-b82b-ab57d8d1677f
set "_key=H7R7V-WPNXQ-WCYYC-76BGV-VT7GH" &:: Groove
exit /b

:a30b8040-d68a-423f-b0b5-9ce292ea5a8f
set "_key=DKT8B-N7VXH-D963P-Q4PHY-F8894" &:: InfoPath
exit /b

:1b9f11e3-c85c-4e1b-bb29-879ad2c909e3
set "_key=2MG3G-3BNTT-3MFW9-KDQW3-TCK7R" &:: Lync
exit /b

:efe1f3e6-aea2-4144-a208-32aa872b6545
set "_key=TGN6P-8MMBC-37P2F-XHXXK-P34VW" &:: OneNote
exit /b

:771c3afa-50c5-443f-b151-ff2546d863a0
set "_key=QPN8Q-BJBTJ-334K3-93TGY-2PMBT" &:: Outlook
exit /b

:8c762649-97d1-4953-ad27-b7e2c25b972e
set "_key=4NT99-8RJFH-Q2VDH-KYG2C-4RD4F" &:: Powerpoint
exit /b

:00c79ff1-6850-443d-bf61-71cde0de305f
set "_key=PN2WF-29XG2-T9HJ7-JQPJR-FCXK4" &:: Publisher
exit /b

:d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3
set "_key=6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7" &:: Word
exit /b

:TheEnd
if exist "%_temp%\crv*.txt" del /f /q "%_temp%\crv*.txt"
echo.
echo %_ln%
echo %msg%
echo %_ln%
echo.
echo Press any key to exit...
if %_Debug% EQU 1 goto :eof
pause >nul
goto :eof
