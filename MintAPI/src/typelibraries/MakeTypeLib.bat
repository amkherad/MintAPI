echo off

echo.
echo by Ali Mousavi Kherad alimousavikherad@gmail.com





echo.
echo ===============================================
echo   @mintcore.tlb   Mint Core
echo             -----------------------
mktyplib /nologo MintCore/MintCore.odl /tlb ../../bin/mintcore.tlb /h MintCore/pack.h /I MintCore/include
if errorlevel 1 goto errorOccured
regtlibv12 "C:\Documents and Settings\Ali Mousavi Kherad\My Documents\Projects\MintAPI\MintAPI/bin/mintcore.tlb"
echo             -----------------------
echo SUCCESSFUL: @mintcore.tlb   Mint Core
echo =================================================
echo.





echo.
echo =================================================
echo   @mintcml.tlb   Mint Cross Module Linkage
echo             -----------------------
mktyplib /nologo MintCrossModuleLinkage/MintCrossModule.odl /tlb ../../bin/mintcml.tlb /h MintCrossModuleLinkage/pack.h /I MintCrossModuleLinkage/include
if errorlevel 1 goto errorOccured
regtlibv12 "C:\Documents and Settings\Ali Mousavi Kherad\My Documents\Projects\MintAPI\MintAPI/bin/mintcml.tlb"
echo             -----------------------
echo SUCCESSFUL: @mintcml.tlb   Mint Cross Module Linkage
echo =================================================
echo.





echo.
echo =================================================
echo   @mintkbh.tlb   Mint Kernel Back Holder
echo             -----------------------
mktyplib /nologo Modules/Kernel/MintAPIKernelBackHolder.odl /tlb ../../bin/mintkbh.tlb /h Modules/Kernel/pack.h /I Modules/Kernel/include
if errorlevel 1 goto errorOccured
regtlibv12 "C:\Documents and Settings\Ali Mousavi Kherad\My Documents\Projects\MintAPI\MintAPI/bin/mintkbh.tlb"
echo             -----------------------
echo SUCCESSFUL: @mintkbh.tlb   Mint Kernel Back Holder
echo =================================================
echo.






goto end

:errorOccured
echo             -----------------------
echo BUILD FAILED  XXXXXXXXXXXXXXXX
echo =================================================
echo.
pause

:end