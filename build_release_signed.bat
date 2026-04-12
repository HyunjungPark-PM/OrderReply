@echo off
setlocal

set VERSION=%1
if "%VERSION%"=="" set VERSION=1.0.12

powershell -NoProfile -ExecutionPolicy Bypass -Command \
  "& '.\build_release_signed.ps1' -Version '%VERSION%' -StoreThumbprint '3F87D815085767D06BCD496D0BFB7D34605AEB73' -StoreScope CurrentUser -EnsureLocalTrust"

if errorlevel 1 (
  echo Signed release build failed.
  exit /b 1
)

echo Signed release build completed.
echo Output: pnet_order_reply_v%VERSION%.zip
endlocal