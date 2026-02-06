$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BinDir = Join-Path $ScriptDir "bin"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Building CMDT - Run as TrustedInstaller" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

$VSBASE  = "C:\Program Files\Microsoft Visual Studio\18\Enterprise\VC\Tools\MSVC\14.50.35717\bin\Hostx64"
$ML32    = "$VSBASE\x86\ml.exe"
$ML64    = "$VSBASE\x64\ml64.exe"
$LINK32  = "$VSBASE\x86\link.exe"
$LINK64  = "$VSBASE\x64\link.exe"

$SDKBASE       = "C:\Program Files (x86)\Windows Kits\10\Lib\10.0.22621.0"
$SDKBIN        = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64"
$SDKINCLUDE    = "C:\Program Files (x86)\Windows Kits\10\Include\10.0.22621.0"
$LIBPATH32_UM  = "$SDKBASE\um\x86"
$LIBPATH32_UCRT = "$SDKBASE\ucrt\x86"
$LIBPATH64_UM  = "$SDKBASE\um\x64"
$LIBPATH64_UCRT = "$SDKBASE\ucrt\x64"

$env:PATH += ";$SDKBIN"
$env:INCLUDE = "$SDKINCLUDE\um;$SDKINCLUDE\shared;$SDKINCLUDE\ucrt"

if (-not (Test-Path $BinDir)) {
    New-Item -ItemType Directory -Path $BinDir | Out-Null
}

$FILES = @("main", "token", "process", "window")
$LIBS = @("kernel32.lib", "user32.lib", "advapi32.lib", "shlwapi.lib", "shell32.lib", "gdi32.lib", "comdlg32.lib", "userenv.lib", "ole32.lib")
$BuildSuccess = $true

Write-Host ""
Write-Host ">>> Architecture: x86" -ForegroundColor Cyan
Push-Location $ScriptDir
& rc /c65001 /fo cmdt_x86.res cmdt.rc
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Resource compilation failed" -ForegroundColor Red
    $BuildSuccess = $false
} else {
    foreach ($f in $FILES) {
        & $ML32 /c /Cp /Cx /Zi /I x86 /Fo"x86\$f.obj" "x86\$f.asm"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Assembly of $f.asm failed" -ForegroundColor Red
            $BuildSuccess = $false
            break
        }
    }
    if ($BuildSuccess) {
        $linkArgs = @("x86\main.obj", "x86\token.obj", "x86\process.obj", "x86\window.obj", "cmdt_x86.res", "/subsystem:windows", "/entry:start", "/out:bin\cmdt_x86.exe", "/MANIFEST:EMBED", "/MANIFESTINPUT:cmdt.manifest", "/LIBPATH:$LIBPATH32_UM", "/LIBPATH:$LIBPATH32_UCRT") + $LIBS
        & $LINK32 $linkArgs
        if ($LASTEXITCODE -ne 0) { 
            $BuildSuccess = $false 
        } else { 
            Write-Host "Build successful: bin\cmdt_x86.exe" -ForegroundColor Green
            Write-Host "Checking imports..." -ForegroundColor Cyan
            $DUMPBIN32 = "$VSBASE\x86\dumpbin.exe"
            & $DUMPBIN32 /imports "$BinDir\cmdt_x86.exe" | Select-String "msvcr|vcruntime|ucrtbase" | ForEach-Object {
                Write-Host "WARNING: CRT dependency found: $_" -ForegroundColor Yellow
                $BuildSuccess = $false
            }
            if ($BuildSuccess) {
                Write-Host "[PASS] No CRT imports detected" -ForegroundColor Green
                # Set file timestamps to 2030-01-01 00:00:00
                $targetFile = "$BinDir\cmdt_x86.exe"
                $futureDate = Get-Date "2030-01-01 00:00:00"
                (Get-Item $targetFile).CreationTime = $futureDate
                (Get-Item $targetFile).LastWriteTime = $futureDate
                Write-Host "Timestamp set to 2030-01-01 00:00:00" -ForegroundColor Cyan
            }
        }
    }
}
Pop-Location

Write-Host ""
Write-Host ">>> Architecture: x64" -ForegroundColor Cyan
Push-Location $ScriptDir
& rc /c65001 /fo cmdt_x64.res cmdt.rc
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Resource compilation failed" -ForegroundColor Red
    $BuildSuccess = $false
} else {
    $x64success = $true
    foreach ($f in $FILES) {
        & $ML64 /c /Cp /Cx /Zi /I x64 /Fo"x64\$f.obj" "x64\$f.asm"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Assembly of $f.asm failed" -ForegroundColor Red
            $x64success = $false
            $BuildSuccess = $false
            break
        }
    }
    if ($x64success) {
        $linkArgs = @("x64\main.obj", "x64\token.obj", "x64\process.obj", "x64\window.obj", "cmdt_x64.res", "/subsystem:windows", "/entry:mainCRTStartup", "/out:bin\cmdt_x64.exe", "/MANIFEST:EMBED", "/MANIFESTINPUT:cmdt.manifest", "/LIBPATH:$LIBPATH64_UM", "/LIBPATH:$LIBPATH64_UCRT") + $LIBS
        & $LINK64 $linkArgs
        if ($LASTEXITCODE -ne 0) { 
            $BuildSuccess = $false 
        } else { 
            Write-Host "Build successful: bin\cmdt_x64.exe" -ForegroundColor Green
            Write-Host "Checking imports..." -ForegroundColor Cyan
            $DUMPBIN64 = "$VSBASE\x64\dumpbin.exe"
            & $DUMPBIN64 /imports "$BinDir\cmdt_x64.exe" | Select-String "msvcr|vcruntime|ucrtbase" | ForEach-Object {
                Write-Host "WARNING: CRT dependency found: $_" -ForegroundColor Yellow
                $BuildSuccess = $false
            }
            if ($BuildSuccess) {
                Write-Host "[PASS] No CRT imports detected" -ForegroundColor Green
                # Set file timestamps to 2030-01-01 00:00:00
                $targetFile = "$BinDir\cmdt_x64.exe"
                $futureDate = Get-Date "2030-01-01 00:00:00"
                (Get-Item $targetFile).CreationTime = $futureDate
                (Get-Item $targetFile).LastWriteTime = $futureDate
                Write-Host "Timestamp set to 2030-01-01 00:00:00" -ForegroundColor Cyan
            }
        }
    }
}
Pop-Location

Write-Host ""
Write-Host "Cleaning up intermediate files..." -ForegroundColor Yellow
Remove-Item "$ScriptDir\x86\*.obj" -ErrorAction SilentlyContinue
Remove-Item "$ScriptDir\x64\*.obj" -ErrorAction SilentlyContinue
Remove-Item "$ScriptDir\*.obj" -ErrorAction SilentlyContinue
Remove-Item "$ScriptDir\*.res" -ErrorAction SilentlyContinue

Write-Host ""
if ($BuildSuccess) {
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "STATUS: SUCCESS" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    exit 0
} else {
    Write-Host "============================================" -ForegroundColor Red
    Write-Host "STATUS: FAILED" -ForegroundColor Red
    Write-Host "============================================" -ForegroundColor Red
    exit 1
}
