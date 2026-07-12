$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BinDir = Join-Path $ScriptDir "bin"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Building CMDT - Run as TrustedInstaller" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

function Get-LatestVCToolsPath {
    $vswhere = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        throw "Nie znaleziono vswhere.exe pod $vswhere. Zainstaluj/napraw Visual Studio Installer."
    }
    $vsInstallPath = & $vswhere -latest -products * `
        -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
        -property installationPath
    if (-not $vsInstallPath) {
        throw "vswhere nie znalazl instalacji VS z komponentem VC.Tools.x86.x64"
    }
    $verFile = Join-Path $vsInstallPath "VC\Auxiliary\Build\Microsoft.VCToolsVersion.default.txt"
    $vcVersion = (Get-Content $verFile -Raw).Trim()
    return Join-Path $vsInstallPath "VC\Tools\MSVC\$vcVersion\bin\Hostx64"
}

function Get-LatestWinSDK {
    $sdkRoot = "C:\Program Files (x86)\Windows Kits\10"
    $latest = Get-ChildItem "$sdkRoot\Include" -Directory |
        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
        Sort-Object { [version]$_.Name } -Descending |
        Select-Object -First 1
    if (-not $latest) {
        throw "Nie znaleziono zadnej wersji Windows SDK w $sdkRoot\Include"
    }
    return $latest.Name
}

$VSBASE  = Get-LatestVCToolsPath
$ML32    = "$VSBASE\x86\ml.exe"
$ML64    = "$VSBASE\x64\ml64.exe"
$LINK32  = "$VSBASE\x86\link.exe"
$LINK64  = "$VSBASE\x64\link.exe"

$SDKVER        = Get-LatestWinSDK
Write-Host ">>> Using VC Tools: $VSBASE" -ForegroundColor DarkGray
Write-Host ">>> Using Windows SDK: $SDKVER" -ForegroundColor DarkGray

$SDKBASE       = "C:\Program Files (x86)\Windows Kits\10\Lib\$SDKVER"
$SDKBIN        = "C:\Program Files (x86)\Windows Kits\10\bin\$SDKVER\x64"
$SDKBINX86     = "C:\Program Files (x86)\Windows Kits\10\bin\$SDKVER\x86"
$SDKINCLUDE    = "C:\Program Files (x86)\Windows Kits\10\Include\$SDKVER"
$LIBPATH32_UM  = "$SDKBASE\um\x86"
$LIBPATH32_UCRT = "$SDKBASE\ucrt\x86"
$LIBPATH64_UM  = "$SDKBASE\um\x64"
$LIBPATH64_UCRT = "$SDKBASE\ucrt\x64"

# rc.exe czasem siedzi tylko w folderze x86, niezaleznie od architektury builda
$RC = "$SDKBIN\rc.exe"
if (-not (Test-Path $RC)) { $RC = "$SDKBINX86\rc.exe" }
if (-not (Test-Path $RC)) { throw "Nie znaleziono rc.exe ani w $SDKBIN ani w $SDKBINX86" }

$env:PATH += ";$SDKBIN"
$env:INCLUDE = "$SDKINCLUDE\um;$SDKINCLUDE\shared;$SDKINCLUDE\ucrt"

if (-not (Test-Path $BinDir)) {
    New-Item -ItemType Directory -Path $BinDir | Out-Null
}

$FILES_X86 = @("main", "token", "process", "window", "strutil", "help", "install", "relay", "cli")
$FILES_X64 = @("main", "token", "process", "window", "strutil", "help", "install", "relay", "cli")
$LIBS = @("kernel32.lib", "user32.lib", "advapi32.lib", "shlwapi.lib", "shell32.lib", "gdi32.lib", "comdlg32.lib", "userenv.lib", "ole32.lib", "dwmapi.lib", "OleAut32.lib")
$BuildSuccess = $true

Write-Host ""
Write-Host ">>> Architecture: x86" -ForegroundColor Cyan
Push-Location $ScriptDir
& $RC /c65001 /fo cmdt_x86.res cmdt.rc
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Resource compilation failed" -ForegroundColor Red
    $BuildSuccess = $false
} else {
    foreach ($f in $FILES_X86) {
        & $ML32 /c /Cp /Cx /Zi /I x86 /Fo"x86\$f.obj" "x86\$f.asm"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Assembly of $f.asm failed" -ForegroundColor Red
            $BuildSuccess = $false
            break
        }
    }
    if ($BuildSuccess) {
        $linkArgs = @("x86\main.obj", "x86\token.obj", "x86\process.obj", "x86\window.obj", "x86\strutil.obj", "x86\help.obj", "x86\install.obj", "x86\relay.obj", "x86\cli.obj", "cmdt_x86.res", "/subsystem:windows", "/entry:start@0", "/Brepro", "/out:bin\cmdt_x86.exe", "/MANIFEST:EMBED", "/MANIFESTINPUT:cmdt.manifest", "/LIBPATH:$LIBPATH32_UM", "/LIBPATH:$LIBPATH32_UCRT") + $LIBS
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
& $RC /c65001 /fo cmdt_x64.res cmdt.rc
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Resource compilation failed" -ForegroundColor Red
    $BuildSuccess = $false
} else {
    $x64success = $true
    foreach ($f in $FILES_X64) {
        & $ML64 /c /Cp /Cx /Zi /I x64 /Fo"x64\$f.obj" "x64\$f.asm"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Assembly of $f.asm failed" -ForegroundColor Red
            $x64success = $false
            $BuildSuccess = $false
            break
        }
    }
    if ($x64success) {
        $linkArgs = @("x64\main.obj", "x64\token.obj", "x64\process.obj", "x64\window.obj", "x64\strutil.obj", "x64\help.obj", "x64\install.obj", "x64\relay.obj", "x64\cli.obj", "cmdt_x64.res", "/subsystem:windows", "/entry:mainCRTStartup", "/Brepro", "/out:bin\cmdt_x64.exe", "/MANIFEST:EMBED", "/MANIFESTINPUT:cmdt.manifest", "/LIBPATH:$LIBPATH64_UM", "/LIBPATH:$LIBPATH64_UCRT") + $LIBS
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