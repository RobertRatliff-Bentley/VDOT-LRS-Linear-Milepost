<#
VDOT LRS Linear Milepost - installer

Copies VDOT_LRS_Linear_Milepost.py to a per-user folder and adds it to the Python
Manager list (Utilities > Python) of MicroStation 2026 and OpenRoads Designer 2025.
No administrator rights are needed. Run Install.cmd, or:

    powershell -ExecutionPolicy Bypass -File Install.ps1 [-Uninstall]

-InstallDir and -PrefsRoot exist for testing against a copy of the preferences.
#>
[CmdletBinding()]
param(
    [switch]$Uninstall,
    [switch]$BundledQt,        # install the bundled PyQt5 even when the built-in one works
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'VDOT_LRS_Linear_Milepost'),
    [string]$PrefsRoot = (Join-Path $env:LOCALAPPDATA 'Bentley')
)

$ErrorActionPreference = 'Stop'
$ScriptName = 'VDOT_LRS_Linear_Milepost.py'
$MacroName = 'VDOT_LRS_Linear_Milepost.mvba'
$LibDir = Join-Path $InstallDir 'lib'
$QtRequirement = 'PyQt5==5.15.10'
$Variable = 'MS_USERADDEDPYTHONSCRIPTS'
$PythonRoot = 'C:\ProgramData\Bentley\PowerPlatformPython\python'

# Products the tool has been checked against. Folder/Version locate the per-user
# preferences; Process is used to see whether the product is running.
$Products = @(
    @{ Name = 'MicroStation 2026';       Folder = 'MicroStation';      Version = '26.*'; Process = 'microstation' },
    @{ Name = 'OpenRoads Designer 2025'; Folder = 'OpenRoadsDesigner'; Version = '25.*'; Process = 'OpenRoadsDesigner' }
)

$Target = Join-Path $InstallDir $ScriptName
$TargetEntry = $Target.Replace('\', '/')          # the products store paths with forward slashes
$EntryPattern = '^\s*' + [regex]::Escape($Variable) + '\s*([=>])\s*(.*?)\s*$'
$Utf8Bom = New-Object System.Text.UTF8Encoding($true)

# Earlier builds installed the tool as vdot_lrs_linear_event.py; that copy and its entry are removed.
$LegacyTarget = Join-Path (Join-Path $env:LOCALAPPDATA 'VDOT_LRS_EventTool') 'vdot_lrs_linear_event.py'
$LegacyEntry = $LegacyTarget.Replace('\', '/')

function Test-OurEntry([string]$Line, [string]$Entry = $TargetEntry) {
    $match = [regex]::Match($Line, $EntryPattern)
    return $match.Success -and ($match.Groups[2].Value.Replace('\', '/') -ieq $Entry)
}

function Get-PrefsFiles($Product) {
    $root = Join-Path $PrefsRoot $Product.Folder
    if (-not (Test-Path $root)) { return @() }
    return @(Get-ChildItem $root -Directory -Filter $Product.Version |
        ForEach-Object { Join-Path $_.FullName 'prefs\Personal.ucf' } |
        Where-Object { Test-Path $_ })
}

function Backup-Prefs([string]$Path) {
    $backup = "$Path.before-vdot-lrs"
    if (-not (Test-Path $backup)) { Copy-Item $Path $backup }
}

function Add-Entry([string]$Path) {
    $lines = [System.Collections.Generic.List[string]]([System.IO.File]::ReadAllLines($Path))
    foreach ($line in $lines) { if (Test-OurEntry $line) { return 'already listed' } }
    $last = -1
    for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match $EntryPattern) { $last = $i } }
    Backup-Prefs $Path
    if ($last -ge 0) {
        $lines.Insert($last + 1, "$Variable > $TargetEntry")      # append to the existing list
    } else {
        $lines.Add("$Variable = $TargetEntry")
    }
    [System.IO.File]::WriteAllLines($Path, $lines, $Utf8Bom)
    return 'added'
}

function Remove-Entry([string]$Path, [string]$Entry = $TargetEntry) {
    $lines = [System.IO.File]::ReadAllLines($Path)
    $kept = [System.Collections.Generic.List[string]]@($lines | Where-Object { -not (Test-OurEntry $_ $Entry) })
    if ($kept.Count -eq $lines.Count) { return 'not listed' }
    # The list must still start with "=": promote the first remaining entry if needed.
    $first = -1
    $hasAssignment = $false
    for ($i = 0; $i -lt $kept.Count; $i++) {
        $match = [regex]::Match($kept[$i], $EntryPattern)
        if (-not $match.Success) { continue }
        if ($first -lt 0) { $first = $i }
        if ($match.Groups[1].Value -eq '=') { $hasAssignment = $true }
    }
    if ($first -ge 0 -and -not $hasAssignment) {
        $kept[$first] = [regex]::Replace($kept[$first], '^(\s*' + [regex]::Escape($Variable) + '\s*)>', '$1=')
    }
    Backup-Prefs $Path
    [System.IO.File]::WriteAllLines($Path, $kept, $Utf8Bom)
    return 'removed'
}

Write-Host ''
Write-Host 'VDOT LRS Linear Milepost' -ForegroundColor Cyan
Write-Host ('  ' + $(if ($Uninstall) { 'Uninstalling' } else { 'Installing' }) + " for user $env:USERNAME")
Write-Host ''
$problems = 0

if (-not $Uninstall) {
    $source = Join-Path $PSScriptRoot $ScriptName
    if (-not (Test-Path $source)) {
        Write-Host "ERROR: $ScriptName was not found beside the installer." -ForegroundColor Red
        Write-Host '       Extract the whole zip to a folder first, then run Install.cmd.'
        exit 1
    }
    New-Item $InstallDir -ItemType Directory -Force | Out-Null
    Copy-Item $source $Target -Force
    Write-Host "Copied the tool to:  $Target"
    $mvba = Join-Path $PSScriptRoot $MacroName
    if (Test-Path $mvba) {
        Copy-Item $mvba (Join-Path $InstallDir $MacroName) -Force
        Write-Host "Copied the VBA version to:  $(Join-Path $InstallDir $MacroName)"
    }

    $python = Join-Path $PythonRoot 'python.exe'
    if (-not (Test-Path $python)) {
        $problems++
        Write-Host ''
        Write-Host 'WARNING: Bentley PowerPlatform Python was not found at' -ForegroundColor Yellow
        Write-Host "         $PythonRoot" -ForegroundColor Yellow
        Write-Host '         It is installed with MicroStation 2026 / OpenRoads Designer 2025.'
    } else {
        $ErrorActionPreference = 'Continue'     # python's stderr must not stop the script
        $check = 'import sys; sys.path.extend(sys.argv[1:]); import PyQt5.QtWidgets'
        & $python -c $check 2>$null
        $haveQt = ($LASTEXITCODE -eq 0)
        if ($haveQt -and -not $BundledQt) {
            Write-Host 'PyQt5:  using the copy that ships with MicroStation / OpenRoads.'
        } else {
            $wheels = Join-Path $PSScriptRoot 'wheels'
            if (-not (Test-Path (Join-Path $wheels 'PyQt5-*.whl'))) {
                $problems++
                Write-Host 'WARNING: PyQt5 is missing and the wheels folder is not beside the installer.' -ForegroundColor Yellow
            } else {
                Write-Host "PyQt5:  installing the bundled copy into $LibDir ..."
                foreach ($dir in $LibDir, (Join-Path $InstallDir '__pycache__')) {
        if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
    }
                & $python -m pip install --no-index --find-links $wheels --target $LibDir `
                    --disable-pip-version-check --no-warn-script-location -q $QtRequirement 2>&1 |
                    Where-Object { "$_" -notmatch 'Ignoring invalid distribution' } | ForEach-Object { Write-Host "    $_" }
                & $python -c $check $LibDir 2>$null
                if ($LASTEXITCODE -eq 0) {
                    Write-Host '        done - the tool uses it only when the built-in PyQt5 is missing.' -ForegroundColor Green
                } else {
                    $problems++
                    Write-Host 'ERROR: the bundled PyQt5 could not be installed or loaded.' -ForegroundColor Red
                }
            }
        }
        $ErrorActionPreference = 'Stop'
    }
}

Write-Host ''
$registered = 0
foreach ($product in $Products) {
    $files = Get-PrefsFiles $product
    if ($files.Count -eq 0) {
        Write-Host ("  {0,-26} not found for this user (start the product once, then run this again)" -f $product.Name)
        continue
    }
    if (Get-Process -Name $product.Process -ErrorAction SilentlyContinue) {
        $problems++
        Write-Host ("  {0,-26} is RUNNING - skipped. Close it and run this again." -f $product.Name) -ForegroundColor Yellow
        continue
    }
    foreach ($file in $files) {
        try {
            $legacy = Remove-Entry $file $LegacyEntry
            if ($legacy -eq 'removed') {
                Write-Host ("  {0,-26} removed the old vdot_lrs_linear_event entry" -f $product.Name) -ForegroundColor Green
            }
            $result = if ($Uninstall) { Remove-Entry $file } else { Add-Entry $file }
            Write-Host ("  {0,-26} Python Manager list: {1}" -f $product.Name, $result) -ForegroundColor Green
            if (-not $Uninstall) { $registered++ }
        } catch {
            $problems++
            Write-Host ("  {0,-26} could not update {1}: {2}" -f $product.Name, $file, $_.Exception.Message) -ForegroundColor Red
        }
    }
}

if (Test-Path $LegacyTarget) {
    Remove-Item $LegacyTarget -Force
    $legacyDir = Split-Path $LegacyTarget
    if (-not (Get-ChildItem $legacyDir -Force)) { Remove-Item $legacyDir -Force }
    Write-Host "Removed the old copy:  $LegacyTarget"
}

Write-Host ''
if ($Uninstall) {
    if (Test-Path $Target) { Remove-Item $Target -Force }
    $macro = Join-Path $InstallDir $MacroName
    if (Test-Path $macro) { Remove-Item $macro -Force }
    foreach ($dir in $LibDir, (Join-Path $InstallDir '__pycache__')) {
        if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
    }
    if ((Test-Path $InstallDir) -and -not (Get-ChildItem $InstallDir -Force)) { Remove-Item $InstallDir -Force }
    Write-Host "Removed the tool from:  $InstallDir"
} else {
    Write-Host 'To run the tool:'
    if ($registered -gt 0) {
        Write-Host '  Utilities > Python (Python Manager), choose VDOT_LRS_Linear_Milepost, Run'
    }
    Write-Host '  or key in:  PYTHON LOAD "' -NoNewline
    Write-Host $TargetEntry -NoNewline
    Write-Host '"'
    if (Test-Path (Join-Path $InstallDir $MacroName)) {
        Write-Host '  VBA version:  VBA LOAD "' -NoNewline
        Write-Host ((Join-Path $InstallDir $MacroName).Replace('\', '/')) -NoNewline
        Write-Host '"'
        Write-Host '                then  VBA RUN [VDOT_LRS_Linear_Milepost]modLrsMain.VdotLrsEvent'
    }
}
Write-Host ''
if ($problems -gt 0) {
    Write-Host "Finished with $problems item(s) needing attention (see above)." -ForegroundColor Yellow
    exit 2
}
Write-Host 'Finished.' -ForegroundColor Green
exit 0
