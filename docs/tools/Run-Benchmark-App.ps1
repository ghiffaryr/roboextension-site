#Requires -Version 5.1
<#
.SYNOPSIS
    Black-box benchmark: times the actual RoboExtension.exe binary against competitors.

.DESCRIPTION
    Drives RoboExtension.exe via standard command-line verbs so the copy and delete
    engines run as a real process showing the actual progress UI that end users see.
    Competitors (Explorer, TeraCopy, FastCopy) are driven the same way as always.

    Use this script when publishing benchmark results: timed runs show the real UI,
    proving results without exposing implementation details.

.PARAMETER RoboExe
    Full path to RoboExtension.exe. Auto-detected from the installed location only. Install the app before running.

.PARAMETER SsdPath
    A writable folder on your SSD (e.g. C:\BenchTemp). Created if absent.

.PARAMETER HddPath
    A writable folder on your HDD (e.g. D:\BenchTemp). Only SSD tests run if omitted.

.PARAMETER TeraCopyExe
    Full path to TeraCopy.exe. Auto-detected from common install paths if omitted.

.PARAMETER FastCopyExe
    Full path to FastCopy.exe. Auto-detected from common install paths if omitted.

.PARAMETER Runs
    Number of timed repetitions per cell. Median is reported. Default = 3.

.PARAMETER OutputJson
    Where to write the JSON results file. Defaults to .\results-app.json.

.PARAMETER ForceRegen
    Delete and recreate existing source datasets from scratch.

.PARAMETER ScenarioFilter
    Only run scenarios whose name contains this substring (e.g. 'large').

.PARAMETER ComboFilter
    Only run storage combos whose label matches any comma-separated substring
    (e.g. 'SSD->SSD' or 'HDD').

.EXAMPLE
    # Easiest: double-click Run-Benchmark-App.bat (handles elevation + execution policy automatically)
    .\Run-Benchmark-App.bat C:\BenchTemp D:\BenchTemp

.EXAMPLE
    # Advanced / custom paths:
    powershell -ExecutionPolicy Bypass -File .\Run-Benchmark-App.ps1 -SsdPath C:\BenchTemp -HddPath D:\BenchTemp
#>
param(
    [string]$RoboExe      = "",

    [Parameter(Mandatory)]
    [string]$SsdPath,

    [string]$HddPath      = "",
    [string]$TeraCopyExe  = "",
    [string]$FastCopyExe  = "",
    [int]   $Runs         = 3,
    [string]$OutputJson   = "",
    [switch]$ForceRegen,

    # Optional: path to a folder of real small files (1 KB-100 KB). If provided,
    # ~8,000 files are sampled from it instead of generating synthetic data.
    [string]$SmallFilesSource = '',

    # Optional: path to a folder of real large files (>= 100 MB). If provided,
    # up to 5 files are copied instead of generating synthetic 512 MB files.
    [string]$LargeFilesSource = '',

    [string]$ScenarioFilter = '',
    [string]$ComboFilter    = '',

    # If set, only run tools whose name contains any of these comma-separated substrings
    # (e.g. 'Explorer,RoboExtension' or just 'RoboExtension').
    [string]$ToolFilter     = '',

    # Controls which sections are executed:
    #   full   (default) - run copy benchmark then delete benchmark
    #   copy             - run copy benchmark only
    #   delete           - run delete benchmark only
    [ValidateSet('full','copy','delete')]
    [string]$Mode = 'full',

    # When set, loads existing results from OutputJson and preserves all rows
    # that the current run won't overwrite:
    #   -Mode delete -Resume  ->  copy rows preserved, delete rows re-run
    #   -Mode copy   -Resume  ->  delete rows preserved, copy rows re-run
    #   -Mode full   -Resume  ->  nothing preserved (full fresh run)
    [switch]$Resume
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $OutputJson) { $OutputJson = Join-Path $PSScriptRoot 'results-app.json' }

# ---------------------------------------------------------------------------
# Elevation check - required for page-cache flush
# ---------------------------------------------------------------------------
$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isElevated) {
    Write-Host '' 
    Write-Host '  ERROR  Administrator privileges required.' -ForegroundColor Red
    Write-Host '         Cache flush needs SeProfileSingleProcessPrivilege.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Re-run in an elevated PowerShell window.' -ForegroundColor Cyan
    Write-Host ''
    exit 1
}

# ---------------------------------------------------------------------------
# Locate RoboExtension.exe
# ---------------------------------------------------------------------------
function Resolve-Exe([string]$hint, [string[]]$defaults) {
    if ($hint -and (Test-Path $hint)) { return $hint }
    foreach ($d in $defaults) {
        $e = [Environment]::ExpandEnvironmentVariables($d)
        if (Test-Path $e) { return $e }
    }
    return $null
}

$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$roboExePath = Resolve-Exe $RoboExe @(
    "$env:ProgramFiles\RoboExtension\RoboExtension.exe"
)
if (-not $roboExePath) {
    Write-Host '' 
    Write-Host '  ERROR  RoboExtension.exe not found.' -ForegroundColor Red
    Write-Host '         Install the app first (run dist\RoboExtension-Setup.exe), or pass -RoboExe <path>.' -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

$teraExe = Resolve-Exe $TeraCopyExe @(
    '%ProgramFiles%\TeraCopy\TeraCopy.exe',
    '%ProgramFiles(x86)%\TeraCopy\TeraCopy.exe'
)
$fastExe = Resolve-Exe $FastCopyExe @(
    '%USERPROFILE%\FastCopy\FastCopy.exe',
    '%ProgramFiles%\FastCopy\FastCopy.exe',
    '%ProgramFiles(x86)%\FastCopy\FastCopy.exe',
    '%LocalAppData%\Programs\FastCopy\FastCopy.exe',
    '%USERPROFILE%\FastCopy\fcp.exe'
)

Write-Host "`n=== TOOL DETECTION ===" -ForegroundColor Cyan
Write-Host "  RoboExtension   : $roboExePath" -ForegroundColor Green
Write-Host "  robocopy        : built-in" -ForegroundColor Gray
if ($teraExe) { Write-Host "  TeraCopy        : $teraExe" -ForegroundColor Green }
else          { Write-Warning "TeraCopy not found - its tests will be skipped." }
if ($fastExe)  { Write-Host "  FastCopy        : $fastExe"  -ForegroundColor Green }
else           { Write-Warning "FastCopy not found - its tests will be skipped." }

# ---------------------------------------------------------------------------
# Windows Shell helper (Explorer benchmark)
# ---------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'WinShell').Type) {
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WinShell {
    const uint   FO_COPY            = 2;
    const ushort FOF_SILENT         = 0x0004;
    const ushort FOF_NOCONFIRMATION = 0x0010;
    const ushort FOF_NOERRORUI      = 0x0400;
    const ushort FOF_NOCONFIRMMKDIR = 0x0200;
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct SHFILEOPSTRUCT {
        public IntPtr hwnd; public uint wFunc; public string pFrom;
        public string pTo; public ushort fFlags;
        [MarshalAs(UnmanagedType.Bool)] public bool fAnyOperationsAborted;
        public IntPtr hNameMappings; public string lpszProgressTitle;
    }
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);
    public static void CopyContents(string src, string dest) {
        var op = new SHFILEOPSTRUCT {
            wFunc  = FO_COPY,
            pFrom  = src.TrimEnd('\\') + "\\*\0",
            pTo    = dest.TrimEnd('\\') + "\0",
            fFlags = FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_NOCONFIRMMKDIR
        };
        int r = SHFileOperation(ref op);
        if (r != 0) throw new InvalidOperationException("SHFileOperation error: 0x" + r.ToString("X"));
    }
}
'@
}

# ---------------------------------------------------------------------------
# Page-cache flush
# ---------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'CacheUtil').Type) {
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CacheUtil {
    [DllImport("ntdll.dll")] static extern uint NtSetSystemInformation(int cls, ref int info, int len);
    [DllImport("advapi32.dll", SetLastError=true)] static extern bool OpenProcessToken(IntPtr h, uint access, out IntPtr token);
    [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)] static extern bool LookupPrivilegeValue(string sys, string name, out long luid);
    [DllImport("advapi32.dll", SetLastError=true)] static extern bool AdjustTokenPrivileges(IntPtr token, bool disable, ref TOKEN_PRIVILEGES tp, int len, IntPtr prev, IntPtr retlen);
    [DllImport("kernel32.dll")] static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool SetProcessWorkingSetSize(IntPtr hProcess, IntPtr dwMin, IntPtr dwMax);
    [StructLayout(LayoutKind.Sequential, Pack=4)] struct TOKEN_PRIVILEGES { public int Count; public long Luid; public int Attr; }
    const uint TOKEN_ADJUST_PRIVILEGES = 0x0020;
    const uint TOKEN_QUERY             = 0x0008;
    const int  SE_PRIVILEGE_ENABLED    = 0x0002;
    static void EnablePrivilege(IntPtr token, string name) {
        long luid;
        if (!LookupPrivilegeValue(null, name, out luid)) return;
        var tp = new TOKEN_PRIVILEGES { Count = 1, Luid = luid, Attr = SE_PRIVILEGE_ENABLED };
        AdjustTokenPrivileges(token, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    }
    public static void FlushAll() {
        SetProcessWorkingSetSize(GetCurrentProcess(), new IntPtr(-1), new IntPtr(-1));
        IntPtr token;
        if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, out token))
            throw new InvalidOperationException("OpenProcessToken failed: " + Marshal.GetLastWin32Error());
        try {
            EnablePrivilege(token, "SeProfileSingleProcessPrivilege");
            EnablePrivilege(token, "SeIncreaseQuotaPrivilege");
        } finally { CloseHandle(token); }
        uint r; int v;
        v = 3; r = NtSetSystemInformation(80, ref v, 4);
        if ((r & 0x80000000u) != 0) throw new InvalidOperationException("NtSetSystemInformation(80,3) failed: 0x" + r.ToString("X8"));
        v = 4; r = NtSetSystemInformation(80, ref v, 4);
        if ((r & 0x80000000u) != 0) throw new InvalidOperationException("NtSetSystemInformation(80,4) failed: 0x" + r.ToString("X8"));
    }
}
'@
}
function Clear-PageCache {
    try { [CacheUtil]::FlushAll() }
    catch { Write-Warning "Page-cache flush failed: $_" }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Median([double[]]$arr) {
    $s = @($arr | Sort-Object); $n = $s.Count
    if ($n -eq 0) { return [double]::NaN }
    if ($n % 2 -eq 1) { return $s[($n-1)/2] }
    return ($s[$n/2 - 1] + $s[$n/2]) / 2.0
}
function Get-FolderSize([string]$path) {
    [long](Get-ChildItem -LiteralPath $path -Recurse -File | Measure-Object -Property Length -Sum).Sum
}
function Format-Secs([double]$s) {
    if ([double]::IsNaN($s) -or $s -le 0) { return 'N/A' }
    if ($s -lt 60) { return ('{0:F1} s' -f $s) }
    '{0}m {1}s' -f [int]($s/60), [int]($s % 60)
}
function Format-MBps([long]$bytes, [double]$s) {
    if ([double]::IsNaN($s) -or $s -le 0) { return 'N/A' }
    '{0:F1} MB/s' -f ($bytes / $s / 1MB)
}

# ---------------------------------------------------------------------------
# Dataset generation (same synthetic seed-42 data as Run-Benchmark.ps1)
# ---------------------------------------------------------------------------
function Write-TestFile([string]$path, [long]$sizeBytes) {
    $bufLen = [Math]::Min([long]256KB, [long]$sizeBytes)
    $buf    = [byte[]]::new($bufLen)
    for ($i = 0; $i -lt $bufLen; $i++) { $buf[$i] = [byte]($i % 251) }
    $fs = [System.IO.FileStream]::new($path,
        [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write,
        [System.IO.FileShare]::None, $bufLen)
    try {
        $rem = $sizeBytes
        while ($rem -gt 0) { $n = [Math]::Min($bufLen, $rem); $fs.Write($buf, 0, $n); $rem -= $n }
    } finally { $fs.Dispose() }
}

function Initialize-Datasets([string]$workspace) {
    Write-Host "`n[DATA] Initializing datasets in $workspace" -ForegroundColor Cyan

    # ── small: 8,000 files (4-64 KB synthetic, or sampled from SmallFilesSource) ──
    $dir = Join-Path $workspace "src\small"
    if ($ForceRegen -and (Test-Path $dir)) { Remove-Item $dir -Recurse -Force }
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        if ($SmallFilesSource -and (Test-Path $SmallFilesSource)) {
            Write-Host "  Sampling ~8,000 small files from $SmallFilesSource ..." -NoNewline -ForegroundColor Gray
            $pool   = @(Get-ChildItem -LiteralPath $SmallFilesSource -Recurse -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Length -ge 1KB -and $_.Length -lt 100KB })
            $sample = $pool | Get-Random -Count ([Math]::Min(8000, $pool.Count))
            $i = 1
            foreach ($f in $sample) {
                Copy-Item -LiteralPath $f.FullName `
                    -Destination (Join-Path $dir ("f{0}{1}" -f $i, $f.Extension)) -Force
                $i++
            }
            Write-Host " $($i-1) files" -ForegroundColor Green
        } else {
            Write-Host "  Generating small (8,000 files, 4-64 KB each)..." -NoNewline -ForegroundColor Gray
            $rng = [System.Random]::new(42)
            for ($i = 1; $i -le 8000; $i++) {
                Write-TestFile (Join-Path $dir "f$i.dat") (4KB + $rng.Next(60KB))
            }
            Write-Host " done" -ForegroundColor Green
        }
    } else { Write-Host "  small  : already exists - use -ForceRegen to rebuild" -ForegroundColor DarkGray }

    # ── large: 5 large files (512 MB synthetic, or sampled from LargeFilesSource) ──
    $dir = Join-Path $workspace "src\large"
    if ($ForceRegen -and (Test-Path $dir)) { Remove-Item $dir -Recurse -Force }
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        if ($LargeFilesSource -and (Test-Path $LargeFilesSource)) {
            Write-Host "  Sampling up to 5 large files from $LargeFilesSource ..." -NoNewline -ForegroundColor Gray
            $pool   = @(Get-ChildItem -LiteralPath $LargeFilesSource -Recurse -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Length -ge 100MB })
            $sample = $pool | Get-Random -Count ([Math]::Min(5, $pool.Count))
            $i = 1
            foreach ($f in $sample) {
                Copy-Item -LiteralPath $f.FullName `
                    -Destination (Join-Path $dir ("f{0}{1}" -f $i, $f.Extension)) -Force
                $i++
            }
            Write-Host " $($i-1) files" -ForegroundColor Green
        } else {
            Write-Host "  Generating large (5 x 512 MB)..." -ForegroundColor Gray
            for ($i = 1; $i -le 5; $i++) {
                Write-Host ("    [{0}/5] writing 512 MB..." -f $i) -NoNewline -ForegroundColor Gray
                Write-TestFile (Join-Path $dir "f$i.dat") 512MB
                Write-Host " done" -ForegroundColor Green
            }
        }
    } else { Write-Host "  large  : already exists - use -ForceRegen to rebuild" -ForegroundColor DarkGray }

    # ── mixed: 3,000 files auto-sampled from small+large sources, or synthetic ──
    $dir = Join-Path $workspace "src\mixed"
    if ($ForceRegen -and (Test-Path $dir)) { Remove-Item $dir -Recurse -Force }
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $hasSmallSrc = $SmallFilesSource -and (Test-Path $SmallFilesSource)
        $hasLargeSrc = $LargeFilesSource -and (Test-Path $LargeFilesSource)
        if ($hasSmallSrc -or $hasLargeSrc) {
            Write-Host "  Sampling ~3,000 mixed files from source folders..." -NoNewline -ForegroundColor Gray
            $pool = @()
            if ($hasSmallSrc) {
                $pool += @(Get-ChildItem -LiteralPath $SmallFilesSource -Recurse -File -ErrorAction SilentlyContinue |
                           Where-Object { $_.Length -ge 1KB })
            }
            if ($hasLargeSrc) {
                $pool += @(Get-ChildItem -LiteralPath $LargeFilesSource -Recurse -File -ErrorAction SilentlyContinue |
                           Where-Object { $_.Length -ge 1MB })
            }
            $sample = $pool | Get-Random -Count ([Math]::Min(3000, $pool.Count))
            $i = 1
            foreach ($f in $sample) {
                Copy-Item -LiteralPath $f.FullName `
                    -Destination (Join-Path $dir ("f{0}{1}" -f $i, $f.Extension)) -Force
                $i++
            }
            Write-Host " $($i-1) files" -ForegroundColor Green
        } else {
            Write-Host "  Generating mixed (3,000 files, 4 KB-10 MB)..." -NoNewline -ForegroundColor Gray
            $rng   = [System.Random]::new(42)
            $sizes = @(4KB, 8KB, 16KB, 32KB, 64KB, 128KB, 256KB, 512KB, 1MB, 2MB, 5MB, 10MB)
            for ($i = 1; $i -le 3000; $i++) {
                Write-TestFile (Join-Path $dir "f$i.dat") $sizes[$rng.Next($sizes.Length)]
            }
            Write-Host " done" -ForegroundColor Green
        }
    } else { Write-Host "  mixed  : already exists - use -ForceRegen to rebuild" -ForegroundColor DarkGray }
}

# ---------------------------------------------------------------------------
# Individual tool runners
# ---------------------------------------------------------------------------
function Invoke-Explorer([string]$src, [string]$dst) {
    [WinShell]::CopyContents($src, $dst)
}

function Invoke-RoboCopy([string]$src, [string]$dst) {
    $p = Start-Process robocopy `
        -ArgumentList "`"$src`"", "`"$dst`"", '/E', '/R:0', '/W:0', '/NJH', '/NJS', '/NP' `
        -PassThru -Wait -WindowStyle Hidden
    if ($p.ExitCode -ge 8) { throw "robocopy exit=$($p.ExitCode)" }
}

# Black-box: invokes the real installed binary WITH progress UI (dst must not pre-exist so isDirect fires without --silent).
function Invoke-RoboExtension([string]$exe, [string]$src, [string]$dst) {
    # -PassThru + WaitForExit() instead of -Wait: avoids WaitForInputIdle() which adds
    # ~500ms for GUI-subsystem binaries before the engine even starts.
    $p = Start-Process $exe -ArgumentList @('copy', "`"$src`"", "`"$dst`"") -PassThru -WindowStyle Normal
    $p.WaitForExit()
    if ($p.ExitCode -ne 0) { throw "RoboExtension copy exit=$($p.ExitCode)" }
}

function Invoke-TeraCopy([string]$exe, [string]$src, [string]$dst) {
    $p = Start-Process $exe `
        -ArgumentList 'Copy', "`"$src`"", "`"$dst`"", '/Wait', '/Close', '/SkipAll' `
        -PassThru -Wait -WindowStyle Normal
    $copied = @(Get-ChildItem -LiteralPath $dst -Recurse -File -ErrorAction SilentlyContinue).Count
    if ($copied -eq 0) { throw "TeraCopy copied 0 files to $dst" }
}

function Invoke-FastCopy([string]$exe, [string]$src, [string]$dst) {
    $p = Start-Process $exe `
        -ArgumentList '/cmd=force_copy', '/no_ui', '/auto_close', '/bufsize=256M', "`"$src`"", "/to=`"$dst\`"" `
        -PassThru -Wait -WindowStyle Hidden
    if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 1) { throw "FastCopy exit=$($p.ExitCode)" }
}

# cmd /C del: Explorer-equivalent single-threaded delete
function Invoke-CmdDel([string]$folder) {
    $p = Start-Process cmd -ArgumentList '/C', "del /F /S /Q `"$folder`" > nul 2>&1 && rd /S /Q `"$folder`"" `
        -PassThru -Wait -WindowStyle Hidden
    # del+rd always exits 0; verify folder is gone
    if (Test-Path $folder) { throw "cmd del did not remove $folder" }
}

# PowerShell Remove-Item: alternative baseline
function Invoke-PSDelete([string]$folder) {
    Remove-Item -LiteralPath $folder -Recurse -Force
}

# Black-box: invokes the real delete engine WITH progress UI, skipping only the confirm dialog.
function Invoke-RoboExtensionDelete([string]$exe, [string]$folder) {
    # -PassThru + WaitForExit() instead of -Wait: avoids WaitForInputIdle() which adds
    # ~500ms for GUI-subsystem binaries before the engine even starts.
    $p = Start-Process $exe -ArgumentList @('delete', "`"$folder`"", '--f', '--yes') -PassThru -WindowStyle Normal
    $p.WaitForExit()
    if ($p.ExitCode -ne 0) { throw "RoboExtension delete --f exit=$($p.ExitCode)" }
}

# FastCopy
function Invoke-FastCopyDelete([string]$exe, [string]$folder) {
    $p = Start-Process $exe `
        -ArgumentList '/cmd=delete', '/no_ui', '/auto_close', "`"$folder`"" `
        -PassThru -Wait -WindowStyle Hidden
    if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 1) { throw "FastCopy delete exit=$($p.ExitCode)" }
}

# ---------------------------------------------------------------------------
# Timed measurement (copy)
# ---------------------------------------------------------------------------
function Measure-Tool([string]$toolId, [string]$toolExe, [string]$src, [string]$dst) {
    if (Test-Path $dst) { Remove-Item -LiteralPath $dst -Recurse -Force }
    # RoboExtension needs dst absent so the isDirect path fires and ProgressForm is shown.
    # Other tools (Explorer, TeraCopy) need the target directory to already exist.
    if ($toolId -ne 'roboextension') { New-Item -ItemType Directory -Path $dst -Force | Out-Null }
    Clear-PageCache

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        switch ($toolId) {
            'explorer'      { Invoke-Explorer       $src $dst }
            'robocopy'      { Invoke-RoboCopy        $src $dst }
            'roboextension' { Invoke-RoboExtension  $toolExe $src $dst }
            'teracopy'      { Invoke-TeraCopy       $toolExe $src $dst }
            'fastcopy'      { Invoke-FastCopy       $toolExe $src $dst }
        }
    } catch { $sw.Stop(); throw }
    $sw.Stop()
    return $sw.Elapsed.TotalSeconds
}

# Timed measurement (delete): copies src to a temp folder then deletes that copy
function Measure-Delete([string]$toolId, [string]$toolExe, [string]$src, [string]$workspace) {
    # Pre-copy with robocopy (headless) so we measure only the delete, not the copy
    $tmp = Join-Path $workspace "del_tmp_$toolId"
    if (Test-Path $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force }
    Invoke-RoboCopy $src $tmp
    # Flush cache so delete starts cold
    Clear-PageCache

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        switch ($toolId) {
            'windows'       { Invoke-CmdDel              $tmp }
            'roboextension' { Invoke-RoboExtensionDelete  $toolExe $tmp }
            'fastcopy'      { Invoke-FastCopyDelete       $toolExe $tmp }
        }
    } catch { $sw.Stop(); throw }
    $sw.Stop()
    # Ensure tmp is gone regardless (cleanup)
    if (Test-Path $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    return $sw.Elapsed.TotalSeconds
}

# ---------------------------------------------------------------------------
# Build tool list
# ---------------------------------------------------------------------------
$toolList = [System.Collections.Specialized.OrderedDictionary]::new()
$toolList['Windows Explorer']         = [pscustomobject]@{ Id='explorer';      Exe='' }
$toolList['Robocopy (plain /E)']      = [pscustomobject]@{ Id='robocopy';      Exe='' }
$toolList['RoboExtension (adaptive)'] = [pscustomobject]@{ Id='roboextension'; Exe=$roboExePath }
if ($teraExe) { $toolList['TeraCopy'] = [pscustomobject]@{ Id='teracopy'; Exe=$teraExe } }
if ($fastExe)  { $toolList['FastCopy'] = [pscustomobject]@{ Id='fastcopy';  Exe=$fastExe  } }
if ($ToolFilter) {
    $tokens = $ToolFilter -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($key in @($toolList.Keys)) {
        $matched = $false
        foreach ($tok in $tokens) { if ($key -like "*$tok*") { $matched = $true; break } }
        if (-not $matched) { $toolList.Remove($key) }
    }
}

# ---------------------------------------------------------------------------
# Storage combos
# ---------------------------------------------------------------------------
$combos = [System.Collections.Generic.List[pscustomobject]]::new()
$combos.Add([pscustomobject]@{ Label='SSD->SSD'; SrcRoot=$SsdPath; DstRoot=$SsdPath })
if ($HddPath) {
    $combos.Add([pscustomobject]@{ Label='HDD->HDD'; SrcRoot=$HddPath; DstRoot=$HddPath })
    $combos.Add([pscustomobject]@{ Label='SSD->HDD'; SrcRoot=$SsdPath; DstRoot=$HddPath })
    $combos.Add([pscustomobject]@{ Label='HDD->SSD'; SrcRoot=$HddPath; DstRoot=$SsdPath })
}

# ---------------------------------------------------------------------------
# Generate datasets
# ---------------------------------------------------------------------------
$wsSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($c in $combos) { $wsSet.Add($c.SrcRoot) | Out-Null }
foreach ($ws in $wsSet)  { Initialize-Datasets $ws }

$scenarioMeta = @(
    [pscustomobject]@{ Name='small'; Label="Many Small Files (8,000 x 4-64 KB)" }
    [pscustomobject]@{ Name='large'; Label="Few Large Files (5 x 512 MB)" }
    [pscustomobject]@{ Name='mixed'; Label="Mixed Workload (3,000 x 4 KB-10 MB)" }
)
$scenBytes = @{}
foreach ($sc in $scenarioMeta) {
    $scenBytes[$sc.Name] = Get-FolderSize (Join-Path $SsdPath "src\$($sc.Name)")
}

# Delete tools (ordered): cmd del, PowerShell, FastCopy delete, RoboExtension
$deleteToolList = [System.Collections.Specialized.OrderedDictionary]::new()
$deleteToolList['Windows Explorer (Shift+Del)'] = [pscustomobject]@{ Id='windows';       Exe='' }
if ($fastExe) { $deleteToolList['FastCopy'] = [pscustomobject]@{ Id='fastcopy'; Exe=$fastExe } }
$deleteToolList['RoboExtension'] = [pscustomobject]@{ Id='roboextension'; Exe=$roboExePath }
if ($ToolFilter) {
    $tfTokens = $ToolFilter -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $keysToRemove = @($deleteToolList.Keys | Where-Object { $dk = $_; -not ($tfTokens | Where-Object { $dk -like "*$_*" }) })
    foreach ($dk in $keysToRemove) { $deleteToolList.Remove($dk) }
}

# ---------------------------------------------------------------------------
# Main benchmark loop
# ---------------------------------------------------------------------------
$allResults       = [System.Collections.Generic.List[pscustomobject]]::new()

# -Resume: keep all existing rows that the current run won't overwrite.
if ($Resume -and (Test-Path $OutputJson)) {
    Write-Host "`n[INFO] Resume: loading existing results from $OutputJson" -ForegroundColor Cyan
    $existing = (Get-Content -Raw $OutputJson | ConvertFrom-Json).Results
    foreach ($r in $existing) {
        $isDeleteRow = $r.Storage -like '*delete*'
        # Keep delete rows when re-running copy; keep copy rows when re-running delete.
        $keep = ($isDeleteRow -and $Mode -eq 'copy') -or
                (-not $isDeleteRow -and $Mode -eq 'delete')
        if ($keep) { $allResults.Add([pscustomobject]$r) }
    }
    Write-Host "[INFO] Loaded $($allResults.Count) preserved result(s)." -ForegroundColor DarkGray
}

if ($Mode -ne 'delete') { foreach ($combo in $combos) {
    if ($ComboFilter) {
        $tokens  = $ComboFilter -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $matched = $false
        foreach ($tok in $tokens) { if ($combo.Label -like "*$tok*") { $matched = $true; break } }
        if (-not $matched) { continue }
    }

    Write-Host ("`n`n=== COPY BENCHMARK ({0}) ===" -f $combo.Label) -ForegroundColor Cyan

    foreach ($sc in $scenarioMeta) {
        if ($ScenarioFilter -and $sc.Name -notlike "*$ScenarioFilter*") { continue }

        $srcDir = Join-Path $combo.SrcRoot "src\$($sc.Name)"
        $bytes  = $scenBytes[$sc.Name]

        Write-Host ("`n--- {0} ---" -f $sc.Label) -ForegroundColor Yellow

        foreach ($kv in $toolList.GetEnumerator()) {
            $toolName = $kv.Key
            $toolMeta = $kv.Value
            $dstDir   = Join-Path $combo.DstRoot "dst\$($sc.Name)_$($toolMeta.Id)"

            Write-Host ("  {0,-38}" -f $toolName) -NoNewline -ForegroundColor White

            $times   = [System.Collections.Generic.List[double]]::new()
            $errored = $false

            for ($run = 1; $run -le $Runs; $run++) {
                Write-Host "[$run]" -NoNewline -ForegroundColor DarkGray
                try {
                    $t = Measure-Tool $toolMeta.Id $toolMeta.Exe $srcDir $dstDir
                    $times.Add($t)
                    Write-Host (" {0:F1}s" -f $t) -NoNewline -ForegroundColor Gray
                } catch {
                    Write-Host " ERR" -NoNewline -ForegroundColor Red
                    Write-Host (" ($_)") -ForegroundColor DarkRed
                    $errored = $true
                    break
                }
            }

            if (-not $errored) {
                $med = Median($times.ToArray())
                Write-Host ("  -> {0}  {1}" -f (Format-Secs $med), (Format-MBps $bytes $med)) `
                    -ForegroundColor Cyan
                $allResults.Add([pscustomobject]@{
                    Scenario   = $sc.Label
                    Storage    = $combo.Label
                    Tool       = $toolName
                    Files      = @(Get-ChildItem -LiteralPath $srcDir -Recurse -File).Count
                    TotalMB    = [Math]::Round($bytes / 1MB, 1)
                    MedianSec  = [Math]::Round($med, 2)
                    MBps       = [Math]::Round($bytes / $med / 1MB, 1)
                    AllRunsSec = ($times | ForEach-Object { [Math]::Round($_, 2) }) -join ','
                })
            }
        }
    }
} } # end if ($Mode -ne 'delete')

# ---------------------------------------------------------------------------
# Tray helper functions — start/stop/warm the RoboExtension tray process so
# CLI invocations delegate over IPC instead of cold-starting each time.
# ---------------------------------------------------------------------------
$script:TrayProcess = $null

function Start-RoboExtensionTray([string]$exe) {
    Write-Host '  [tray] Starting RoboExtension tray process...' -ForegroundColor DarkGray
    $script:TrayProcess = Start-Process $exe -PassThru -WindowStyle Hidden
    # Poll the named pipe until the server is ready (up to 10 s).
    $pipeName  = 'RoboExtension_IPC'
    $deadline  = [DateTime]::UtcNow.AddSeconds(10)
    $connected = $false
    while ([DateTime]::UtcNow -lt $deadline) {
        try {
            $pipe = [System.IO.Pipes.NamedPipeClientStream]::new('.', $pipeName,
                [System.IO.Pipes.PipeDirection]::InOut)
            $pipe.Connect(200)
            $pipe.Dispose()
            $connected = $true
            break
        } catch { Start-Sleep -Milliseconds 200 }
    }
    if (-not $connected) {
        Write-Warning '[tray] Pipe not ready after 10s -- cold-start fallback active.'
        return
    }
    # Warm-up run: delete an empty temp dir to trigger JIT of the hot paths.
    $tmp = Join-Path $env:TEMP ("re_warmup_$(Get-Random)")
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    $p = Start-Process $exe -ArgumentList @('delete', "`"$tmp`"", '--f', '--yes') -PassThru -WindowStyle Hidden
    $p.WaitForExit(5000) | Out-Null
    if (Test-Path $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host "  [tray] Ready (PID $($script:TrayProcess.Id), pipe warm)." -ForegroundColor DarkGray
}

function Stop-RoboExtensionTray {
    if ($null -ne $script:TrayProcess -and -not $script:TrayProcess.HasExited) {
        Write-Host "  [tray] Stopping tray process (PID $($script:TrayProcess.Id))..." -ForegroundColor DarkGray
        try { $script:TrayProcess.Kill() } catch {}
        $script:TrayProcess.WaitForExit(5000) | Out-Null
    }
    $script:TrayProcess = $null
}

if ($Mode -ne 'copy') {
# ---------------------------------------------------------------------------
# Delete benchmark loop (SSD + HDD when available)
# ---------------------------------------------------------------------------
$deleteResults = [System.Collections.Generic.List[pscustomobject]]::new()

$deleteCombos = [System.Collections.Generic.List[pscustomobject]]::new()
$deleteCombos.Add([pscustomobject]@{ Label='SSD (delete)'; Path=$SsdPath })
if ($HddPath) { $deleteCombos.Add([pscustomobject]@{ Label='HDD (delete)'; Path=$HddPath }) }

# Start the tray process so RE delegations skip cold-start
$reInDeleteList = $deleteToolList.Keys | Where-Object { $_ -like '*RoboExtension*' }
if ($reInDeleteList) { Start-RoboExtensionTray $roboExePath }
try {
foreach ($dc in $deleteCombos) {
    Write-Host ("`n`n=== DELETE BENCHMARK ({0}) - Permanent Delete ===" -f $dc.Label.Split(' ')[0]) -ForegroundColor Cyan

    foreach ($sc in $scenarioMeta) {
        if ($ScenarioFilter -and $sc.Name -notlike "*$ScenarioFilter*") { continue }
        $srcDir = Join-Path $dc.Path "src\$($sc.Name)"
        $bytes  = $scenBytes[$sc.Name]
        $files  = @(Get-ChildItem -LiteralPath $srcDir -Recurse -File).Count

        Write-Host ("`n--- {0} ---" -f $sc.Label) -ForegroundColor Yellow

        foreach ($kv in $deleteToolList.GetEnumerator()) {
            $toolName = $kv.Key
            $toolMeta = $kv.Value

            Write-Host ("  {0,-38}" -f $toolName) -NoNewline -ForegroundColor White

            $times   = [System.Collections.Generic.List[double]]::new()
            $errored = $false

            for ($run = 1; $run -le $Runs; $run++) {
                Write-Host "[$run]" -NoNewline -ForegroundColor DarkGray
                try {
                    $t = Measure-Delete $toolMeta.Id $toolMeta.Exe $srcDir $dc.Path
                    $times.Add($t)
                    Write-Host (" {0:F1}s" -f $t) -NoNewline -ForegroundColor Gray
                } catch {
                    Write-Host " ERR" -NoNewline -ForegroundColor Red
                    Write-Host (" ($_)") -ForegroundColor DarkRed
                    $errored = $true
                    break
                }
            }

            if (-not $errored) {
                $med = Median($times.ToArray())
                Write-Host ("  -> {0}" -f (Format-Secs $med)) -ForegroundColor Cyan
                $deleteResults.Add([pscustomobject]@{
                    Scenario   = $sc.Label
                    Storage    = $dc.Label
                    Tool       = $toolName
                    Files      = $files
                    TotalMB    = [Math]::Round($bytes / 1MB, 1)
                    MedianSec  = [Math]::Round($med, 2)
                    MBps       = 'N/A'
                    AllRunsSec = ($times | ForEach-Object { [Math]::Round($_, 2) }) -join ','
                })
            }
        }
    }
}
$allResults.AddRange($deleteResults)
} finally {
    if ($reInDeleteList) { Stop-RoboExtensionTray }
}
} # end if ($Mode -ne 'copy')

# ---------------------------------------------------------------------------
# Summary table
# ---------------------------------------------------------------------------
Write-Host ("`n`n{0}" -f ('=' * 80)) -ForegroundColor Green
Write-Host "  BENCHMARK RESULTS (black-box) - median of $Runs runs" -ForegroundColor Green
Write-Host ('=' * 80) -ForegroundColor Green
$allResults | Format-Table `
    @{ L='Storage';  E={ $_.Storage  }; W=10 },
    @{ L='Scenario'; E={ $_.Scenario }; W=42 },
    @{ L='Tool';     E={ $_.Tool     }; W=38 },
    @{ L='Time';     E={ Format-Secs $_.MedianSec }; W=10 },
    @{ L='MB/s';     E={ '{0:F1}' -f $_.MBps };      W=8  } `
    -AutoSize

# ---------------------------------------------------------------------------
# JSON export
# ---------------------------------------------------------------------------
$cpu = try { (Get-CimInstance Win32_Processor -ErrorAction Stop).Name } catch { 'unknown' }
$os  = try { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption } catch { 'unknown' }
$ver = try {
    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($roboExePath)
    $vi.FileVersion
} catch { 'unknown' }

$export = [pscustomobject]@{
    GeneratedAt      = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    Machine          = $env:COMPUTERNAME
    OS               = $os
    CPU              = $cpu
    RoboExtensionExe = $roboExePath
    RoboExtensionVer = $ver
    Runs             = $Runs
    Results          = $allResults
}

$export | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputJson -Encoding UTF8
Write-Host "`nResults written to: $OutputJson" -ForegroundColor Green
