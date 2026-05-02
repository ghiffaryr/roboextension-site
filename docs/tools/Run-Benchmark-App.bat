@echo off
:: ============================================================
::  RoboExtension Public Benchmark Launcher
::
::  Usage:
::    Run-Benchmark-App.bat [-SsdPath <path>] [-HddPath <path>] [-TeraCopyExe <path>] [-FastCopyExe <path>]
::                          [-Runs <n>] [-Mode <full|copy|delete>] [-ScenarioFilter <str>]
::                          [-ComboFilter <str>] [-ToolFilter <str>] [-ForceRegen] [-Resume]
::
::  Parameters:
::    -SsdPath        Writable folder on SSD (default: C:\BenchTemp)
::    -HddPath        Writable folder on HDD (omit to run SSD-only)
::    -TeraCopyExe    Path to TeraCopy.exe (optional; auto-detected if omitted)
::    -FastCopyExe    Path to FastCopy.exe (optional; auto-detected if omitted)
::    -Runs           Number of timed runs per scenario (default: 3)
::    -Mode           full (default), copy, or delete
::    -ScenarioFilter Comma-separated scenario name substrings to include
::    -ComboFilter    Comma-separated combo substrings to include
::    -ToolFilter     Comma-separated tool name substrings to include
::    -SmallFilesSource  Folder of real small files (1 KB-100 KB) to sample instead of generating
::    -LargeFilesSource  Folder of real large files (>=100 MB) to use instead of generating
::    -ForceRegen        Regenerate source datasets even if they already exist
::    -Resume            Keep existing rows not being re-run
::
::  Examples:
::    .\Run-Benchmark-App.bat
::    .\Run-Benchmark-App.bat -SsdPath C:\BenchTemp
::    .\Run-Benchmark-App.bat -SsdPath C:\BenchTemp -HddPath D:\BenchTemp
::    .\Run-Benchmark-App.bat -SsdPath C:\BenchTemp -Mode delete
::    .\Run-Benchmark-App.bat -SsdPath C:\BenchTemp -Runs 5 -ToolFilter RoboExtension
::    .\Run-Benchmark-App.bat -SsdPath C:\BenchTemp -HddPath D:\BenchTemp -Mode delete -Resume
::    .\Run-Benchmark-App.bat -SsdPath C:\BenchTemp -WslSource "\\wsl.localhost\Ubuntu-22.04\home\user\project"
:: ============================================================

setlocal EnableDelayedExpansion

:: Capture script directory before any shift calls corrupt %0
set "SCRIPT_DIR=%~dp0"

:: Defaults
set "SSD_PATH=C:\BenchTemp"
set "HDD_PATH="
set "TERA_EXE="
set "FAST_EXE="
set "RUNS="
set "MODE=full"
set "SCENARIO_FILTER="
set "COMBO_FILTER="
set "TOOL_FILTER="
set "SMALL_SRC="
set "LARGE_SRC="
set "FORCE_REGEN="
set "RESUME="
set "WSL_SRC="

:parse_args
if "%~1"=="" goto :args_done
if /I "%~1"=="-SsdPath"        ( set "SSD_PATH=%~2"        & shift & shift & goto :parse_args )
if /I "%~1"=="-HddPath"        ( set "HDD_PATH=%~2"        & shift & shift & goto :parse_args )
if /I "%~1"=="-TeraCopyExe"    ( set "TERA_EXE=%~2"        & shift & shift & goto :parse_args )
if /I "%~1"=="-FastCopyExe"    ( set "FAST_EXE=%~2"        & shift & shift & goto :parse_args )
if /I "%~1"=="-Runs"           ( set "RUNS=%~2"            & shift & shift & goto :parse_args )
if /I "%~1"=="-Mode"           ( set "MODE=%~2"            & shift & shift & goto :parse_args )
if /I "%~1"=="-ScenarioFilter" goto :_parse_scenario_filter
if /I "%~1"=="-ComboFilter"    goto :_parse_combo_filter
if /I "%~1"=="-ToolFilter"     goto :_parse_tool_filter
if /I "%~1"=="-SmallFilesSource" ( set "SMALL_SRC=%~2"       & shift & shift & goto :parse_args )
if /I "%~1"=="-LargeFilesSource" ( set "LARGE_SRC=%~2"       & shift & shift & goto :parse_args )
if /I "%~1"=="-ForceRegen"       ( set "FORCE_REGEN=1"       & shift         & goto :parse_args )
if /I "%~1"=="-Resume"           ( set "RESUME=1"            & shift         & goto :parse_args )
if /I "%~1"=="-WslSource"        ( set "WSL_SRC=%~2"         & shift & shift & goto :parse_args )
echo  Unknown argument: %~1 & exit /b 1

:: Filter accumulators: re-join tokens that were split by cmd.exe on commas
:_parse_scenario_filter
set "SCENARIO_FILTER=%~2"
shift & shift
:_sf_cont
if "%~1"=="" goto :args_done
set "_c=%~1"
if "!_c:~0,1!"=="-" goto :parse_args
set "SCENARIO_FILTER=!SCENARIO_FILTER!,%~1"
shift
goto :_sf_cont

:_parse_combo_filter
set "COMBO_FILTER=%~2"
shift & shift
:_cf_cont
if "%~1"=="" goto :args_done
set "_c=%~1"
if "!_c:~0,1!"=="-" goto :parse_args
set "COMBO_FILTER=!COMBO_FILTER!,%~1"
shift
goto :_cf_cont

:_parse_tool_filter
set "TOOL_FILTER=%~2"
shift & shift
:_tf_cont
if "%~1"=="" goto :args_done
set "_c=%~1"
if "!_c:~0,1!"=="-" goto :parse_args
set "TOOL_FILTER=!TOOL_FILTER!,%~1"
shift
goto :_tf_cont

:args_done

:: ── Locate RoboExtension.exe ─────────────────────────────────────────────
set "ROBO_EXE="

:: Try common locations: relative publish\, dist\, installed paths
for %%P in (
    "%ProgramFiles%\RoboExtension\RoboExtension.exe"
) do (
    if exist %%P (
        set "ROBO_EXE=%%~fP"
        goto :found_robo
    )
)
echo.
echo  ERROR: RoboExtension.exe not found.
echo         Build the project first (run build.bat in the repo root),
echo         or install RoboExtension from https://github.com/ghiffaryr/RoboExtension
echo.
pause
exit /b 1
:found_robo

:: ── Verify companion PS1 is present ─────────────────────────────────────
if not exist "%SCRIPT_DIR%Run-Benchmark-App.ps1" (
    echo.
    echo  ERROR: Run-Benchmark-App.ps1 not found.
    echo         Run-Benchmark-App.bat and Run-Benchmark-App.ps1 must be in the same folder.
    echo         Expected: %SCRIPT_DIR%Run-Benchmark-App.ps1
    echo.
    pause
    exit /b 1
)

:: ── Self-elevate if not running as Administrator ──────────────────────────
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  Requesting Administrator privileges for page-cache flush...
    echo.
    set "_TMP=%TEMP%\bench_run_%RANDOM%.ps1"
    (
        echo $p = @{ SsdPath = '!SSD_PATH!'; RoboExe = '!ROBO_EXE!'; Mode = '!MODE!'; OutputJson = '!SCRIPT_DIR!results-app.json' }
        echo if ^('!HDD_PATH!'^)        { $p.HddPath        = '!HDD_PATH!' }
        echo if ^('!TERA_EXE!'^)        { $p.TeraCopyExe    = '!TERA_EXE!' }
        echo if ^('!FAST_EXE!'^)        { $p.FastCopyExe    = '!FAST_EXE!' }
        echo if ^('!RUNS!'^)            { $p.Runs           = [int]'!RUNS!' }
        echo if ^('!SCENARIO_FILTER!'^) { $p.ScenarioFilter = '!SCENARIO_FILTER!' }
        echo if ^('!COMBO_FILTER!'^)    { $p.ComboFilter    = '!COMBO_FILTER!' }
        echo if ^('!TOOL_FILTER!'^)     { $p.ToolFilter       = '!TOOL_FILTER!' }
        echo if ^('!SMALL_SRC!'^)       { $p.SmallFilesSource = '!SMALL_SRC!' }
        echo if ^('!LARGE_SRC!'^)       { $p.LargeFilesSource = '!LARGE_SRC!' }
        echo if ^('!FORCE_REGEN!'^)     { $p.ForceRegen       = $true }
        echo if ^('!RESUME!'^)          { $p.Resume         = $true }
        echo if ^('!WSL_SRC!'^)         { $p.WslSource       = '!WSL_SRC!' }
        echo ^& '!SCRIPT_DIR!Run-Benchmark-App.ps1' @p
    ) > "!_TMP!"
    powershell -NoProfile -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -NoExit -File ""!_TMP!""' -Verb RunAs -Wait"
    del "!_TMP!" 2>nul
    exit /b
)

:: ── Build PowerShell argument string ─────────────────────────────────────
set "PS_ARGS=-SsdPath "%SSD_PATH%" -RoboExe "%ROBO_EXE%" -Mode %MODE%"
if not "%HDD_PATH%"==""        set "PS_ARGS=%PS_ARGS% -HddPath "%HDD_PATH%""
if not "%TERA_EXE%"==""        set "PS_ARGS=%PS_ARGS% -TeraCopyExe "%TERA_EXE%""
if not "%FAST_EXE%"==""        set "PS_ARGS=%PS_ARGS% -FastCopyExe "%FAST_EXE%""
if not "%RUNS%"==""            set "PS_ARGS=%PS_ARGS% -Runs %RUNS%"
if not "%SCENARIO_FILTER%"==""    set "PS_ARGS=%PS_ARGS% -ScenarioFilter "%SCENARIO_FILTER%""
if not "%COMBO_FILTER%"==""    set "PS_ARGS=%PS_ARGS% -ComboFilter "%COMBO_FILTER%""
if not "%TOOL_FILTER%"==""     set "PS_ARGS=%PS_ARGS% -ToolFilter "%TOOL_FILTER%""
if not "%SMALL_SRC%"==""       set "PS_ARGS=%PS_ARGS% -SmallFilesSource "%SMALL_SRC%""
if not "%LARGE_SRC%"==""       set "PS_ARGS=%PS_ARGS% -LargeFilesSource "%LARGE_SRC%""
if defined FORCE_REGEN         set "PS_ARGS=%PS_ARGS% -ForceRegen"
if defined RESUME              set "PS_ARGS=%PS_ARGS% -Resume"
if not "%WSL_SRC%"==""        set "PS_ARGS=%PS_ARGS% -WslSource "%WSL_SRC%""

echo.
echo  ================================================================
echo   RoboExtension Benchmark
echo  ================================================================
echo   SSD path : %SSD_PATH%
if not "%HDD_PATH%"==""     echo   HDD path : %HDD_PATH%
if not "%TERA_EXE%"==""     echo   TeraCopy : %TERA_EXE%
if not "%FAST_EXE%"==""     echo   FastCopy : %FAST_EXE%
echo   Binary   : %ROBO_EXE%
echo   Mode     : %MODE%
if not "%RUNS%"==""         echo   Runs     : %RUNS%
if not "%SCENARIO_FILTER%"=="" echo   Scenario : %SCENARIO_FILTER%
if not "%COMBO_FILTER%"==""    echo   Combo    : %COMBO_FILTER%
if not "%TOOL_FILTER%"==""     echo   Tools    : %TOOL_FILTER%
if not "%SMALL_SRC%"==""       echo   SmallSrc : %SMALL_SRC%
if not "%LARGE_SRC%"==""       echo   LargeSrc : %LARGE_SRC%
if not "%WSL_SRC%"==""         echo   WslSource: %WSL_SRC%
echo   Output   : %SCRIPT_DIR%results-app.json
echo  ================================================================
echo.
if "%MODE%"=="delete" (
    echo  Running delete benchmark only...
) else if "%MODE%"=="copy" (
    echo  Running copy benchmark only...
) else (
    echo  Running full benchmark ^(copy + delete^)...
)
if defined FORCE_REGEN echo  ForceRegen: source datasets will be regenerated
if defined RESUME      echo  Resume: preserving existing rows from results-app.json
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Run-Benchmark-App.ps1" ^
    %PS_ARGS% ^
    -OutputJson "%SCRIPT_DIR%results-app.json"

echo.
if %errorlevel% equ 0 (
    echo  Done! Results written to: %SCRIPT_DIR%results-app.json
) else (
    echo  Benchmark failed with exit code %errorlevel%.
)
echo.
pause
