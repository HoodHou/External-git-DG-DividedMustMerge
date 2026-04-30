param(
    [string]$Mode = "",
    [string]$Target = ""
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$exePath = $null
$distDir = Join-Path $scriptDir "dist"
if (Test-Path $distDir) {
    $exePath = Get-ChildItem -Path $distDir -Recurse -Filter "*.exe" -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
}
$stateDir = if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA "FenJiuBiHe" } else { Join-Path $scriptDir "runtime" }
$stateFile = Join-Path $stateDir "context_first_file.txt"

function Show-Info([string]$message) {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [System.Windows.Forms.MessageBox]::Show($message, "FenJiuBiHe") | Out-Null
}

function Resolve-PythonCommand {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return @($python.Source)
    }
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return @($py.Source, "-3")
    }
    return $null
}

function Launch-App([string[]]$paths) {
    if (Test-Path $exePath) {
        Start-Process -FilePath $exePath -ArgumentList $paths | Out-Null
        return
    }
    $pythonCmd = Resolve-PythonCommand
    if (-not $pythonCmd) {
        throw "Python not found and exe was not built."
    }
    $appPath = Join-Path $scriptDir "app.py"
    if ($pythonCmd.Count -eq 1) {
        Start-Process -FilePath $pythonCmd[0] -ArgumentList @($appPath) + $paths | Out-Null
        return
    }
    Start-Process -FilePath $pythonCmd[0] -ArgumentList @($pythonCmd[1], $appPath) + $paths | Out-Null
}

try {
    switch ($Mode) {
        "--context-set-first" {
            if (-not $Target) {
                Show-Info "No file path was provided for File 1."
                exit 1
            }
            New-Item -ItemType Directory -Force -Path $stateDir | Out-Null
            Set-Content -Path $stateFile -Value ([System.IO.Path]::GetFullPath($Target)) -Encoding UTF8
            Show-Info ("Selected File 1: " + [System.IO.Path]::GetFileName($Target))
            exit 0
        }
        "--context-compare" {
            if (-not $Target) {
                Show-Info "No second file was provided for comparison."
                exit 1
            }
            if (-not (Test-Path $stateFile)) {
                Show-Info "Please select File 1 first."
                exit 1
            }
            $firstFile = (Get-Content -Path $stateFile -Raw -Encoding UTF8).Trim()
            if (-not $firstFile) {
                Show-Info "Stored File 1 is empty. Please select it again."
                exit 1
            }
            $secondFile = [System.IO.Path]::GetFullPath($Target)
            if ($firstFile -ieq $secondFile) {
                Show-Info "File 1 and File 2 are the same file."
                exit 1
            }
            Launch-App @($firstFile, $secondFile)
            exit 0
        }
        "--context-clear" {
            if (Test-Path $stateFile) {
                Remove-Item -Path $stateFile -Force -ErrorAction SilentlyContinue
            }
            Show-Info "Cleared File 1."
            exit 0
        }
        default {
            $paths = @()
            if ($Mode) {
                $paths += $Mode
            }
            if ($Target) {
                $paths += $Target
            }
            Launch-App $paths
            exit 0
        }
    }
} catch {
    Show-Info $_.Exception.Message
    exit 1
}
