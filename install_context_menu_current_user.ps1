param()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$launcher = Join-Path $scriptDir "open_with_fenjiubihe.ps1"
$registered = 0
$failed = 0

function Get-ProgId([string]$ext) {
    $assocOutput = cmd /c "assoc $ext" 2>$null
    if (-not $assocOutput) {
        return $null
    }
    foreach ($line in $assocOutput) {
        if ($line -like "$ext=*") {
            return ($line -split "=", 2)[1]
        }
    }
    return $null
}

function Register-MenuRoot([string]$rootKey) {
    try {
        if (-not (Test-Path $rootKey)) {
            New-Item -Path $rootKey -Force | Out-Null
        }

        $legacyKey = Join-Path $rootKey "FenJiuBiHe.Open"
        if (Test-Path $legacyKey) {
            Remove-Item -Path $legacyKey -Recurse -Force -ErrorAction SilentlyContinue
        }

        $entries = @(
            @{ Key = "FenJiuBiHe.SetFirst"; Text = "Set as File 1"; Args = '--context-set-first "%1"' },
            @{ Key = "FenJiuBiHe.CompareWithFirst"; Text = "Compare with File 1"; Args = '--context-compare "%1"' },
            @{ Key = "FenJiuBiHe.ClearFirst"; Text = "Clear File 1"; Args = '--context-clear "%1"' }
        )

        foreach ($entry in $entries) {
            $baseKey = Join-Path $rootKey $entry.Key
            $cmdKey = Join-Path $baseKey "command"
            New-Item -Path $baseKey -Force | Out-Null
            New-Item -Path $cmdKey -Force | Out-Null
            Set-ItemProperty -Path $baseKey -Name '(default)' -Value $entry.Text
            Set-ItemProperty -Path $baseKey -Name 'Icon' -Value "powershell.exe"
            Set-ItemProperty -Path $cmdKey -Name '(default)' -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$launcher`" $($entry.Args)"
        }
        return $true
    } catch {
        Write-Host ("Write failed: " + $rootKey)
        Write-Host $_.Exception.Message
        return $false
    }
}

foreach ($ext in @(".xml", ".xlsx", ".csv")) {
    foreach ($root in @(
        "HKCU:\Software\Classes\$ext\shell",
        "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell"
    )) {
        if (Register-MenuRoot $root) {
            $registered++
        } else {
            $failed++
        }
    }

    $progId = Get-ProgId $ext
    if ($progId) {
        $progRoot = "HKCU:\Software\Classes\$progId\shell"
        if (Register-MenuRoot $progRoot) {
            $registered++
        } else {
            $failed++
        }
    }
}

if ($registered -eq 0) {
    Write-Host "No context menu entries were written."
    exit 1
}

Write-Host ("Installed two-step context menu entries into {0} root locations." -f $registered)
if ($failed -gt 0) {
    Write-Host ("{0} root locations failed, but the menu can still work if at least one root succeeded." -f $failed)
}
exit 0
