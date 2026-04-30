param()

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

function Remove-MenuRoot([string]$rootKey) {
    foreach ($name in @("FenJiuBiHe.Open", "FenJiuBiHe.SetFirst", "FenJiuBiHe.CompareWithFirst", "FenJiuBiHe.ClearFirst")) {
        $target = Join-Path $rootKey $name
        if (Test-Path $target) {
            Remove-Item -Path $target -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

$removed = 0
foreach ($ext in @(".xml", ".xlsx", ".csv")) {
    foreach ($root in @(
        "HKCU:\Software\Classes\$ext\shell",
        "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell"
    )) {
        Remove-MenuRoot $root
        $removed++
    }

    $progId = Get-ProgId $ext
    if ($progId) {
        Remove-MenuRoot ("HKCU:\Software\Classes\$progId\shell")
        $removed++
    }
}

Write-Host ("Removed context menu entries from {0} root locations." -f $removed)
exit 0
