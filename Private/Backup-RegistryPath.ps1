function Backup-RegistryPath {
    param(
        [string] $Key,
        [string] $BackupPath = "$($env:USERPROFILE)\Desktop",
        [string] $BackupName
    )

    $Date = Get-Date
    $FileName = "$BackupName-$($Date.Year)-$($Date.Month)-$($Date.Day).$($Date.Hour).$($Date.Minute).$($Date.Second).reg"

    $BackupPlace = "$BackupPath\$FileName"

    if (Test-Path -Path $BackupPlace) {
        return $null
    } else {
        try {
            if (Test-Path Registry::$Key) {
                $Registry = Start-MyProgram -Program 'reg.exe' -cmdArgList "export", "$Key", "$BackupPlace"
                return $BackupPlace
            }
        } catch {

            return $null
        }
    }
}