function Find-OutlookKeys {
    param(
        $OutlookVersion = '2016',
        $EmailFind
    )
    $MainKey = [ordered] @{
        'Outlook 2016' = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles'
        'Outlook 2013' = 'HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Profiles'
    }


    $AllData = foreach ($OutlookVersion in $MainKey.Keys) {
        $OutlookRegistryKey = $MainKey.$OutlookVersion

        if ($OutlookVersion -eq 'Outlook 2016') {
            $SearchValue = '001f6641'
        } elseif ($OutlookVersion -eq 'Outlook 2013') {
            $SearchValue = '001f662b'
        } else {
            Exit
        }
        <#
Property      : {001f300a, 001f3d13, 00033e03, 00033009...}
PSPath        : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\Nowy\cb755c91ea7e0b4c97fca67db4f0486b
PSParentPath  : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\Nowy
PSChildName   : cb755c91ea7e0b4c97fca67db4f0486b
PSDrive       : HKCU
PSProvider    : Microsoft.PowerShell.Core\Registry
PSIsContainer : True
SubKeyCount   : 0
View          : Default
Handle        : Microsoft.Win32.SafeHandles.SafeRegistryHandle
ValueCount    : 20
Name          : HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\Nowy\cb755c91ea7e0b4c97fca67db4f0486b
#>

        $RegistryKey = Get-ChildItem -Path "Registry::$OutlookRegistryKey" -Recurse
        $Special = $RegistryKey | Where-Object { $_.Property -eq $SearchValue } #| Select -Last 1 *
        $Keys = foreach ($S in $Special) {
            #$Special.PSPath
            $Path = "$($S.Name)"
            <#

    001f6641     : {83, 0, 77, 0...}
    PSPath       : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\Nowy\0666d1f4813a9a4ba2d9462100225710
    PSParentPath : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\Nowy
    PSChildName  : 0666d1f4813a9a4ba2d9462100225710
    PSProvider   : Microsoft.PowerShell.Core\Registry
    #>
            $MyValue = Get-ItemProperty -Path Registry::$Path -Name $SearchValue
            #$Email = [System.Text.Encoding]::Unicode.GetString( )
            $Email = Convert-BinaryToString ($MyValue.$SearchValue)

            $ParentPath = $S.PSParentPath
            $ChildName = $S.PSChildName
            #Write-Color "Path ", $ParentPath, " will delete value ", $ChildName, ' email ', $Email -Color White, Yellow, White, Yellow

            if ($ChildName -ne 'GroupsStore') {
                [PsCustomObject] @{
                    OutlookVersion  = $OutlookVersion
                    Profile         = ($ParentPath -split '\\')[-1]
                    ProfilePath     = $ParentPath
                    RegistryKeyName = $ChildName
                    RegistryKey     = "$ParentPath\$ChildName"
                    Email           = ($Email -replace 'SMTP:', '').ToLower()
                }
            }
            <#
            if ($Email -like "*$EmailFind*") {
                $ParentPath = $S.PSParentPath
                $ChildName = $S.PSChildName
                #$Path
                #Write-Color "ParentPath ", $ParentPath, " Path: ", $Path -Color White, Yellow, White, Yellow
                if ($ChildName -ne 'GroupsStore') {
                    #Write-Color "Path ", $ParentPath, " will delete value ", $ChildName -Color White, Yellow, White, Yellow
                    # Remove-Item -Path "$ParentPath\$ChildName" -Recurse
                }
            }
            #>
            #break
        }
        #Search-Registry -KeyName '001f6641' -Recurse $Path
        return $Keys
    }
    return $AllData
}