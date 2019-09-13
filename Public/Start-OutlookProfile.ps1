function Start-OutlookProfile {
    param(
        [string] $RemoveAccount,
        [string] $PrimaryAccount,
        [string] $BackupPath = "$($env:USERPROFILE)\Desktop",
        [switch] $GUI,
        [switch] $DisplayProgress,
        [switch] $NoBackup,
        [switch] $DebugOutput,
        [switch] $WhatIf
    )


    $MainKey = [ordered] @{
        'Outlook 2016' = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles'
        'Outlook 2013' = 'HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Profiles'
    }

    $AllData = foreach ($OutlookVersion in $MainKey.Keys) {
        $OutlookRegistryKey = $MainKey.$OutlookVersion

        # Find All Outlook Profiles
        $OutlookProfiles = Get-ChildItem -path Registry::$OutlookRegistryKey -ErrorAction SilentlyContinue | Select *
        if ($null -eq $OutlookProfiles) {
            continue
        }


        $i = 1
        $Outlooks = foreach ($Outlook in $OutlookProfiles) {
            [pscustomobject] @{
                Number      = $i

                ProfileName = $Outlook.PSChildName
                ParentPath  = $Outlook.PSParentPath
                Path        = $Outlook.pSPath
            }
            $i++
        }

        # Loop thru one or more Outlook Profiles
        foreach ($OutlookProfile in $Outlooks) {

            $ProfilePath = $OutlookProfile.Path


            # default section - this is the key storing Default Value
            $DefaultKey = '0a0d020000000000c000000000000046'
            $RegKeyDefault = "$ProfilePath\$DefaultKey"

            $DefaultValue = (Get-ItemProperty -Path $RegKeyDefault -Name '01023d15' -ErrorAction SilentlyContinue).'01023d15'
            if ($null -ne $DefaultValue) {
                $DefaultValueHexServiceUID = Convert-BinaryToHex -Bin $DefaultValue

                if ($DisplayProgress) {
                    Write-Color -Text '[i] ', 'Default Registry Key: ', $RegKeyDefault -Color Blue, White, Green, White
                    Write-Color -Text '[i] ', 'Binary Value ', $DefaultValue -Color Blue, White, Green, White
                    Write-Color -Text '[i] ', 'Hex version ', $DefaultValueHexServiceUID -Color Blue, White, Green, White -LinesAfter 1
                }
            }

            # Scan each account in Outlook Profile
            $Array = foreach ($Outlook in $OutlookProfile) {
                $ProfilePath = $Outlook.Path
                $SpecialKey = '9375CFF0413111d3B88A00104B2A6676'
                $SubKeyProfile = "$ProfilePath\$SpecialKey"

                $OneProfile = Get-ChildItem -Path $SubKeyProfile


                $RootProfile = Get-ItemProperty -Path $SubKeyProfile
                $Val1 = $RootProfile.'{ED475418-B0D6-11D2-8C3B-00104B2A6676}'
                $Val2 = $RootProfile.'{ED475419-B0D6-11D2-8C3B-00104B2A6676}'
                $Val3 = $RootProfile.'{ED475420-B0D6-11D2-8C3B-00104B2A6676}'
                #Convert-BinaryTohex -Binary $Val1
                #Convert-BinaryTohex -Binary $Val2
                #Convert-BinaryTohex -Binary $Val3

                if ($DisplayProgress) {
                    Write-Color '[i] ', 'Profile path ', $ProfilePath -Color Blue, White, Yellow
                }


                foreach ($One in $OneProfile) {

                    $AccountName = (Get-ItemProperty -Path Registry::$One -Name 'Account Name' -ErrorAction SilentlyContinu).'Account Name'
                    $ServiceUID = (Get-ItemProperty -Path Registry::$One -Name 'Service UID' -ErrorAction SilentlyContinu).'Service UID'
                    $HexServiceUID = Convert-BinaryToHex -Binary $ServiceUID

                    $PreferencesUID = (Get-ItemProperty -Path Registry::$One -Name 'Preferences UID' -ErrorAction SilentlyContinue).'Preferences UID'
                    $HexPreferencesUID = Convert-BinaryToHex -Binary $PreferencesUID

                    $XPProviderUID = (Get-ItemProperty -Path Registry::$One -Name 'XP Provider UID' -ErrorAction SilentlyContinue).'XP Provider UID'
                    $HexXPProviderUID = Convert-BinaryToHex -Binary $XPProviderUID

                    if ($DisplayProgress) {
                        Write-Color '[Account] ', 'Name ', $AccountName, ' Service UID ', $ServiceUID, ' Hex version ', $HexServiceUID -Color Blue, White, Yellow, White, Yellow, WHite, Yellow
                    }
                    $RegKey = "$ProfilePath\$HexServiceUID"

                    # Find Service UID of account that will be used or is already set as Primary Account
                    $MyValue = (Get-ItemProperty -Path $RegKey -Name '01023d15' -ErrorAction SilentlyContinue).'01023d15'
                    $MyValueHexServiceUID = Convert-BinaryToHex -Binary $MyValue

                    if ($DisplayProgress) {
                        Write-Color '[Account Update] ', 'Key: ', $RegKey, ' Primary Service UID ', $MyValue, ' Hex ', $MyValueHexServiceUID -Color Blue, White, Green, White, Green, White, Green
                    }
                    [PsCustomobject] @{
                        OutlookVersion     = $OutlookVersion
                        Profile            = $OutlookProfile.ProfileName
                        ProfileNumber      = $One.PSChildName
                        AccountName        = $AccountName
                        #ServiceUIDBefore  = $ServiceUID
                        ServiceUID         = $HexServiceUID
                        #PrimaryServiceUIDBefore = $MyValue
                        RequiredServiceUID = $MyValueHexServiceUID
                        #PreferencesUIDBefore = $PreferencesUID
                        PreferencesUID     = $HexPreferencesUID

                        XPProviderUID      = $HexXPProviderUID
                        ProfilePath        = $ProfilePath

                    }


                }
            }
            if ($DisplayProgress) {
                Write-Color -LinesAfter 1
            }
            $Array
        }
    }
    <#
    if (-not $GUI) {
        # Assing all profiles
        $OutlookProfiles = $Outlooks
    } else {
        # Show GUI
        $Line = '==================================='
        do {
            Clear-Host
            Write-Color $line -LinesBefore 1
            Write-Color 'Outlook Profile Fixer' -C Green -StartTab 1
            Write-Color $line

            foreach ($Outlook in $Outlooks) {
                Write-Color -Text $Outlook.Number, ' - Profile Name: ', $Outlook.ProfileName -Color Yellow, White, Green
            }
            Write-Color '0', ' - ', 'Quit' -Color Yellow, White, Green -LinesAfter 1

            $Input = Read-Host 'Select'
            If ($Input -eq 0) {
                Exit
            } elseif ($Outlooks.Number -contains $Input) {
                break
            } else {
                Write-Color 'Wrong choice.', ' Press any key to restart!' -Color Red, Yellow -LinesBefore 1
                [void][System.Console]::ReadKey($true)
            }
        } while ($Input -ne '0')
        Clear-Host
        $OutlookProfiles = $Outlooks[$Input - 1]
        # End Gui
    }
    #>

    if (-not $NoBackup) {
        # Backup
        $Backups = foreach ($OutlookVersion in $MainKey.Keys) {
            $OutlookRegistryKey = $MainKey.$OutlookVersion

            [string] $BackupName = "$OutlookVersion-RegistryProfile"
            # Make registry Backup
            #Write-Color "[i] ", 'Backup of Key ', $OutlookRegistryKey, ' to ', $BackupPath -Color Blue, White, Yellow, White, Yellow
            $Backup = Backup-RegistryPath -Key $OutlookRegistryKey -BackupPath $BackupPath -BackupName $BackupName
            if ($null -ne $Backup) {
                #try {
                #    Write-Color "[i] ", "Backup of Outlook profiles made to ", $Backup -Color Blue, White, Yellow -LinesAfter 1
                #} catch {
                    Write-Color "[i] Backup of Outlook profiles made to $Backup"
                #}
            } else {

            }
        }

    }

    foreach ($Mail in $AllData) {
        if ($Mail.ProfilePath) {
            if ($Mail.PreferencesUID -and $Mail.XPProviderUID -and $Mail.PreferencesUID -and $Mail.ProfileNumber) {
                # Check if user wants to remove any account
                if ($RemoveAccount) {
                    if ($Mail.AccountName -match $RemoveAccount) {

                        $Keys = @(
                            "$($Mail.ProfilePath)\$($Mail.PreferencesUID)"
                            "$($Mail.ProfilePath)\$($Mail.ServiceUID)"
                            "$($Mail.ProfilePath)\$($Mail.XPProviderUID)"
                            "$($Mail.ProfilePath)\9375CFF0413111d3B88A00104B2A6676\$($Mail.ProfileNumber)"
                        )

                        foreach ($Key in $Keys) {
                            #try {
                            #    Write-Color '[i] ', "Removing key ", $Key -Color Blue, White, Yellow
                            #} catch {
                                Write-host "[i] Removing key $Key"
                            #}
                            if (-Not $WhatIf) {
                                Remove-Item -Path $Key -Confirm:$false #-WhatIf
                            }
                        }

                    }
                }
                if ($PrimaryAccount) {
                    if ($Mail.AccountName -match $PrimaryAccount) {
                        $Default = "$($Mail.ProfilePath)\0a0d020000000000c000000000000046"
                        if ($Mail.RequiredServiceUID) {
                            #Try {
                            #    Write-Color "[i] ", "Setting default profile ", $Default, ' with ', $Mail.RequiredServiceUID -Color Blue, White, Yellow, White, Green
                            #} catch {
                                Write-Host "[i] Setting default profile $Default with $($Mail.RequiredServiceUID)"
                            #}
                            if (-not $WhatIf) {
                                Set-ItemProperty -Path $Default -Name '01023d15' -Value $Mail.MyValue -Type Binary #-WhatIf
                            }
                        }

                        [byte[]] $ByteArray = Convert-HexToBinary -Hex $Mail.ProfileNumber
                        $SubValue = "$($Mail.ProfilePath)\9375CFF0413111d3B88A00104B2A6676"
                        #Try {
                        #    Write-Color "[i] ", "Setting default profile ", $SubValue, ' in ', "{ED475418-B0D6-11D2-8C3B-00104B2A6676}", ' with ', $ByteArray -Color Blue, White, Yellow, White, Green, White, Yellow
                        #} catch {
                            Write-Color "[i] Setting default profile $SubValue in {ED475418-B0D6-11D2-8C3B-00104B2A6676} with $ByteArray"
                        #}
                        if (-not $WhatIf) {
                            Set-ItemProperty -Path $SubValue -Name "{ED475418-B0D6-11D2-8C3B-00104B2A6676}" -Value $ByteArray -Type Binary #-WhatIf
                        }
                    }
                }
            }
        }
    }

    # Remove leftovers
    $LeftOvers = Find-OutlookKeys
    foreach ($Left in $LeftOvers) {
        # Check if user wants to remove any account
        if ($RemoveAccount) {
            if ($Left.Email -match "$RemoveAccount") {
                #try {
                #    Write-Color '[i] ', 'Removing leftovers key ', $Left.RegistryKey -Color Blue, White, Yellow
                #} catch {
                    Write-Host "[i] Removing leftovers key $($Left.RegistryKey)"
                #}
                if (-not $WhatIf) {
                    Remove-Item -Path $Left.RegistryKey -Recurse -Confirm:$False
                }
            }
        }
    }
    if ($DebugOutput) {
        $AllData
    }
}