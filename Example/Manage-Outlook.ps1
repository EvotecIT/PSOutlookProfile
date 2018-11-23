Clear-Host

Import-Module PSOutlookProfile -Force

Start-OutlookProfile -WhatIf -RemoveAccount 'przemyslaw.klys@euvic.pl' -PrimaryAccount 'przemyslaw.klys@evotec.pl'

#Start-OutlookProfile -NoBackup -WhatIf -RemoveAccount 'przemyslaw.klys@euvic.pl' -PrimaryAccount 'przemyslaw.klys@evotec.pl'