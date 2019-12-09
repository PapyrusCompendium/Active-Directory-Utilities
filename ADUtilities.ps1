$script:banner = "
 






      ___       _   _            ______ _               _                     _   _ _   _ _ _ _   _           
     / _ \     | | (_)           |  _  (_)             | |                   | | | | | (_) (_) | (_)          
    / /_\ \ ___| |_ ___   _____  | | | |_ _ __ ___  ___| |_ ___  _ __ _   _  | | | | |_ _| |_| |_ _  ___  ___ 
    |  _  |/ __| __| \ \ / / _ \ | | | | | '__/ _ \/ __| __/ _ \| '__| | | | | | | | __| | | | __| |/ _ \/ __|
    | | | | (__| |_| |\ V /  __/ | |/ /| | | |  __/ (__| || (_) | |  | |_| | | |_| | |_| | | | |_| |  __/\__ \
    \_| |_/\___|\__|_| \_/ \___| |___/ |_|_|  \___|\___|\__\___/|_|   \__, |  \___/ \__|_|_|_|\__|_|\___||___/
                                                                      __/ /                                  
                                                                     |___/                 - PapyrusCompendium
"

#Author: PapyrusCompendium
#Description: Help with quick utility functions in Active Directory. Mainly this is for my own use.
Import-Module "$(Get-Location)\ADUtiliesModule.ps1"
$script:foundADObject
$script:alternateCreds

function HelpDisplay {
    "   ┌────────────────────────┐"
    "   │------Search Users------│"
    "   │By Name:               0│"
    "   │By Username:           1│"
    "   │By Employee ID:        2│"
    "   │                        │"
    "   │----Search Computers----│"
    "   │By Computer Name       3│"
    "   │By IP Address          4│"
    "   │                        │"
    "   │---------Config---------│"
    "   │Use Alternate Creds    5│"
    "   │                        │"
    "   ├────────────────────────┤"
    "   │< Exit Script          #│"
    "   └────────────────────────┘"
}

function DisplayUserData{
    "  ┌────────────────────────┐       ┌───────────────────────Requested Data──────────────────────"
    "  │------Display Data------│       │"
    "  │Employee ID:           0│       │"
    "  │Phones:                1│       │"
    "  │Email:                 2│       │"
    "  │Full Name:             3│       │"
    "  │Account Creation:      4│       │"
    "  │Password Details:      5│       │"
    "  │Custom Query:          6│       │"
    "  │All Data:              *│       │"
    "  │                        │       │"
    "  ├────────────────────────┤       │"
    "  │-----Account Configs----│       │"
    "  │Change Employee ID:    7│       │"
    "  │Change Password        8│       │"
    "  │                        │       │"
    "  ├────────────────────────┤       │"
    "  │< Go Back              #│       │"
    "  └────────────────────────┘"
}

function DisplayComputerData{
    "  ┌────────────────────────┐       ┌───────────────────────Requested Data──────────────────────"
    "  │------Display Data------│       │"
    "  │Get Last User:         0│       │"
    "  │Team Viewer  ID:       1│       │"
    "  │Dell Service Tag       2│       │"
    "  │Installed Software:    3│       │"
    "  │Custom Query:          4│       │"
    "  │All Data:              *│       │"
    "  ├────────────────────────┤       │"
    "  │! Restart Machine      R│       │"
    "  │< Go Back              #│       │"
    "  └────────────────────────┘"
}

function ClearBanner{
    Clear-Host
    $script:banner
}

function SetColourTheme{
    $host.ui.RawUI.WindowTitle = "Active Directory Utilities"
    $host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
    $host.UI.RawUI.ForegroundColor = 'Green'
    $host.PrivateData.ErrorForegroundColor = 'Red'
    $host.PrivateData.ErrorBackgroundColor = $bckgrnd
    $host.PrivateData.WarningForegroundColor = 'Magenta'
    $host.PrivateData.WarningBackgroundColor = $bckgrnd
    $host.PrivateData.DebugForegroundColor = 'Yellow'
    $host.PrivateData.DebugBackgroundColor = $bckgrnd
    $host.PrivateData.VerboseForegroundColor = 'Green'
    $host.PrivateData.VerboseBackgroundColor = $bckgrnd
    $host.PrivateData.ProgressForegroundColor = 'Cyan'
    $host.PrivateData.ProgressBackgroundColor = $bckgrnd
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Clear-Host
}

function SetCursorLocation{param($x, $y)
    $host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates $x, $y
}

function WriteInKeyPairSidePanel{param($textData)
    $lineNumber = 20
    $width = 37

    if($textData.GetType().Name -like "Object*"){
        $textData | ForEach-Object {
            SetCursorLocation $width $lineNumber
            Write-Host("$($_.Key) : $($_.Value)")
            $lineNumber++
        }
    } else {
        SetCursorLocation $width $lineNumber
        Write-Host("$($textData.Key) : $($textData.Value)")
    }

    SetCursorLocation 0 36
}

function WriteListInSidePanel{param($textData)
    $lineNumber = 20
    $width = 37

    $textData | ForEach-Object {
        SetCursorLocation $width $lineNumber
        Write-Host($_)
        $lineNumber++
    }

    SetCursorLocation 0 36
}

function WriteTextInSidePanel{param($textData)
    $lineNumber = 20
    $width = 37

    SetCursorLocation $width $lineNumber
    Write-Host($textData)

    SetCursorLocation 0 36
}

function MainMenu{
    ClearBanner
    HelpDisplay
    $optionSelection = Read-Host -Prompt "    Option"
    
    switch($optionSelection)
    {   
        "0" {SearchByName}
        "1" {SearchByUsername}
        "2" {SearchByEmployeeID}
        "3" {SearchByComputerName}
        "4" {SearchByIP}
        "5" {AlternateCredential}
        "#" {Exit}
    }
}

function UserDataMenu
{
    for(;;){
        ClearBanner
        "   $($script:foundADObject.Name) - $($script:foundADObject.SamAccountName)"
        $host.ui.RawUI.WindowTitle = "Active Directory Utilities - $($script:foundADObject.Name)"
        DisplayUserData

        $optionSelection = Read-Host -Prompt "   Option"

        ""
        switch($optionSelection)
        {
            "#" {return}
            "*" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -like "*"})}
            "0" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -eq "EmployeeID"})}
            "1" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -like "*phone*"})}
            "2" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -like "*email*"})}
            "3" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -eq "Name"})}
            "4" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {($_.Key -like "*account*") -or $_.Key -like "*created*"})}
            "5" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -like "*password*"})}
            "6" {
                $query = Read-Host -Prompt "Query"
                WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {($_.Key -like "*$($query)*") -or ($_.Value -like "*$($query)*")})
            }
            "7" {ChangeEmployeeID}
            "8" {ChangeUserPassword}
        }
        ""

        Pause
    }
}

function ComputerDataMenu
{
    for(;;){
        ClearBanner
        "   $($script:foundADObject.Name) - $($script:foundADObject.IPv4Address)"
        $host.ui.RawUI.WindowTitle = "Active Directory Utilities - $($script:foundADObject.Name)"
        DisplayComputerData

        $optionSelection = Read-Host -Prompt "   Option"
        ""
        switch($optionSelection)
        {
            "#" {return}
            "r" {RestartMachine($script:foundADObject.Name)}
            "*" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -like "*"})}
            "0" {WriteListInSidePanel(GetLastUser($script:foundADObject.Name))}
            "1" {WriteTextInSidePanel("TeamViewer ID: $(GetTeamViewerID($script:foundADObject.Name))")}
            "2" {
                $serviceTag = GetDellServiceTag($script:foundADObject.Name)
                $machineData = @()
                $machineData += GetMachineType($serviceTag)
                $machineData += "Dell Service Tag: $($serviceTag)"
                WriteListInSidePanel($machineData)
            }
            "3" {WriteListInSidePanel($(GetInstalledSoftware($script:foundADObject.Name)))}
            "4" {
                $query = Read-Host -Prompt "Query"
                WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {($_.Key -like "*$($query)*") -or ($_.Value -like "*$($query)*")})
            }
        }
        ""

        Pause
    }
}

#Start Script Here
SetColourTheme
for(;;){
    MainMenu
    $host.ui.RawUI.WindowTitle = "Active Directory Utilities"
}
