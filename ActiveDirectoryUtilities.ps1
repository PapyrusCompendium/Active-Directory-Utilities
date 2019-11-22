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

$script:foundADUser

function HelpDisplay {
    "   ┌────────────────────────┐"
    "   │----Search Functions----│"
    "   │By Name:               0│"
    "   │By Username:           1│"
    "   │By Employee ID:        2│"
    "   ├────────────────────────┤"
    "   │< Exit Script          #│"
    "   └────────────────────────┘"
}

function DisplayData{
    "  ┌────────────────────────┐"
    "  │------Display Data------│"
    "  │Employee ID:           0│"
    "  │Phones:                1│"
    "  │Email:                 2│"
    "  │Full Name:             3│"
    "  │Account Creation:      4│"
    "  │Password Details:      5│"
    "  │Custom Query:          6│"
    "  │All Data:              *│"
    "  ├────────────────────────┤"
    "  │-----Account Configs----│"
    "  │Change Employee ID:    7│"
    "  │Change Password        8│"
    "  ├────────────────────────┤"
    "  │< Go Back              #│"
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
    $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates $x, $y
}

function WriteInSidePanel{param($textData)
    $lineNumber = 17
    $width = 35

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

function ChnageEmployeeID{
    $employeeID = Read-Host -Prompt "New Employee ID"
    Set-ADUser $script:foundADUser -EmployeeID  $employeeID
}

function ChangeUserPassword{
    $newPassword = Read-Host -Prompt "New Password" -AsSecureString
    Set-ADAccountPassword $script:foundADUser -Reset -NewPassword $newPassword
}

function SelectADUser { param ($adUsers)
    ClearBanner

    if($adUsers.GetType().Name -ne "ADUser"){
        for($x = 0; $x -lt $adUsers.Count; $x++){
            "   $($adUsers[$x].Name)[$($x)]"
        }
        "   Go Back[-1]"
    
        [int16]$index = Read-Host -Prompt "   Option"
        if($index -eq -1){
            return
        }
    
        $script:foundADUser =  $adUsers[$index]
    } else {
        $script:foundADUser = $adUsers
    }

    DataMenu
}

function SearchByName{
    ClearBanner
    $name = "*$(Read-Host -Prompt "Name")*".Replace(" ", "*")

    if(Get-ADUser -F {Name -like $name}){
        SelectADUser(Get-ADUser -F {Name -like $name} -Properties *)
    } else {
        "No employees found!"
        Pause
    }
}

function SearchByUsername{
    ClearBanner
    $username = "*$(Read-Host -Prompt "Username")*"

    if(Get-ADUser -F {SamAccountName -like $username}){
        SelectADUser(Get-ADUser -F {SamAccountName -like $username} -Properties *)
    } else {
        "Employee not found!"
        Pause
    }
}

function SearchByEmployeeID{
    ClearBanner
    $employeeID = Read-Host -Prompt "Employee ID"

    if(Get-ADUser -F {EmployeeID -eq $employeeID}){
        SelectADUser(Get-ADUser -F {EmployeeID -eq $employeeID} -Properties *)
    } else {
        "Employee not found!"
        Pause
    }
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
        "#" {Exit}
    }
}

function DataMenu{
    for(;;){
        ClearBanner
        "   $($script:foundADUser.Name) - $($script:foundADUser.SamAccountName)"
        $host.ui.RawUI.WindowTitle = "Active Directory Utilities - $($script:foundADUser.Name)"
        DisplayData

        $optionSelection = Read-Host -Prompt "   Option"

        ""
        switch($optionSelection)
        {
            "#" {return}
            "*" {WriteInSidePanel($script:foundADUser.GetEnumerator() | ? {$_.Key -like "*"})}
            "0" {WriteInSidePanel($script:foundADUser.GetEnumerator() | ? {$_.Key -eq "EmployeeID"})}
            "1" {WriteInSidePanel($script:foundADUser.GetEnumerator() | ? {$_.Key -like "*phone*"})}
            "2" {WriteInSidePanel($script:foundADUser.GetEnumerator() | ? {$_.Key -like "*email*"})}
            "3" {WriteInSidePanel($script:foundADUser.GetEnumerator() | ? {$_.Key -eq "Name"})}
            "4" {WriteInSidePanel($script:foundADUser.GetEnumerator() | ? {($_.Key -like "*account*") -or $_.Key -like "*created*"})}
            "5" {WriteInSidePanel($script:foundADUser.GetEnumerator() | ? {$_.Key -like "*password*"})}
            "6" {
                Write-Host -NoNewline "   Query: "
                $query = Read-Host
                WriteInSidePanel($script:foundADUser.GetEnumerator() | ? {($_.Key -like "*$($query)*") -or ($_.Value -like "*$($query)*")})
            }
            "7" {ChnageEmployeeID}
            "8" {ChangeUserPassword}
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
