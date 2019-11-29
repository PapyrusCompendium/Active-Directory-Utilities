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
    "  │IPv4 Address:          0│       │"
    "  │Team Viewer  ID:       1│       │"
    "  │Dell Service Tag       2│       │"
    "  │Installed Software:    3│       │"
    "  │Custom Query:          4│       │"
    "  │All Data:              *│       │"
    "  ├────────────────────────┤       │"
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
    
        $script:foundADObject =  $adUsers[$index]
    } else {
        $script:foundADObject = $adUsers
    }

    UserDataMenu
}

function SelectADComputer { param ($adComputers)
    ClearBanner
    if($adComputers.GetType().Name -ne "ADComputer"){
        for($x = 0; $x -lt $adComputers.Count; $x++){
            "   $($adComputers[$x].Name)[$($x)]"
        }
        "   Go Back[-1]"
    
        [int16]$index = Read-Host -Prompt "   Option"
        if($index -eq -1){
            return
        }
    
        $script:foundADObject =  $adComputers[$index]
    } else {
        $script:foundADObject = $adComputers
    }

    ComputerDataMenu
}

function ChangeEmployeeID{
    $employeeID = Read-Host -Prompt "New Employee ID"
    Set-ADUser $script:foundADObject -EmployeeID  $employeeID
}

function ChangeUserPassword{
    $newPassword = Read-Host -Prompt "New Password" -AsSecureString
    Set-ADAccountPassword $script:foundADObject -Reset -NewPassword $newPassword
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

function SearchByComputerName {
    ClearBanner
    $computerName = "*$(Read-Host -Prompt "Computer Name")*".Replace(" ", "*")
    
    if(Get-ADComputer -F {Name -like $computerName}){
        SelectADComputer(Get-ADComputer -F {Name -like $computerName} -Properties *)
    } else {
        "Computer not found!"
        Pause
    }
}

function SearchByIP {
    ClearBanner
    $deviceIP = Read-Host -Prompt "Computer IP"
    
    if(Get-ADComputer -F {IPv4Address -eq $deviceIP} -Properties *){
        SelectADComputer(Get-ADComputer -F {IPv4Address -eq $deviceIP} -Properties *)
    } else {
        "Computer not found!"
        Pause
    }
}

function AlternateCredential {
    ClearBanner
    $username = Read-Host -Prompt "Username"
    $password = Read-Host -Prompt "Password" -AsSecureString
    $script:alternateCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
}

function GetTeamViewerID {param($computerName)
    $wmiObject
    if(!$script:alternateCreds)
    {
        $wmiObject = Get-WmiObject -List StdRegProv -ComputerName $computerName
    } else {
        $wmiObject = Get-WmiObject -List StdRegProv -ComputerName $computerName -Credential $script:alternateCreds
    }

    return $wmiObject.GetDWORDValue(2147483650, "SOFTWARE\WOW6432Node\TeamViewer", "ClientID").uValue
}

function GetDellServiceTag {param($computerName)
    if(!$script:alternateCreds)
    {
        return (Get-WmiObject -Class win32_bios -ComputerName $computerName).SerialNumber
    }

    return (Get-WmiObject -Class win32_bios -ComputerName $computerName -Credential $script:alternateCreds).SerialNumber
}

function GetInstalledSoftware {param($computerName)
    $softwareList = @()
    if(!$script:alternateCreds)
    {
        $wmiObject = Get-WmiObject -List StdRegProv -ComputerName $computerName -Credential $adCreds
    } else {
        $wmiObject = Get-WmiObject -List StdRegProv -ComputerName $computerName
    }

    foreach ($item in $wmiObject.EnumKey(2147483650, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall").sNames |
    ? {($wmiObject.GetDWORDValue(2147483650, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($_)", "SystemComponent").uValue -ne 1)}) {
        $softwareList += $wmiObject.GetStringValue(2147483650, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($item)", "DisplayName").sValue
    }

    foreach ($item in $wmiObject.EnumKey(2147483650, "SOFTWARE\Wow6432node\Microsoft\Windows\CurrentVersion\Uninstall").sNames |
    ? {($wmiObject.GetDWORDValue(2147483650, "SOFTWARE\Wow6432node\Microsoft\Windows\CurrentVersion\Uninstall\$($_)", "SystemComponent").uValue -ne 1)}) {
        $softwareList += $wmiObject.GetStringValue(2147483650, "SOFTWARE\Wow6432node\Microsoft\Windows\CurrentVersion\Uninstall\$($item)", "DisplayName").sValue
    }

    return $softwareList | Where-Object {$_ -notlike "*Update for*"} | Sort-Object
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
            "*" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -like "*"})}
            "0" {WriteInKeyPairSidePanel($script:foundADObject.GetEnumerator() | Where-Object {$_.Key -eq "IPv4Address"})}
            "1" {WriteTextInSidePanel("TeamViewer ID: $(GetTeamViewerID($script:foundADObject.Name))")}
            "2" {WriteTextInSidePanel("Dell Service Tag: $(GetDellServiceTag($script:foundADObject.Name))")}
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
