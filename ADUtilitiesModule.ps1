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

    return "$($wmiObject.GetDWORDValue(2147483650, "SOFTWARE\WOW6432Node\TeamViewer", "ClientID").uValue)"
}

function GetDellServiceTag {param($computerName)
    if(!$script:alternateCreds)
    {
        return (Get-WmiObject -Class win32_bios -ComputerName $computerName).SerialNumber
    }

    return (Get-WmiObject -Class win32_bios -ComputerName $computerName -Credential $script:alternateCreds).SerialNumber
}

function GetMachineType {param($serviceTag)
    $htmlText = [System.Net.WebClient]::new().DownloadString("https://www.dell.com/support/home/us/en/04/product-support/servicetag/$($serviceTag)/configuration")
    $html = New-Object -Com "HTMLFile"
    $html.IHTMLDocument2_write($htmlText)
    return (($html.all.tags("h1") | % InnerText) | Out-String).Trim()
}

function GetInstalledSoftware {param($computerName)
    $softwareList = @()
    if(!$script:alternateCreds)
    {
        $wmiObject = Get-WmiObject -List StdRegProv -ComputerName $computerName
    } else {
        $wmiObject = Get-WmiObject -List StdRegProv -ComputerName $computerName -Credential $script:alternateCreds
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

function GetLastUser {param($computerName)
    $wmiObject
    if(!$script:alternateCreds)
    {
        $wmiObject = Get-WmiObject -Class win32_loggedonuser -ComputerName $computerName
    } else {
        $wmiObject = Get-WmiObject -Class win32_loggedonuser -ComputerName $computerName -Credential $alternateCreds
    }
    # | Where-Object {$_.Antecedent -like "*DOMAIN*"} 
    return ($wmiObject| Select-Object Antecedent)
}


function RestartMachine {param($computerName)
    $response = Read-Host -Prompt "Are you sure? (Y/N)"
    if($response.ToLower() -ne "y"){
        return
    }

    Restart-Computer -ComputerName $computerName
}
