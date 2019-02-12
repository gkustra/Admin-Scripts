##################################################################################
# Shared Mailbox permission adding by greg.kustra@carlsonwagonlit.com ®
# ver. 1.0 - All loops and main functionality implemented - tested
# ver. 1.1 - Logic implemented to add users list separated by comma into AD Group 
# ver. 1.2 - Bug fixes
##################################################################################
$ver = "1.2"
############# Main loop ########################
$continue = "Y"
do {
    
    #################### Script info ########################
    Write-Host "SMB Granting Permission Script ver. $ver" -ForegroundColor Green
    Write-Host "Press CTRL+C to exit the script at any time" -ForegroundColor yellow -BackgroundColor Blue
    
    #################### Get SMB Displayed name ########################
    $groupArray = @()
    [string]$smbName = ""
    while ($smbName -eq "") {
        $smbName = Read-Host "Please input Shared MailBox name (Displayed Name)"
    }
    $groups = Get-MailboxPermission -Identity $smbName #PfizerEventsCWTUK
    #################### Display all CWT groups and CWT users with access ########################
    Write-Host "Below users & groups have the access currently" -ForegroundColor White
    foreach($group in $groups) {   
        if ($group.User.Contains("CWT\")) {  
            Write-host $group.User -ForegroundColor Yellow
            #if ($group.User.Contains(".")) {
            #    $groupArray += $group.User         
            #}
            $cwtUser = $group.User -notmatch "CWT\\U[0-9A-Z]{6}"
            if ($cwtUser) {
                $groupArray += $group.User            
            }
         }
    }
    Write-Host "Wait .....You can add the user(s) to the group - chose the number......" -ForegroundColor Magenta
    if ($groupArray.Length -gt 0) {
        [int]$j = 0
        for ($i=0; $i -lt $groupArray.Length; $i++ ){
            $j = $i + 1
            Write-host $j.ToString() $groupArray[$i]
        }
    #################### Menu to display CWT Groups if exists ########################
        [Int]$d = 0
        while ($d -lt 1 -or $d -gt $groupArray.Length){
            $d = read-host "Please enter an option from 1 to" $groupArray.Length "to add user id to the group or hit enter to skip the adding"
            if ($d -eq "") {break}
        }
    }
    #################### Get User ID & action it if not blank ########################
    [String]$UserID = ""
    while ($UserID -eq "") {
        $UserID = Read-Host "Please enter User ID" 
    }
    #################### Condition to check if more user ids input ########################
    $UserID = $UserID.Replace(" ","")
    $listMatch = $UserID -match "[U-u][0-9A-Za-z]{6},"
    if ($listMatch) {
        $UserList = $UserID.split(",")
    }  
    if ($UserID -ne "") {
        Write-Host "Adding $UserID to ....." -ForegroundColor Yellow
        if ($d -gt 0 -and $d -le $groupArray.Length) {
            [String]$CWTGroup = ""
            $CWTGroup = $groupArray[$d-1].Replace("CWT\","")     
            Write-Host $groupArray[$d-1] -ForegroundColor Green
            if ($listMatch) {
                Add-ADGroupMember -Identity $CWTGroup -Members $UserList
            }
            else{
                Add-ADGroupMember -Identity $CWTGroup -Members $UserID
            }            
        }
        else {
            Write-Host $smbName -ForegroundColor Green 
            if ($listMatch) {
                Write-Host "Only first user from the list will be handled - $UserList[1]" -ForegroundColor Red
                Add-MailboxPermission -Identity $smbName -User $UserList[1] -AccessRights FullAccess
            }
            else {
                Add-MailboxPermission -Identity $smbName -User $UserID -AccessRights FullAccess
            }
            #Check if Send-As access is required 
            $Yes = Read-Host "Type in Y to grant also Send-As Access for this user"
            if ($Yes = "Y") {
                if ($listMatch) {
                    Write-Host "Only first user from the list will be handled - $UserList[1]" -ForegroundColor Red
                    Add-ADPermission $smbName -User $UserList[1] -ExtendedRights Send-As    
                }
                else {
                    Add-ADPermission $smbName -User $UserID -ExtendedRights Send-As                
                }                
            }    
        }         
    } 
############ Condition of main loop ##################
$continue = Read-Host "Type in Y to continue with another SMB or any char to break"    
} while ($continue -eq "Y")

    
