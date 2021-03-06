#############################################################################
# Calendar permission handling by greg.kustra@carlsonwagonlit.com ®
# ver. 1.0 - all loops and functionality implemented - tested 
# ver. 1.1 - added polish & german lang in switch statement
# ver. 1.2 - added Remove permission option
# ver. 1.3 - some error handling added
# ver. 1.4 - added french, portuguese, spanish & italian lang in switch statement
#############################################################################
$ver = "1.4"
############# Main loop ########################
$continue = "Y"
do {
       
    #################### Select the language ########################
    Write-Host "Calendar Access Program ver. $ver" -ForegroundColor Green 
    Write-Host "Press CTRL+C to exit the script at any time" -ForegroundColor red -BackgroundColor Blue
    Write-Host "Please select the language" -ForegroundColor Yellow
    [int]$xMenuLanguage = 0
    while ($xMenuLanguage -lt 1 -or $xMenuLanguage -gt 7){
        Write-host "1. English for calendar"
        Write-host "2. German for kalender"
        Write-host "3. Polish for kalendarz"
        Write-host "4. French for calendrier"
        Write-host "5. Portuguese for calendario"
        Write-host "6. Spanish for calendario"
        Write-host "7. Italian for calendario"
             
        [Int]$xMenuLanguage = read-host "Please enter an option 1 to 7..." 
    }
    #################### set object type based on language chosen ###################
    [string]$cal = ""  
    
    switch ($xMenuLanguage) { 
        1 {$cal = "calendar"} 
        2 {$cal = "kalender"} 
        3 {$cal = "kalendarz"} 
        4 {$cal = "calendrier"} 
        5 {$cal = "calendário"}
        6 {$cal = "calendario"}
        7 {$cal = "calendario"}
        default {$cal = "calendario"}
    }
         
    ############### Get user id / Alias of Shared MailBox ###########################
    [string]$owner = ""
    while ($owner -eq "") {
       $owner = Read-Host "Please enter User id (or an Alias of Shared MailBox) of Calendar Owner"
    } 
    #Get-MailboxRegionalConfiguration -Identity $owner
       
    ############### Display permission for calendar so you know if Get or Set should be used ###########################
    Write-Output "Wait ................"
    $FolderPermission = Get-MailboxFolderPermission -identity "${owner}:\$cal" #-erroraction silentlycontinue
    Write-Output $FolderPermission 
    
    if ($FolderPermission -eq $null) {
        Write-host "....... Validating ........"
        Write-host "Wrong calendar language setting chosen, please start the script again and choose proper language" -ForegroundColor Magenta
        break # it doesn't make sense to continue
    }
        
    ################ Decision to continue ########################
    $continue = Read-Host "Do you want to continue (Y/N) ?)"
    if ($continue -eq "N") {break}
    
    #######################Set User id ###################################
    [string]$guest = ""
    while ($guest -eq "") {
        $guest = Read-Host "Please enter User ID of Guest"
    }    
    
    ######################## SET, ADD or REMOVE Permission #######################
    Write-Host "Please select Set, Add  or Remove permission" -ForegroundColor Green
    [int]$SetAdd = 0
    while ($SetAdd -lt 1 -or $SetAdd -gt 3){
        Write-host "1. Add permission (no permission set up yet)"
        Write-host "2. Set permission (change existing permision)"
        Write-host "3. Remove permission (delete existing permision)"
               
        [Int]$SetAdd = read-host "Please enter an option 1 to 3..." 
    } 
    
    
    if ($SetAdd -lt 3) {     
        ################### Present Multiple Choice Menu of Roles #############################
        $Roles = @("None","Owner","PublishingEditor","Editor","PublishingAuthor","Author","NonEditingAuthor", "Reviewer", "Contributor")
        For ($i=0; $i -lt $Roles.Length; $i++) {
            $i.toString() + ". " + $Roles[$i].toString()
        }
    
        ################### Select the role ########################
        Write-Host "Please select the rights" -ForegroundColor Magenta
        [int]$xMenuChoiceA = 0
        $max = $Roles.Length - 1 
        while ($xMenuChoiceA -lt 1 -or $xMenuChoiceA -gt $max){
            [Int]$xMenuChoiceA = read-host "Please enter an option 1 to" $max
        }
    }
       
    Write-Output "Wait ................"
    
    $PermissionLevel = $Roles[$xMenuChoiceA].toString()
    #Write-Output $PermissionLevel
      
    if ($SetAdd -eq 1) {
        Add-MailboxFolderPermission -identity "${owner}:\$cal"  –user "$guest” -AccessRights "$PermissionLevel"   
    }
    elseif ($SetAdd -eq 2) {
        Set-MailboxFolderPermission -identity "${owner}:\$cal"  –user "$guest” -AccessRights "$PermissionLevel" 
    }
    else {
        Remove-MailboxFolderPermission -identity “${owner}:\$cal” –user “$guest”
    }
    Write-Host "....... Review Current Permission ..........." -ForegroundColor Cyan
    
    ################show the access level#######################
    Get-MailboxFolderPermission -identity "${owner}:\$cal"
    
############ Condition of main loop ##################
$continue = Read-Host "Type in Y to continue with another user/calendar or any char to break"
    
} while ($continue -eq "Y")