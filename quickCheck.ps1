function service{
    $D = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $PDC = $D.PdcRoleOwner 
 
    # Setup the DirectorySearcher object. 
    $Searcher = New-Object System.DirectoryServices.DirectorySearcher 
    $Searcher.PageSize = 200 
    $Searcher.SearchScope = "subtree" 
    $Searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
    # Specify attribute values to retrieve. 
    $Searcher.PropertiesToLoad.Add("distinguishedName") |Out-Null 
    $Searcher.PropertiesToLoad.Add("modifyTimeStamp") |Out-Null 
    $Searcher.PropertiesToLoad.Add("sAMAccountName") |Out-Null 
    $Searcher.PropertiesToLoad.Add("lastLogonTimeStamp") |Out-Null 

    $HashTable = @{ }

    $DC = $(Get-ADDomain $D.Name).distinguishedName    
    $Base = "LDAP://$D"
    $Searcher.SearchRoot = $Base 
    $Results = $Searcher.FindAll()
    $count =0
    If($Results) 
    { 
        # Output one line for each account. 
        ForEach ($Result In $Results) 
        { 
    # Retrieve the values. 
            $DN = $Result.Properties.Item("distinguishedName")[0]  
            $sam = $Result.Properties.Item("sAMAccountName")[0]  
            $logon = $Result.Properties.Item("lastLogonTimeStamp")[0]   
 
            if(($sam -match "^[pP]98[5..7]") -and ($sam -notmatch "^[pP]981")){ 
                
                $lastLogon = [datetime]::fromfiletime($logon)
          
                $currentDate = get-date 
                if($lastLogon -lt $currentDate.AddDays(-105)){
                "$sam,$lastLogon"
                 $count++
                }
            }
               
        }
           
    } 
        Else 
    { 
         Write-Host "ERROR: Failed to connect to DC $Server" -foregroundcolor red -backgroundcolor black 
        "<DC not found>" 
    } 
    write-host "-----------------------------"       
    write-host $count -ForegroundColor Green
}
service
function users{
$D = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$PDC = $D.PdcRoleOwner 
 
# Setup the DirectorySearcher object. 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.PageSize = 200 
$Searcher.SearchScope = "subtree" 
$Searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
# Specify attribute values to retrieve. 
$Searcher.PropertiesToLoad.Add("distinguishedName") |Out-Null 
$Searcher.PropertiesToLoad.Add("modifyTimeStamp") |Out-Null 
$Searcher.PropertiesToLoad.Add("sAMAccountName") |Out-Null 
$Searcher.PropertiesToLoad.Add("lastLogonTimeStamp") |Out-Null 

$HashTable = @{ }

$DC = $(Get-ADDomain $D.Name).distinguishedName    
$Base = "LDAP://$D/CN=Users,$DC"
$Searcher.SearchRoot = $Base 
$Results = $Searcher.FindAll()
$count =0
If($Results) 
{ 
    # Output one line for each account. 
    ForEach ($Result In $Results) 
    { 
# Retrieve the values. 
        $DN = $Result.Properties.Item("distinguishedName")[0]  
        $sam = $Result.Properties.Item("sAMAccountName")[0]  
        $logon = $Result.Properties.Item("lastLogonTimeStamp")[0]   
 
        if(($sam -notmatch "^[pP]98[5..7]")){ 
                
            $lastLogon = [datetime]::fromfiletime($logon)
          
            $currentDate = get-date 
            if($lastLogon -lt $currentDate.AddDays(-105)){
            "$sam,$lastLogon"
             $count++
            }
        }
               
    }
           
} 
    Else 
{ 
     Write-Host "ERROR: Failed to connect to DC $Server" -foregroundcolor red -backgroundcolor black 
    "<DC not found>" 
} 
write-host "-----------------------------"       
write-host $count -ForegroundColor Green
}
users
function admin{
$D = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$PDC = $D.PdcRoleOwner 
 
# Setup the DirectorySearcher object. 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.PageSize = 200 
$Searcher.SearchScope = "subtree" 
$Searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
# Specify attribute values to retrieve. 
$Searcher.PropertiesToLoad.Add("distinguishedName") |Out-Null 
$Searcher.PropertiesToLoad.Add("modifyTimeStamp") |Out-Null 
$Searcher.PropertiesToLoad.Add("sAMAccountName") |Out-Null 
$Searcher.PropertiesToLoad.Add("lastLogonTimeStamp") |Out-Null 

$HashTable = @{ }

$DC = $(Get-ADDomain $D.Name).distinguishedName    
$Base = "LDAP://$D/OU=Admin Accounts,OU=Admin Roles,$DC"
$Searcher.SearchRoot = $Base 
$Results = $Searcher.FindAll()
$count =0
If($Results) 
{ 
    # Output one line for each account. 
    ForEach ($Result In $Results) 
    { 
# Retrieve the values. 
        $DN = $Result.Properties.Item("distinguishedName")[0]  
        $sam = $Result.Properties.Item("sAMAccountName")[0]     
        $logon = $Result.Properties.Item("lastLogonTimeStamp")[0]      
        $lastLogon = [datetime]::fromfiletime($logon)
        $currentDate = get-date 
        if($lastLogon -lt $currentDate.AddDays(-105)){
        "$sam,$lastLogon"
         $count++
        }              
    }     
} 
    Else 
{ 
     Write-Host "ERROR: Failed to connect to DC $Server" -foregroundcolor red -backgroundcolor black 
    "<DC not found>" 
} 
write-host "-----------------------------"       
write-host $count -ForegroundColor Green
}
write-host
admin
