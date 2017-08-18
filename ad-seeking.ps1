#Dev by Ender Loc Phan
<#Requirements:

Import-Module ActiveDirectory
#>
<#Usage 1.0
 - Suppy the objectClass (Eg: user, group, person...)

 - Just Enumerate Distinguished name

 .\adTracking.ps1 -dna           # Enumerate Distinguished name and print it to console
 .\adTracking.ps1 -dna -addToReport   # Write Distinguished name to text file
 .\adTracking.ps1 -dna -addToReport -amount 100   # Write Distinguished name to text file with specific amout of data
 .\adTracking.ps1 -dna -amount 100       # Print given amount of Distinguished name to console

 - Get All attributes

 .\adTracking.ps1               # Enumerate  all supplied LDAP Attributes and print it to console
 .\adTracking.ps1 -addToReport  # Write all data to CSV file
 .\adTracking.ps1 -addToReport -amount 100 # Write data to CSV file with given amount of data
 .\adTracking.ps1 -amount 100       # Print given amount of data to console   
#>    

<# Update 1.1
- Added the trusted domain method
- Fixed Account expires function
- Fixed PasswordLS
- change parameters to optional methods

Usage 1.1: Just flow the options given by the tool
#>

$activeMo = Import-Module ActiveDirectory -ErrorAction Stop
Write-Verbose -Message  "This tool is running under PowerShell version $($PSVersionTable.PSVersion.Major)" -Verbose
write-host 
write-host " 1. Run on current domain "
write-host " 2. Run on trusted domains "
write-host 
$type =  Read-Host -Prompt "Option "
if ($type -eq 1) 
{
  # Get the Current Domain data  
  $Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
}
elseif($type -eq 2) 
{
    write-host
    write-host " 1. Enter trusted domain manually "
    write-host " 2. Get all trusted domain automatically"
    write-host
    $trust = Read-Host -Prompt "Option "
    if($trust -eq 1){
        
        $trustDN = Read-Host -Prompt "Domain "
        write-host
        $TrustedDomain = $trustDN
    }
    elseif($trust -eq 2){
    
        $trustedD = Get-ADTrust -Filter * | select Name | Out-String
        $trustedD             
        $trustDN = Read-Host -Prompt "Domain "
        write-host
        $TrustedDomain = $trustDN            
    }
    else{
        Write-Verbose -Message  "Unknown entered option" -Verbose
        exit 
    }

    $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$TrustedDomain)
    Try 
    {
        $Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
        Write-Verbose -Message "Connect to $Domain successfully" -Verbose
    }
        Catch [exception] {
        $Host.UI.WriteErrorLine("ERROR: $($_.Exception.Message)")
        Exit
    }
}
else
{
    Write-Verbose -Message  "Option is not valid" -Verbose
    exit
}
$objectCategory =  Read-Host -Prompt "objectCategory "
if($objectCategory -eq ""){
    Write-Verbose -Message  "objectCategory can't be null" -Verbose
    exit    
}
$objectClass =  Read-Host -Prompt "objectClass "
if($objectClass -eq ""){

    Write-Verbose -Message  "Objectclass can't be null" -Verbose
    exit    
}
$PDC = $Domain.PdcRoleOwner
$ADSearch = New-Object System.DirectoryServices.DirectorySearcher
#new empty ad search, search engine someth we can send queries to find out
$ADSearch.SearchRoot ="LDAP://$PDC"
#where we wanna look in LDAP is Domain, because we don't wanna search from root
#root is: $objDomain = New-Object System.DirectoryServices.DirectoryEntry
$ADSearch.SearchScope = "subtree"

$ADSearch.PageSize = 100
$ADSearch.Filter = "(&(objectCategory=$objectCategory)(objectClass=$objectClass))"
#where objectClass attribute are -eq to user
#Atribute to search for: ObjectClass
# value of attribute : user
#exp: $ADSearch.Filter = "(Name=Ender)"
$connect = [ADSI] "LDAP://$($Domain)" 
$lockoutDuration = $connect.lockoutDuration.Value
$lockoutThreshold  =$connect.lockoutThreshold
$maxPwdAge =$connect.maxPwdAge.Value
$maxPwdAgeValue =  $connect.ConvertLargeIntegerToInt64($maxPwdAge)
$duraValue = $connect.ConvertLargeIntegerToInt64($lockoutDuration)
#values in array are atttibutes of LDAP
$properies =@("distinguishedName",
"sAMAccountName",
"mail",
"lastLogonTimeStamp",
"pwdLastSet",
"badpwdcount",
"accountExpires",
"userAccountControl",
"modifyTimeStamp",
"lockoutTime"
"badPasswordTime",
"maxPwdAge ",
"Description"
)
foreach($pro in $properies)
{
    $ADSearch.PropertiesToLoad.add($pro)| out-null
    #the name of property of the object, search will load the name in an array #properties
}
$ProgressBar = $True
$userObjects = $ADSearch.FindAll()
$dnarr = New-Object System.Collections.ArrayList
$modiValues = New-object System.Collections.ArrayList
$Delimiter = ","
$userCount =  $userObjects.Count
$result = @()
$count = 0
# Creating csv file
$invalidChars = [io.path]::GetInvalidFileNameChars()
$dateTimeFile = ((Get-Date -Format s).ToString() -replace "[$invalidChars]","-")
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$outFileService = $($PSScriptRoot)+"\$($Domain)-Report Service Accounts-$($dateTimeFile).csv"
$outFileAdmin = $($PSScriptRoot)+"\$($Domain)-Report Admin Accounts-$($dateTimeFile).csv"
$outFileUser= $($PSScriptRoot)+"\$($Domain)-Report User Accounts-$($dateTimeFile).csv"
$outFileTxt = $($PSScriptRoot)+"\Report-$($dateTimeFile).txt"
$outFileHTMLService = $($PSScriptRoot)+"\$($Domain)-Report Service Accounts-$($dateTimeFile).html"
$outFileMeg = $($PSScriptRoot)+"\$($Domain)-FinalReport-$($dateTimeFile).csv"
$outFileModi = $($PSScriptRoot)+"\$($Domain)-ReportModi-$($dateTimeFile).csv"


$NeverExpires = 9223372036854775807

# Supplied Attributes
$global:exportedToCSV  = $false
$global:exportedToTxt = $false
$global:servicAcc = $false
$global:ea = 0
$global:last2015 = 0
$global:last2016 = 0
$global:last2017 = 0
$global:otherLast = 0
$global:NeverLogon = 0
$global:noLastSet = 0
$global:passSet2015 = 0
$global:passSet2016 = 0
$global:passSet2017 = 0
$global:otherPassSet = 0
$global:noBadSet= 0
$global:basPassC0= 0
$global:basPassC1= 0
$global:basPassC2= 0
$global:basPassC3= 0
$global:noBadLogSet = 0
$global:uknownBadLog = 0
$global:badlog2015 =0
$global:badlog2016 =0
$global:badlog2017 =0
$global:otherBadlog =0
$global:accNotEx = 0
$global:accEx = 0
$global:accDisStatus=0
$global:smartRe =0
$global:passNotRe= 0
$global:passChangeNotAll = 0
$global:passNExpSet = 0
$global:ageNA = 0
$global:ageDate2017=0
$global:ageDate2016=0
$global:ageDate2015=0
$global:otherAgeDAte=0
$global:modi2015 =0
$global:modi2016=0
$global:modi2017=0
$global:otherModi=0
$global:noneModi=0
function ini{
    $global:dn =  $user.Properties.Item("distinguishedName")[0]    
    $global:sam = $user.Properties.Item("sAMAccountName")[0]
    $global:logon = $user.Properties.Item("lastLogonTimeStamp")[0]
    $global:mail =$user.Properties.Item("mail")[0]
    $global:passwordLS = $user.Properties.Item("pwdLastSet")[0]
    $global:passwordC = $user.Properties.Item("badpwdcount")[0]
    $global:accountEx = $user.Properties.Item("accountExpires")[0]
    $global:accountDis= $user.Properties.Item("userAccountControl")[0] 
    $global:modify= $user.Properties.Item("modifyTimeStamp")[0]
    $global:lockoutTime= $user.Properties.Item("lockoutTime")[0]
	$global:lastFailedAt = $user.Properties.item("badPasswordTime")[0]
    $global:Description = $user.Properties.item("Description")[0] 
    $global:passSet = $false
    $global:passTrue =$false
}


Function tracking
{
    param($fileName)
        #Last Logon
        $global:lastLogon = [datetime]::fromfiletime($global:logon)

        $currentDate = get-date 
        if($global:lastLogon -lt $currentDate.AddDays(-105)){
            <#
            if ($logon -eq $null) {
                Write-Host "Logon is `$null"
            }
            if ($lastLogon -eq $null) {
                Write-Host "Date is `$null"
            }
            $lastLogon.Year#>        
            $global:lastLogon= $global:lastLogon.ToString("yyyy/MM/dd")           
            if($global:lastLogon.split("/")[0] -eq 2015){
                $global:last2015++
            }     
            elseif ($global:lastLogon.split("/")[0] -eq 2016){
                $global:last2016++
            }
            elseif ($global:lastLogon.split("/")[0] -eq 2017){
                $global:last2017++
            }elseif ($global:lastLogon.split("/")[0] -eq 1601){
                $global:lastLogon = "Never"
                $global:NeverLogon++
            }else{
                $global:otherLast++
            }
               
            #password last set
            if($global:passwordLS -eq 0)
            {         
                 $global:value = "Never"
                 $global:noLastSet++
            }
            else
            {         
                 $global:value = [datetime]::fromfiletime($global:passwordLS)                   
                 $global:value = $global:value.ToString("yyyy/MM/dd")
                 if($global:value.split("/")[0] -eq 2015){
                     $global:passSet = $true
                     $global:passSet2015++
                 }     
                 elseif ($global:value.split("/")[0] -eq 2016){
                     $global:passSet = $true
                     $global:passSet2016++
                 }
                 elseif ($global:value.split("/")[0] -eq 2017){
                     $global:passSet = $true
                     $global:passSet2017++
                 }
                 elseif ($global:value.split("/")[0] -eq 1601){
                     $global:value = "Never"   
                     $global:noLastSet++ 
                 }
                 else{
                     $global:passSet = $true
                     $global:otherPassSet++
                 }
         
            }     
            #Account expires   
            if(($global:accountEx -eq $NeverExpires) -or ($global:accountEx -gt [Datetime]::MaxValue.Ticks))
            {
                $global:convertAccountEx = "Not Expired"
        
            }
            else
            {
                #$convertDate = [datetime]$accountEx
                $global:convertAccountEx = "Expired"
                $global:accEx++
            }
            #Email
            if([String]::IsNullOrEmpty($global:mail)){
        
                $global:email = "N/A"
        
            }
            else{
                $global:email =$global:mail
                $global:ea++
            }
            #PasswordCount
            if([String]::IsNullOrEmpty($global:passwordC)){

                $global:passwordCStatus = "N/A"
                $global:noBadSet++
            }
            else{

                $global:passwordCStatus = $global:passwordC   
                if($global:passwordC -eq 0){
                    $global:basPassC0++
                }       
                elseif($global:passwordC -eq 1){
                    $global:basPassC1++
                }
                elseif($global:passwordC -eq 2){
                    $global:basPassC2++
                }
                else{
                    $global:basPassC3++
                }
            }  
            #UserInfor
            if($global:accountDis -band 0x0002)
            {
                $global:accountDisStatus = "disabled"
                $global:accDisStatus++
            }
            else
            {
                $global:accountDisStatus = "none-disabled"
            }  
            #If Smartcard Required
            if( $global:accountDis -band 262144)
            {
                $global:smartCDStatus = "Required"
                $global:smartRe++
            }
            else
            {
                $global:smartCDStatus = "Not Required"
            }  

            #If No password is required
            if( $global:accountDis -band 32){
                $global:passwordEnforced ="Not Required"
                $global:passNotRe++
            }
            else
            {
                $global:passwordEnforced = "Required"
            }  

            #Password never expired
            if( $global:accountDis -band 0x10000){
                $global:passNExp ="Never Expires is set"
                $global:passNExpSet++
        
            }
            else
            {
                $global:passNExp = "None Set"
                $global:passTrue = $true
            }  
    
            #Datetime bad Logon
            if ($global:lastFailedAt -eq 0){
                $global:badLogOnTime = "Unknown"
                $global:uknownBadLog++
	        }
	        else{
                $global:badLogOnTime = [datetime]::fromfiletime($global:lastFailedAt)              
                $global:badLogOnTime= $global:badLogOnTime.ToString("yyyy/MM/dd")
                if($global:badLogOnTime.split("/")[0] -eq 2015){
                    $global:badlog2015++
                }       
                elseif($global:badLogOnTime.split("/")[0] -eq 2016){
                    $global:badlog2016++
                }
                elseif($global:badLogOnTime.split("/")[0] -eq 2017){
                    $global:badlog2017++
                }
                elseif($global:badLogOnTime.split("/")[0] -eq 1601){
                     $global:badLogOnTime = "Never"    
                     $global:noBadLogSet++
                }
                else{
                     $global:otherBadlog++
                }	    
           }   	  
            #maxPwdAgeValue to get expiration date
            $global:expDAte = $global:passwordLS - $maxPwdAgeValue    
            $global:expDAte = [datetime]::fromfiletime($global:expDAte) 
            if(($global:passTrue -eq $true)-and ($global:passSet -eq $true)){
                $global:expDAte = $global:expDAte.ToString("yyyy/MM/dd")
                if($global:expDAte.split("/")[0] -eq 2015){
                    $global:ageDate2015++
                }       
                elseif($global:expDAte.split("/")[0] -eq 2016){
                    $global:ageDate2016++
                }
                elseif($global:expDAte.split("/")[0] -eq 2017){
                    $global:ageDate2017++
                }
                elseif($global:expDAte.split("/")[0] -eq 1601){
                    $global:expDAte = "N/A"
                    $global:ageNA++
                }
                else{
                    $global:otherAgeDAte++
                }       
            }
            else{
                $global:expDAte = "N/A"
                $global:ageNA++
            } 
            #$Modify
            if($global:modify -ne $null){   
                $global:modify = $global:modify.ToString("yyyy/MM/dd")
                if($global:modify.split("/")[0] -eq 2015){
                    $global:modi2015++
                }       
                elseif($global:modify.split("/")[0] -eq 2016){
                    $global:modi2016++
                }
                elseif($global:modify.split("/")[0] -eq 2017){
                     $global:modi2017++
                }
                else{
                     $global:otherModi++
                }
            }
            else{
                $global:modify = "N/A"
                $global:noneModi++
            }

            $obj = New-object -TypeName psobject
            $obj | Add-Member -MemberType NoteProperty -Name "Distinguished Name" -Value $global:dn
            $obj | Add-Member -MemberType NoteProperty -Name "Sam account" -Value $global:sam
            $obj | Add-Member -MemberType NoteProperty -Name "Email" -Value $global:email
            $obj | Add-Member -MemberType NoteProperty -Name "Password last changed" -Value $global:value
            $obj | Add-Member -MemberType NoteProperty -Name "Bad password count" -Value $global:passwordCStatus
            $obj | Add-Member -MemberType NoteProperty -Name "Last Bad Attempt" -Value $global:badLogOnTime 
            $obj | Add-Member -MemberType NoteProperty -Name "Last Logon " -Value $global:lastLogon
            $obj | Add-Member -MemberType NoteProperty -Name "Account Expires" -Value $global:convertAccountEx
            $obj | Add-Member -MemberType NoteProperty -Name "Account Status" -Value $global:accountDisStatus  
            $obj | Add-Member -MemberType NoteProperty -Name "Smartcard Required" -Value $global:smartCDStatus 
            $obj | Add-Member -MemberType NoteProperty -Name "Password Required" -Value $global:passwordEnforced  
            #$obj | Add-Member -MemberType NoteProperty -Name "Password Change" -Value $passChange  
            $obj | Add-Member -MemberType NoteProperty -Name "Never Expired Password Set" -Value $global:passNExp  
            $obj | Add-Member -MemberType NoteProperty -Name "Password Expiration Date" -Value $global:expDAte
            $obj | Add-Member -MemberType NoteProperty -Name "Last Modified" -Value $global:modify    
            $obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $global:Description  
            if($exportCheck -eq $true){    
                    $global:exportedToCSV = $true
                    $obj | Export-Csv -Path "$fileName" -NoTypeInformation -append -Delimiter $Delimiter     
            }
            else
            {
                $obj 
            } 
            
    }
}

#outFile --> HTML

function html{

$global:IncludeImages = New-Object System.Collections.ArrayList
$global:check= 0
$global:outFilePicPie = $($PSScriptRoot)+"\Pie-$($dateTimeFile)-$($global:check).jpeg"
#PIE
    #Email
$emailPer = $global:ea 
#$emailPer= [math]::Round($emailPer,2)
$noEmailPer=  $userCount - $emailPer
$mailHash = @{"Available"=$emailPer;"Unavailable"=$noEmailPer}
    #Account expired
$accExPer = $global:accEx
#$accExPer = [math]::Round($accExPer,2)
$accNotExPer = $userCount - $accExPer
$accExHash = @{"Expired"="$accExPer";"Unexpired"="$accNotExPer"}
    #Account Status
$accDisPer = $global:accDisStatus 
#$accDisPer = [math]::Round($accDisPer,2)
$accNoDisPer = $userCount - $accDisPer
$accStatusHash = @{"Disabled"="$accDisPer";"Enabled"="$accNoDisPer"}
    #Smart Card required
$smartRePer = $global:smartRe
#$smartRePer = [math]::Round($smartRePer,2)
$smartNotRePer = $userCount - $smartRePer
$smartReHash = @{"Required"="$smartRePer";"Not Required"="$smartNotRePer"}
    #Password Required
$passReNotPer = $global:passNotRe 
#$passReNotPer = [Math]::Round($passReNotPer,2)
$passRePer =  $userCount - $passReNotPer
$passReHash = @{"Not Required"="$passReNotPer";"Required"="$passRePer"}
    #Password Changed
$passChangeNotAllPer = $global:passChangeNotAll
#$passChangeNotAllPer = [math]::Round($passChangeNotAllPer,2)
$passChangeAllper =  $userCount - $passChangeNotAllPer
$passChangedHash = @{"Allowed"="$passChangeAllper";"Not Allowed"="$passChangeNotAllPer";}
    #Password Never Expired Set
$passExpSetPer =$global:passNExpSet
#$passExpSetPer = [math]::Round($passExpSetPer)
$passExpNoSetPer= $userCount - $passExpSetPer
$passExpHash = @{"Set"="$passExpSetPer";"None-set"="$passExpNoSetPer"}
Function drawPie {
    param($hash,
    [string]$title
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $Series = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Series
    $ChartTypes = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
    $Series.ChartType = $ChartTypes::Pie
    $Chart.Series.Add($Series)
    $Chart.ChartAreas.Add($ChartArea)
    $Chart.Series['Series1'].Points.DataBindXY($hash.keys, $hash.values)
    $Chart.Series[‘Series1’][‘PieLabelStyle’] = ‘Disabled’
    $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $Legend.IsEquallySpacedItems = $True
    $Legend.BorderColor = 'Black'
    $Chart.Legends.Add($Legend)
    $chart.Series["Series1"].LegendText = "#VALX (#VALY)"
    $Chart.Width = 700
    $Chart.Height = 400
    $Chart.Left = 10
    $Chart.Top = 10
    $Chart.BackColor = [System.Drawing.Color]::White
    $Chart.BorderColor = 'Black'
    $Chart.BorderDashStyle = 'Solid'
    $ChartTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
    $ChartTitle.Text = $title
    $Font = New-Object System.Drawing.Font @('Microsoft Sans Serif','12', [System.Drawing.FontStyle]::Bold)
    $ChartTitle.Font =$Font
    $Chart.Titles.Add($ChartTitle)
    $testPath = Test-Path $global:outFilePicPie
    if($testPath -eq $True){
        $global:check += 1      
        $global:outFilePicPie = $($PSScriptRoot)+"\Pie-$($dateTimeFile)-$($global:check).jpeg"                 
    }
    $global:IncludeImages.Add($global:outFilePicPie)
    $Chart.SaveImage($outFilePicPie, 'jpeg')  
}
#BAR
    #lastLogon
$lastLogonHash = [ordered]@{"Never"="$global:NeverLogon";"<2015"="$global:otherLast";"2015"="$global:last2015";"2016"="$global:last2016";"2017"="$global:last2017"}
$global:check1= 0
$global:outFilePicBar = $($PSScriptRoot)+"\Bar-$($dateTimeFile)-$($global:check).jpeg"
    #PassLastSet
$passSetHash = [ordered]@{"Never"="$global:noLastSet";"<2015"="$global:otherPassSet";"2015"="$global:passSet2015";
                        "2016"="$global:passSet2016";"2017"="$global:passSet2017";}
    #BadPassCount
$badPassCHash = [ordered]@{"N/A"="$global:noBadSet";"0"="$global:basPassC0";"1"="$global:basPassC1";
                            "2"="$global:basPassC2";"3"="$global:basPassC3" }
    #Last bad Attempt
$lastBadLogHash = [ordered]@{"Unknown"="$global:uknownBadLog";"Never"="$global:noBadLogSet";"<2015"="$global:otherBadlog";"2015"="$global:badlog2015";"2016"="$global:badlog2016";"2017"="$global:badlog2017"}
    #password Age   
$ageHash = [ordered]@{"N/A"="$global:ageNA";"<2015"="$global:otherAgeDAte";"2015"="$global:ageDate2015";
                                "2016"="$global:ageDate2016";"2017"="$global:ageDate2017" }
    #Last Modi
    
$lastModihash = [ordered]@{ "N/A"=$global:noneModi++;"<2015"="$global:otherModi";"2015"="$global:modi2015";
                                "2016"="$global:modi2016";"2017"="$global:modi2017"}

function drawBar{
    param(
    $hash,[string]$title
    ) 
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization
    $Chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $ChartArea1 = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $Series1 = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Series
    $ChartTypes1 = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
    #$Series1.ChartType = $ChartTypes1::Bar
    $Chart1.Series.Add($Series1)
    $Chart1.ChartAreas.Add($ChartArea1)
    #$Chart1.Series.Add("dataset") | Out-Null
    $Chart1.Series[‘Series1’].Points.DataBindXY($hash.keys, $hash.values)
    $chart1.Series[0].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column 
    #$Chart1.Series['Series1'].Points.DataBindXY($hash.keys, $hash.values)
    $ChartArea1.AxisX.Title = "Years"
    $ChartArea1.AxisY.Title = "Figures"
    $Chart1.Series[‘Series1’].IsValueShownAsLabel = $True
    $Chart1.Series[‘Series1’].SmartLabelStyle.Enabled = $True
    $chart1.Series[‘Series1’]["LabelStyle"] = "TopLeft"
    #$chart1.Series[0]["PieLabelStyle"] = "Outside" 
    ##$chart1.Series[0]["DrawingStyle"] = "Emboss" 
    #$chart1.Series[0]["PieLineColor"] = "Black" 
    #$chart1.Series[0]["PieDrawingStyle"] = "Concave"

    if($global:amount){
        $ChartArea1.AxisY.Maximum = $global:amount
        if($userCount -ge 1000){
            $ChartArea1.AxisY.Interval = $inter - ($inter %100)
            $inter = [math]::Round($userCount/10,0)
        }elseif($userCount -ge 100){
            $ChartArea1.AxisY.Interval = $inter - ($inter %10)
            $inter = [math]::Round($userCount/20,0)
        }else{
            $ChartArea1.AxisY.Interval = $inter - ($inter %10)
            $inter = [math]::Round($userCount/10,0)
        }
    }else{
        $ChartArea1.AxisY.Maximum = $userCount
        
        if($userCount -ge 1000){
            $ChartArea1.AxisY.Interval = $inter - ($inter %100)
            $inter = [math]::Round($userCount/10,0)
        }elseif($userCount -ge 100){
            $ChartArea1.AxisY.Interval = $inter - ($inter %10)
            $inter = [math]::Round($userCount/20,0)
        }else{
            $ChartArea1.AxisY.Interval = $inter - ($inter %10)
            $inter = [math]::Round($userCount/10,0)
        }

    }
    
    $Chart1.Width = 1000
    $Chart1.Height = 700
    $Chart1.Left = 10
    $Chart1.Top = 10
    $Chart1.BackColor = [System.Drawing.Color]::White
    $Chart1.BorderColor = 'Black'
    $Chart1.BorderDashStyle = 'Solid'      
    $ChartTitle1 = New-Object System.Windows.Forms.DataVisualization.Charting.Title
    $ChartTitle1.Text = $title
    $Font1 = New-Object System.Drawing.Font @('Microsoft Sans Serif','12', [System.Drawing.FontStyle]::Bold)
    $ChartTitle1.Font =$Font1
    $Chart1.Titles.Add($ChartTitle1)

    $testPath = Test-Path $global:outFilePicBar
    if($testPath -eq $True){
        $global:check1 += 1      
        $global:outFilePicBar = $($PSScriptRoot)+"\Bar-$($dateTimeFile)-$($global:check1).jpeg"         
    }
    $global:IncludeImages.Add($global:outFilePicBar)
    $Chart1.SaveImage("$outFilePicBar", 'jpeg')
}
drawPie -hash $mailHash -title "Emails Availability" |Out-Null
drawPie -hash $accExHash -title "Expired Accounts"|Out-Null
drawPie -hash $accStatusHash -title "Account Status"|Out-Null
drawPie -hash $smartReHash -title "Smart Cards Required"|Out-Null
drawPie -hash $passReHash -title "Password Required"|Out-Null
#drawPie -hash $passChangedHash -title "Password CANNOT Change"|Out-Null
drawPie -hash $passExpHash -title "Password Never Expired Settings"|Out-Null
drawBar -hash $lastLogonHash -title  "Last Logon Date"|Out-Null
drawBar -hash $passSetHash -title "Password Last Changed"|Out-Null
#drawBar -Hash $badPassCHash -title "Bad Password Count"|Out-Null
drawBar -hash $lastBadLogHash -title "Last Bad Logon Attempts"|Out-Null
drawBar -hash $ageHash -title "Password Expiration Date"|Out-Null
drawBar -hash $lastModihash -title "User's Objects Latest Modification"|Out-Null
$userName = Get-ADUser -filter * -Properties DistinguishedName| ?{$_.sAMAccountName -match $env:UserName }|select Name|Out-String
$userName = $userName -replace '-', ' ' -replace 'Name', ''
$userName = $userName.Trim()
$trustedDo = Get-ADTrust -Filter * -Server $Domain | select Name |Out-String
$trustedDo = $trustedDo  -replace '-','' -replace 'Name','' 
$trustedDo =$trustedDo.Trim()
$adForest =  (get-ADForest -Server $Domain).domains | Out-String
if([string]::IsNullOrEmpty($global:amount)){
    $global:amount = $userCount
}
$admin = Get-ADGroupMember "Domain ADmins" -Server $Domain| select name,distinguishedName |measure
$admin = $admin.count
$domainCName = Get-ADDomainController -Filter * -Server $Domain| select Name|Out-String
$domainCName = $domainCName -replace '-', ' ' -replace 'Name', ''
$domainCName = $domainCName.Trim()
$domainCoper = Get-ADDomainController -Filter * -Server $Domain| select operatingsystem|Out-String
$domainCoper = $domainCoper -replace '-', ' ' -replace 'Name', '' -replace 'operatingsystem',''
$domainCoper = $domainCoper.Trim()
$ipAddress = Get-NetIPAddress | ?{($_.InterfaceAlias -match "Public") -and ($_.AddressFamily -match "Ipv4")}|select IPAddress|Out-String
$ipAddress = $ipAddress -replace '-', ' ' -replace 'IPAddress', ''
$ipAddress = $ipAddress.Trim()
$body =@'
<h1> Forest Report </h1>
<p><ins><b>I.<b> Information<ins></p>
<div class="tabofexecu">
    <table class="tabexecu" >
 
          <tr>
            <td>Object Category:</td>
            <td>{8}</td> 
          </tr>
          <tr>
            <td>Object Class: </td>
            <td>{9}</td> 
          </tr>
  
          <tr>
            <td>Amount of Data: </td>
            <td>{10}</td> 
          </tr>      
    </table>
<div>

<div class="tablehere">
    <table class="tabinfo" > 
          <tr>
            <td>Domain:</td>
            <td>{0}</td> 
          </tr>
          <tr>
            <td>User Domain: </td>
            <td>{1}</td> 
          </tr>
          <tr>
            <td>Computer Name:</td>
            <td>{2}</td> 
          </tr>
          <tr>
            <td>IP Address:</td>
            <td>{14}</td> 
          </tr>
          <tr>
            <td>Reported by: </td>
            <td>{3}</td> 
          </tr>
          <tr>
            <td>Execution Date: </td>
            <td>{4}</td> 
          </tr>
          <tr>
            <td>Retrieved Data from: </td>
            <td>{5}</td> 
          </tr>
    </table>
</div>

<p><ins><b>II.<b> Domain Summary<ins></p>
<div  class="secTable">
    <table class="tabforest" > 
          <tr>
            <td>Number of Domain Admins:</td>
            <td>{11}</td> 
          </tr>

          <tr>
            <td>Forest Domains:</td>
            <td>{6}</td> 
          </tr>
          <tr>
            <td>Trusted Domains: </td>
            <td>{7}</td> 
          </tr>
    </table>
</div>
<div class="tabdomaincon">
    <table class="tabdomain" > 
        <tr>
            <th>Domain Controllers</th>
            <th>Operating System</th> 
        </tr>
        <tr>           
            <td>{12}</td> 
            <td>{13}</td> 
        </tr>      
    </table>
<div>

<p><ins><b>III.<b> Data Illustration<ins></p>
'@ -f  $Domain ,$env:UserDomain, $env:ComputerName,$userName,$(get-date),$outFileMeg,$adForest,$trustedDo,$objectCategory,$objectClass,$global:amount,$admin,$domainCName,$domainCoper,$ipAddress
}
function Generate-Html {
    Param(
       
        $filehtml,
        [string[]]$IncludeImages
    )

    if ($IncludeImages){
        $ImageHTML = $IncludeImages | % {
        $ImageBits = [Convert]::ToBase64String((Get-Content $_ -Encoding Byte))
        "<center><img src=data:image/jpeg;base64,$($ImageBits) alt='My Image'/><center>"
    }
        ConvertTo-Html -Body $body -PreContent $imageHTML -Title "Report on $Domain" -CssUri "style.css" |
        Out-File $filehtml
    }
}



# Service accounts

function seracc{
    $global:servicAcc = $True
    ini
    if(($global:sam -match "^[pP]98[5..7]") -and ($global:sam -notmatch "^[pP]981")){ 
        tracking -fileName $outFileService 
             
    }
 

}
    
# Admin accounts

function adminacc{


}

#Main run here
$cls = cls
function main{
    $ADSearch.SearchRoot ="LDAP://$Domain"
    # distinguished Name method
    $arrayDN = @()
    if($dna -eq $true)
    {
        if($amountCheck -eq $true)
        {
            Write-Host
            Write-Verbose -Message  "Please be patient whilst the script retrieves all $global:amount distinguished names..." -Verbose        
        
            foreach ($user  in $userObjects)
            {
                if($count -lt $global:amount)
                {
                    $sam = $user.Properties.Item("sAMAccountName")[0]
                    $dn =  $user.Properties.Item("distinguishedName")[0]
                               
                    if($exportCheck -eq $true){
                        $global:exportedToTxt = $true
                        $dn | Out-File "$outFileTxt" -Append
                    }
                    elseif($exportCheck -eq $false){
                        $dn
                        #$arrayDN += $dn
                    }                 
                    $count++    
                    $TotalUsersProcessed++   
                
                }
                If ($ProgressBar) 
                {
                    Write-Progress -Activity "Processing $($global:amount) Users" -Status ("Count: 
                    $($TotalUsersProcessed)- Username: {0}" -f $sam) -PercentComplete (($TotalUsersProcessed/$global:amount)*100)
                }
            }
            #$arrayDN
        }    
        elseif($amountCheck -eq $false)
        {
            Write-Host
            Write-Verbose -Message  "Please be patient whilst the script retrieves all $userCount distinguished names..." -Verbose
            foreach ($user  in $userObjects)
            {
                $sam = $user.Properties.Item("sAMAccountName")[0]
                $dn =  $user.Properties.Item("distinguishedName")[0]
                if($exportCheck -eq $true)
                {
                        $global:exportedToTxt = $true
                        $dn | Out-File "$outFileTxt" -Append
                }
                elseif($exportCheck -eq $false)
                {
                        $dn
                        #$arrayDN += $dn
                } 
                $TotalUsersProcessed++
                If ($ProgressBar) 
                {
                    Write-Progress -Activity "Processing $($userCount) Users" -Status ("Count: 
                    $($TotalUsersProcessed)- Username: {0}" -f $sam) -PercentComplete (($TotalUsersProcessed/$userCount)*100)
                }
            }        
            #$arrayDN
        }
    }
    ## Finished distinguished Name method

    elseif($amountCheck -eq $true)
    {
        Write-Host
        Write-Verbose -Message  "Please be patient whilst the script retrieves all $global:amount distinguished names..." -Verbose
        foreach ($user  in $userObjects)
        {
            if($count -lt $global:amount)
            {
                #tracking
                seracc
                $TotalUsersProcessed++
                $count++
                If ($ProgressBar) 
                {                
                    Write-Progress -Activity "Processing $($global:amount) Users" -Status ("Count: 
                    $($TotalUsersProcessed)- Username: {0}" -f $sam) -PercentComplete (($TotalUsersProcessed/$global:amount)*100)              
                }
            
            }
        }
    }
    elseif($amountCheck -eq $false)
    {
        Write-Host
        Write-Verbose -Message  "Please be patient whilst the script retrieves all $userCount distinguished names..." -Verbose
        foreach ($user  in $userObjects)
        {    
            #tracking
            seracc
        
            $TotalUsersProcessed++
            If ($ProgressBar) 
            {
               
                Write-Progress -Activity "Processing $($userCount) Users" -Status ("Count: 
                $($TotalUsersProcessed)- Username: {0}" -f $sam) -PercentComplete (($TotalUsersProcessed/$userCount)*100)
            }
        }
    }
}
#optional choices
function optional{
    write-host
    write-host " 1. Get distinguished name "
    write-host " 2. Get all supplied attributes"
    write-host
    $methods = Read-Host -Prompt "Option "
    if($methods -eq 1)
    {
        $dna = $true
    }
    elseif ($methods -eq 2)
    {
        $dna = $false
    }else
    {
        Write-Verbose -Message  "Option is not valid" -Verbose
        exit
    }

    #Amount
    $global:amount = Read-Host -Prompt "Amount of data (Enter to get all data)"
    if($global:amount -eq ""){        
        $amountCheck = $false
    }
    else
    {        
        $amountCheck = $true
    }    
    #Export
    $export = Read-Host -Prompt "Do you want to export the data? (y/n)"
    if(($export -eq "y") -or ($export -eq ""))
    {
        $exportCheck = $true
    }
    elseif($export -eq "n")
    {
         $exportCheck = $false
    }
    else
    {
        Write-Verbose -Message  "Option is not valid" -Verbose
        exit
    }
    main
    
    
}
#Options
if($type -eq 1)
{  
    optional  
}
elseif($type -eq 2)
{
    optional   
}
else{

    Write-Verbose -Message  "Option is not valid" -Verbose
    exit
}


if($exportedToCSV -eq $true){
        Write-Host
      <#
        modiOne
        
        $CSV1 = Import-Csv $outFileModi
        $CSV2 = Import-Csv $outFileService

        $CSV2 | ForEach-Object -Begin {$i = 0} {  
        $_ | Add-Member -MemberType NoteProperty -Name "User's Objects lastest Modification" -Value $CSV1[$i++].modi -PassThru 
                    } | Export-Csv $outFileMeg -NoTypeInformation
        rm $outFileModi
         $outFileService
         #>
        Write-Host "Data has been exported to $outFileService" -foregroundcolor "magenta"
        
}
if($exportedToTxt -eq $true){
        Write-Host
        Write-Host "Data has been exported to $outFileTxt" -foregroundcolor "magenta"
}

#Finish
Write-Host
Write-Verbose -Message  "Script Finished!!" -Verbose

function writeHTML{

    param($filename)
    html
    Generate-Html -filehtml $filename -IncludeImages $global:IncludeImages
    foreach($image in $global:IncludeImages){
            rm $image 
    }
}

if($global:servicAcc -eq $true){
    writeHTML -filename $outFileHTMLService
}
