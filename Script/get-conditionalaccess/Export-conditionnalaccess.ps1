###########################################################################################################################################################
##############   Version 1.2               ################################################################################################################
##############   Modification : 05/23/2020 ################################################################################################################
###########################################################################################################################################################

###########################################################################################################################################################
######################################################## Part connect to Graph APIv #######################################################################
# The resource URI
$resource = "https://graph.microsoft.com"
# Your Client ID and Client Secret obainted when registering your WebApp
$clientid = "01a98767-df6e-4cdf-9752-c4e2f2d92a91"
$clientSecret = "3l6SnV35..sB.1-m9L0MGYhQS.A8.MJbG-"
$redirectUri = "https://lazyadministrator.net"

# UrlEncode the ClientID and ClientSecret and URL's for special characters 
Add-Type -AssemblyName System.Web
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($clientid)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($clientSecret)
$redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUri)
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resource)
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/user.readwrite.all")

# Function to popup Auth Dialog Windows Form
Function Get-AuthCode {
    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }

    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri        
        if ($Global:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }

    $output
}


# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUriEncoded&client_id=$clientID&resource=$resourceEncoded&prompt=admin_consent&scope=$scopeEncoded"
Get-AuthCode
# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode  = ($uri | Select-string -pattern $regex).Matches[0].Value

Write-output "Received an authCode, $authCode"


#get Access Token
$body = "grant_type=authorization_code&redirect_uri=$redirectUri&client_id=$clientId&client_secret=$clientSecretEncoded&code=$authCode&resource=$resource"
$tokenResponse = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
    -Method Post -ContentType "application/x-www-form-urlencoded" `
    -Body $body `
    -ErrorAction STOP



###########################################################################################################################################################
######################################################## Part for export ##################################################################################

# export file 
$exportconditionalaccess = $PSScriptRoot + ".\Condtionalaccess.csv"

$allconditionalaccess = @()

# get all conditional access 
$apiUrl = 'https://graph.microsoft.com/beta/conditionalAccess/policies'
$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
$conditionalaccess = ($Data | select-object Value).Value

# loop for all conditional access 
foreach ($conditionalaccess in $conditionalaccess)
{
    $Output = New-Object PSObject
    $groupinclude = $null 
    $groupexclude = $null 
    $userInclude = $null
    $userexclude = $null
    $includeApplications = $null
    $excludeApplications = $null
    $IncludeRole = $null
    $ExcludeRole = $null
    $includeLocations = $null
    $excludeLocations = $null


    
    ###################################################### part get global information ################################################## 
    $output | Add-Member NoteProperty -Name "id" -Value "$($conditionalaccess.id)"
    $output | Add-Member NoteProperty -Name "displayname" -value "$($conditionalaccess.displayName)"
    $output | Add-Member NoteProperty -Name "createdatetime" -value "$($conditionalaccess.createdDateTime)"
    $output | Add-Member NoteProperty -Name "modifiedDateTime" -value "$($conditionalaccess.modifiedDateTime)"
    $output | Add-Member NoteProperty -Name "state" -value "$($conditionalaccess.state)"
    $output | Add-Member NoteProperty -Name "sessionControls" -value "$($conditionalaccess.sessionControls)"
    

    ###################################################### part get Conditions ##########################################################
    
    #loop get Include App displayname 
    $($conditionalaccess.conditions.applications.includeApplications) | % {
        if ($_ -eq "ALL")
        {
            $includeApplications+= $_ + " "
        }
        elseif ($_ -eq "Office365")
        {
            $includeApplications+= $_ + " "
        }
        elseif ($_ -eq "none")
        {
            $includeApplications+= $_ + " "
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=appid eq '$_'"
            $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
            $includeApplications+= (($Data | select-object Value).Value).displayName + " "        
        }
    }
    $output | Add-Member NoteProperty -Name "includeApplications" -value  "$includeApplications"

    #loop get exclude App displayname 
    $($conditionalaccess.conditions.applications.excludeApplications) | % {
        if ($_ -eq "ALL")
        {
            $excludeApplications+= $_ + " "
        }
        elseif ($_ -eq "Office365")
        {
            $excludeApplications+= $_ +" "
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=appid eq '$_'"
            $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
            $excludeApplications+= (($Data | select-object Value).Value).displayName + " "        
        }
    }
    $output | Add-Member NoteProperty -Name "excludeApplications" -value  "$excludeApplications"

    $output | Add-Member NoteProperty -Name "includeUserActions" -value "$($conditionalaccess.conditions.applications.includeUserActions)"
    $output | Add-Member NoteProperty -Name "signInRiskLevels" -value "$($conditionalaccess.conditions.signInRiskLevels)"
    $output | Add-Member NoteProperty -Name "clientAppTypes" -value "$($conditionalaccess.conditions.clientAppTypes)"
    $output | Add-Member NoteProperty -Name "deviceStates" -value "$($conditionalaccess.conditions.deviceStates)"
    
    #loop get Include  user displayname 
    $($conditionalaccess.conditions.users.includeusers) | % {
        if ($_ -eq "ALL")
        {
            $userInclude+= $_ +" "
        }
        elseif ($_ -eq "GuestsOrExternalUsers")
        {
            $userInclude+= $_ +" "
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$_'"
            $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
            $userInclude+= (($Data | select-object Value).Value).displayName + " " 
        
        }
    }
    $output | Add-Member NoteProperty -Name "Includeusers" -value  "$userinclude"

    #loop get Exclude user displayname 

    $($conditionalaccess.conditions.users.excludeusers) | % {
        if ($_ -eq "ALL")
        {
            $userexclude+= $_ +" "
        }
        elseif ($_ -eq "GuestsOrExternalUsers")
        {
            $userexclude+= $_ +" "
        }
        else
        {    
            $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$_'"
            $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
            $userexclude+= (($Data | select-object Value).Value).displayName + " "
        }
     }
     $output | Add-Member NoteProperty -Name "excludeUsers" -value "$userexclude"


    #loop get Include group displayname  
    $($conditionalaccess.conditions.users.includeGroups) | % {
        $apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=id eq '$_'"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $groupInclude+= (($Data | select-object Value).Value).displayName + " " 
    }
    $output | Add-Member NoteProperty -Name "includeGroups" -value "$groupInclude"

    #loop get Exclude group displayname
    $($conditionalaccess.conditions.users.excludeGroups) | % {
        $apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=id eq '$_'"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $groupexclude+= (($Data | select-object Value).Value).displayName + " " 
    }
    $output | Add-Member NoteProperty -Name "excludeGroups" -value "$groupexclude"

    #loop get Include groupManagement  displayname
    $($conditionalaccess.conditions.users.includeRoles) | % {
        $apiUrl = "https://graph.microsoft.com/v1.0/directoryRoleTemplates?`$filter=id eq '$_'"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $IncludeRole+= (($Data | select-object Value).Value).displayName + " " 
    }
    $output | Add-Member NoteProperty -Name "includeRoles" -value "$IncludeRole"
    
    #loop get exclude groupManagement  displayname
    $($conditionalaccess.conditions.users.excludeRoles) | % {
        $apiUrl = "https://graph.microsoft.com/v1.0/directoryRoleTemplates?`$filter=id eq '$_'"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $excludeRole+= (($Data | select-object Value).Value).displayName + " " 
    }
    $output | Add-Member NoteProperty -Name "excludeRoles" -value "$excludeRole"

    $output | Add-Member NoteProperty -Name "Includeplatforms" -value "$($conditionalaccess.conditions.platforms.includePlatforms)"
    $output | Add-Member NoteProperty -Name "excludeplatforms" -value "$($conditionalaccess.conditions.platforms.excludePlatforms)"


    #loop get Include location displayname
    $($conditionalaccess.conditions.locations.includeLocations) | % {
        if ($_ -eq "ALL")
        {
            $includeLocations+= $_ 
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/beta/conditionalAccess/namedLocations?`$filter=id eq '$_'"
            $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
            $includeLocations+= (($Data | select-object Value).Value).displayName  
        }
    }
    $output | Add-Member NoteProperty -Name "includeLocations" -value "$includeLocations"
    
    #loop get exclude location  displayname
    $($conditionalaccess.conditions.locations.excludeLocations) | % {
        if ($_ -eq "ALL")
        {
            $excludeLocations+= $_ 
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/beta/conditionalAccess/namedLocations?`$filter=id eq '$_'"
            $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
            $excludeLocations+= (($Data | select-object Value).Value).displayName + " " 
        }
    }
    $output | Add-Member NoteProperty -Name "excludeLocations" -value "$excludeLocations"

    $output | Add-Member NoteProperty -Name "includeDeviceStates" -value "$($conditionalaccess.conditions.devices.includeDeviceStates)"
    $output | Add-Member NoteProperty -Name "excludeDeviceStates" -value "$($conditionalaccess.conditions.devices.excludeDeviceStates)"
    
    ###################################################### part get grant Controls ################################################## 

    $output | Add-Member NoteProperty -Name "grantoperator" -value "$($conditionalaccess.grantControls.operator)"
    $output | Add-Member NoteProperty -Name "grantbuiltInControls" -value "$($conditionalaccess.grantControls.builtInControls)"
    $output | Add-Member NoteProperty -Name "customAuthenticationFactors" -value "$($conditionalaccess.grantControls.customAuthenticationFactors)"
    $output | Add-Member NoteProperty -Name "termsOfUse" -value "$($conditionalaccess.grantControls.termsOfUse)"

    $allconditionalaccess+= $Output

$($conditionalaccess.displayName)
}

# Export to csv 
$allconditionalaccess | Export-Csv -Path $exportconditionalaccess -NoTypeInformation

 
