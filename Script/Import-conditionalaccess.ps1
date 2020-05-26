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
$redirectUri = "http://localhost"

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



####################################################################################################################################################################
######################################################## Part for Import from csv ##################################################################################

#load assembly
Add-Type -AssemblyName "System.Windows.Forms"

# prompt for choosing file to Import 
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    
}
$null = $FileBrowser.ShowDialog()
$file = $filebrowser.filename

#get all users 
$apiUrl = "https://graph.microsoft.com/v1.0/users"
$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
$allusers = (($Data | select-object Value).Value)

#get all users 
$apiUrl = "https://graph.microsoft.com/v1.0/groups"
$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
$allgroups = (($Data | select-object Value).Value)

#get all Roles
$apiUrl = "https://graph.microsoft.com/v1.0/directoryRoleTemplates"
$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
$ALLRole = ($Data | select-object Value).Value

# get all locations
$apiUrl = "https://graph.microsoft.com/beta/conditionalAccess/namedLocations"
$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
$allLocations= ($Data | select-object Value).Value

# get all Applications 
$apiUrl = "https://graph.microsoft.com/beta/servicePrincipals"
$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
$allApplications= ($Data | select-object Value).Value


foreach($line in (Import-Csv "$file" -Delimiter "," ))
{
    
    #define displayname 
    $displayname = $null
    if ($($line.displayname) -eq "")
    {
        $displayname = $null
    }
    else
    {
        $displayname = '"' + "$($line.displayname)" + '"'
    }

    #define State
    $state = $null
    if ( $($line.state) -eq "")
    {
        $state = $null
    }
    else
    {
        $state = '"' + "$($line.state)" + '"'
    }

    # define grantoperator
    $operator = $null
    if ($($line.grantoperator) -eq "")
    {
        $operator = $null
    }
    else
    {
        $operator = '"' + "$($line.grantoperator)" + '"'
    }
    
    # define grantbuiltInControls
    $grantbuiltInControls = $null
    if (($ligne.grantbuiltInControls) -eq "")
    {
        $grantbuiltInControls = $null
    }
    else
    {
        $grantbuiltInControls = '"' + "$($ligne.grantbuiltInControls)" + '"'
    }

    # define customAuthenticationFactors
    $customAuthenticationFactors = $null
    if ($line.customAuthenticationFactors -eq "")
    {
        $customAuthenticationFactors = $null
    }
    else
    {
        $customAuthenticationFactors = '"' + "$($line.customAuthenticationFactors)" + '"'
    }
    # define termofuse 
    $termsOfUse = $null
    if ( $line.termsOfUse -eq "")
    {
        $termsOfUse = $null
    }
    else
    {
        $termsOfUse = '"' + "$($line.termsOfUse)" + '"'
    }

    # define  $signInRiskLevels
    $signInRiskLevels = $null
    if ($line.signInRiskLevels -eq "")
    {
        $signInRiskLevels = $null
    }
    elseif ((($line.signInRiskLevels).Split(" ")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.signInRiskLevels).Split(" "))
        {
            $count++
            if ( $count -eq (($line.signInRiskLevels).Split(" ")).count)
            {
                $signInRiskLevels+= '"' + "$i" + '"'
            }
            else 
            {
                $signInRiskLevels+= '"' +"$i" + '"' + ","
            }
        }
    }


    # define  $clientAppTypes
    $clientAppTypes = $null
    if ($line.clientAppTypes -eq "")
    {
        $clientAppTypes = $null
    }
    elseif ((($line.clientAppTypes).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.clientAppTypes).Split(" "))
        {
            $count++
            if ( $count -eq (($line.clientAppTypes).Split(" ")).count)
            {
                $clientAppTypes+= '"' + $i + '"'
            }
            else 
            {
                $clientAppTypes+= '"' + $i + '"' + ","
            }
        }
    }

    # define  $Includeplatforms
    $Includeplatforms = $null
    if ($line.Includeplatforms -eq "")
    {
        $Includeplatforms = $null
    }
    elseif ((($line.Includeplatforms).Split(" ")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.Includeplatforms).Split(" "))
        {
            $count++
            if ( $count -eq (($line.Includeplatforms).Split(" ")).count)
            {
                $Includeplatforms+= '"' + "$i" + '"'
            }
            else 
            {
                $Includeplatforms+= '"' +"$i" + '"' + ","
            }
        }
    }

    # define  $excludeplatforms
    $excludeplatforms = $null
    if ($line.excludeplatforms -eq "")
    {
        $excludeplatforms = $null
    }
    elseif ((($line.excludeplatforms).Split(" ")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeplatforms).Split(" "))
        {
            $count++
            if ( $count -eq (($line.excludeplatforms).Split(" ")).count)
            {
                $excludeplatforms+= '"' + "$i" + '"'
            }
            else 
            {
                $excludeplatforms+= '"' +"$i" + '"' + ","
            }
        }
    }

    # define  $includeLocations
    $includeLocations = $null
    if ($line.includeLocations -eq "")
    {
        $includeLocations = $null
    }
    elseif ($line.includeLocations -eq "All")
    {
        $includeLocations = '"' + "$($line.includeLocations)" + '"'
    }
    elseif ((($line.includeLocations).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeLocations).Split(";"))
        {
            $Locations= $null
            $count++
            if ( $count -eq (($line.includeLocations).Split(";")).count)
            {
                $locations = ($alllocations | Where-Object { $_.displayname -eq "$i"}).id
                $includeLocations+= '"' + "$locations" + '"'
            }
            else 
            {
                $locations = ($alllocations | Where-Object { $_.displayname -eq "$i"}).id
                $includeLocations+= '"' + "$locations" + '"' + ","
            }
        }
    }
    # define  $excludeLocations
    $excludeLocations = $null
    if ($line.excludeLocations -eq "")
    {
        $excludeLocations = $null
        pause
    }
    elseif ($line.excludeLocations -eq "All")
    {
        $excludeLocations = '"' + "$($line.excludeLocations)" + '"'
        pause
    }
    elseif ((($line.excludeLocations).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeLocations).Split(";"))
        {
            $Locations= $null
            $count++
            if ($count -eq (($line.excludeLocations).Split(";")).count)
            {
                $locations = ($alllocations | Where-Object { $_.displayname -eq "$i"}).id
                $excludeLocations+= '"' + "$locations" + '"'
                $count
            }
            else 
            {
                $locations = ($alllocations | Where-Object { $_.displayname -eq "$i"}).id
                $excludeLocations+= '"' + "$locations" + '"' + ","
                $count
            }
        }
    }

    # define  $includeDeviceStates
    $includeDeviceStates = $null
    if ($line.includeDeviceStates -eq "")
    {
        $includeDeviceStates = $null
    }
    elseif ((($line.includeDeviceStates).Split(" ")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeDeviceStates).Split(" "))
        {
            $count++
            if ( $count -eq (($line.includeDeviceStates).Split(" ")).count)
            {
                $includeDeviceStates+= '"' + "$i" + '"'
            }
            else 
            {
                $includeDeviceStates+= '"' + "$i" + '"' + ","
            }
        }
    }

    # define  $excludeDeviceStates
    $excludeDeviceStates = $null
    if ($line.excludeDeviceStates -eq "")
    {
        $excludeDeviceStates = $null
    }
    elseif ((($line.excludeDeviceStates).Split(" ")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeDeviceStates).Split(" "))
        {
            $count++
            if ( $count -eq (($line.excludeDeviceStates).Split(" ")).count)
            {
                $excludeDeviceStates+= '"' + "$i" + '"'
            }
            else 
            {
                $excludeDeviceStates+= '"' +"$i" + '"' + ","
            }
        }
    }

    # define  $includeApplications
    $includeApplications = $null
    if ($line.includeApplications -eq "")
    {
        $includeApplications = $null
    }
    elseif ($line.includeApplications -eq "All ")
    {
        $includeApplications = '"' + "$($line.includeApplications)" + '"'
    }
    elseif ($line.includeApplications -eq "None ")
    {
        $includeApplications = '"' + "$($line.includeApplications)" + '"'
    }
    elseif ($line.includeApplications -eq "Office365 ")
    {
        $includeApplications = '"' + "$($line.includeApplications)" + '"'
    }
    elseif ((($line.includeApplications).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeApplications).Split(";"))
        {
            $applications= $null
            $count++
            if ( $count -eq (($line.includeApplications).Split(";")).count)
            {
                $applications = ($allapplications | Where-Object { $_.displayname -eq "$i"}).id
                $includeApplications+= '"' + "$applications" + '"'
            }
            else 
            {
                $applications = ($allapplications | Where-Object { $_.displayname -eq "$i"}).id
                $includeApplications+= '"' + "$applications" + '"' + ","
            }
        }
    } 

    # define  $excludeApplications
    $excludeApplications = $null
    if ($line.excludeApplications -eq "")
    {
        $excludeApplications = $null
    }
    elseif ($line.excludeApplications -eq "All ")
    {
        $excludeApplications = '"' + "$($line.excludeApplications)" + '"'
    }
    elseif ($line.excludeApplications -eq "None ")
    {
        $excludeApplications = '"' + "$($line.excludeApplications)" + '"'
    }
    elseif ($line.excludeApplications -eq "Office365 ")
    {
        $excludeApplications = '"' + "$($line.excludeApplications)" + '"'
    }
    elseif ((($line.excludeApplications).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeApplications).Split(";"))
        {
            $applications= $null
            $count++
            if ( $count -eq (($line.excludeApplications).Split(";")).count)
            {
                $applications = ($allapplications | Where-Object { $_.displayname -eq "$i"}).id
                $excludeApplications+= '"' + "$applications" + '"'
            }
            else 
            {
                $applications = ($allapplications | Where-Object { $_.displayname -eq "$i"}).id
                $excludeApplications+= '"' + "$applications" + '"' + ","
            }
        }
    } 

    # define includeUserActions
    $includeUserActions = $null
    if($($line.includeUserActions) -eq "")
    {
        $includeUserActions = $null
    }
    else
    {
        $includeUserActions = "$($line.includeUserActions)"
    }

    # define  $includeUsers
    $includeUsers = $null
    if ($line.includeUsers -eq "")
    {
        $includeUsers = $null
    }
    elseif ($line.includeUsers -eq "All;")
    {
        $includeUsers = '"' + "$($line.includeUsers)" + '"'
    }
    elseif ($line.includeUsers -eq "GuestsOrExternalUsers ")
    {
        $includeUsers = '"' + "$($line.includeUsers)" + '"'
    }
    elseif ((($line.includeUsers).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeUsers).Split(";"))
        {
            $users= $null
            $count++
            if($1 -eq "")
            {
                $users= $null
            }
            elseif ( ($count-1) -eq (($line.includeUsers).Split(";")).count)
            {
                $users = ($allusers | Where-Object { $_.displayname -eq "$i"}).id
                $includeUsers+= '"' + "$users" + '"'
            }
            else 
            {
                $users = ($allusers | Where-Object { $_.displayname -eq "$i"}).id
                $includeUsers+= '"' + "$users" + '"' + ","
            }
        }
    }


# define  $excludeUsers
    $excludeUsers = $null
    if ($line.excludeUsers -eq "")
    {
        $excludeUsers = $null
    }
    elseif ($line.excludeUsers -eq "All;")
    {
        $excludeUsers = '"' + "$($line.excludeUsers)" + '"'
    }
    elseif ($line.excludeUsers -eq "GuestsOrExternalUsers ")
    {
        $excludeUsers = '"' + "$($line.excludeUsers)" + '"'
    }
    elseif ((($line.excludeUsers).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeUsers).Split(";"))
        {
            $users= $null
            $count++
            if($i -eq "")
            {
                $users= $null
            }
            elseif ( $count -eq (($line.excludeUsers).Split(";")).count)
            {
                $users = ($allusers | Where-Object { $_.displayname -eq "$i"}).id
                $excludeUsers+= '"' + "$users" + '"'
            }
            else 
            {
                $users = ($allusers | Where-Object { $_.displayname -eq "$i"}).id
                $excludeUsers+= '"' + "$users" + '"' + ","
            }
        }
    }


    # define  $includeGroups
    $includeGroups = $null
    if ($line.includeGroups -eq "")
    {
        $includeGroups = $null
    }
    elseif ($line.includeGroups -eq "All")
    {
        $includeGroups = '"' + "$($line.includeGroups)" + '"'
    }
    elseif ($line.includeGroups -eq "GuestsOrExternalusers")
    {
        $includeGroups = '"' + "$($line.includeGroups)" + '"'
    }
    elseif ((($line.includeGroups).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeGroups).Split(";"))
        {
            $groups= $null
            $count++
            if ( $count -eq (($line.includeGroups).Split(";")).count)
            {
                $groups = ($allgroups | Where-Object { $_.displayname -eq "$i"}).id
                $includeGroups+= '"' + "$groups" + '"'
            }
            else 
            {
                $groups = ($allgroups | Where-Object { $_.displayname -eq "$i"}).id
                $includeGroups+= '"' + "$groups" + '"' + ","
            }
        }
    }
    

    # define  $excludeGroups
    $excludeGroups = $null
    if ($line.excludeGroups -eq "")
    {
        $excludeGroups = $null
    }
    elseif ($line.excludeGroups -eq "All")
    {
        $excludeGroups = '"' + "$($line.excludeGroups)" + '"'
    }
    elseif ($line.excludeGroups -eq "GuestsOrExternalusers")
    {
        $excludeGroups = '"' + "$($line.excludeGroups)" + '"'
    }
    elseif ((($line.excludeGroups).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeGroups).Split(";"))
        {
            $groups= $null
            $count++
            if ( $count -eq (($line.excludeGroups).Split(";")).count)
            {
                $groups = ($allgroups | Where-Object { $_.displayname -eq "$i"}).id
                $excludeGroups+= '"' + "$groups" + '"'
            }
            else 
            {
                $groups = ($allgroups | Where-Object { $_.displayname -eq "$i"}).id
                $excludeGroups+= '"' + "$groups" + '"' + ","
            }
        }
    }

    # define  $IncludeRoles
    $includeRoles = $null
    if ($line.includeroles -eq "")
    {
        $includeRoles = $null
    }
    elseif ($line.includeroles -eq "All")
    {
        $includeRoles = '"' + "$($line.includeroles)" + '"'
    }
    
    elseif ((($line.includeroles).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeroles).Split(";"))
        {
            $roles= $null
            $count++
            if ( $count -eq (($line.includeroles).Split(";")).count)
            {
                $roles = ($allrole | Where-Object { $_.displayname -eq "$i"}).id
                $includeRoles+= '"' + "$roles" + '"'
            }
            else 
            {
                $roles = ($allrole | Where-Object { $_.displayname -eq "$i"}).id
                $includeRoles+= '"' + "$roles" + '"' + ";"
            }
        }
    }


    # define  $excludeRoles
    $excludeRoles = $null
    if ($line.excludeRoles -eq "")
    {
        $excludeRoles = $null
    }
    elseif ($line.excludeRoles -eq "All")
    {
        $excludeRoles = '"' + "$($line.excludeRoles)" + '"'
    }
    
    elseif ((($line.excludeRoles).Split(";")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeRoles).Split(";"))
        {
            $roles= $null
            $count++
            if ( $count -eq (($line.excludeRoles).Split(";")).count)
            {
                $roles = ($allrole | Where-Object { $_.displayname -eq "$i"}).id
                $excludeRoles+= '"' + "$roles" + '"'
            }
            else 
            {
                $roles = ($allrole | Where-Object { $_.displayname -eq "$i"}).id
                $excludeRoles+= '"' + "$roles" + '"' + ","
            }
        }
    }

    # define $sessionControls
    $sessionControls = $null
    if($($line.sessionControls) -eq "")
    {
        $sessionControls = $null
    }
    else
    {
        $sessionControls = "$($line.sessionControls)"
    }

    #################################################################
    # define .json file 
    $line 
    $ConditionalAccessPolicies = @"
{
    "displayName": $displayname,
    "state": $state,
    "grantControls": {
    "operator": $operator,
    "builtInControls": [
    $grantbuiltInControl
    ],
    "customAuthenticationFactors": [
    $customAuthenticationFactors
    ],
    "termsOfUse": [
    $termsOfUse
    ]
    },
    "conditions": {
    "signInRiskLevels": [
    $signInRiskLevels
    ],
    "clientAppTypes": [
    $clientAppTypes
    ],
    "platforms": {
      "includePlatforms": [
      $Includeplatforms
      ],
      "excludePlatforms": [
      $excludeplatforms
      ]
    },
    "locations": {
      "includeLocations": [
       $includeLocations
      ],
      "excludeLocations": [
      $excludeLocations
      ]
    },
    "devices": {
      "includeDeviceStates": [
      $includeDeviceStates
      ],
      "excludeDeviceStates": [
      $excludeDeviceStates
      ]
    },
    "applications": {
      "includeApplications": [
      $includeApplications
      ],
      "excludeApplications": [
      $excludeApplications
      ],
      "includeUserActions": [
      $includeUserActions
      ]
    },
    "users": {
      "includeUsers": [
      $includeUsers
      ],
      "excludeUsers": [
      $excludeUsers
      ],
      "includeGroups": [
      $includeGroups
      ],
      "excludeGroups": [
      $excludeGroups
      ],
      "includeRoles": [
      $includeRoles
      ],
      "excludeRoles": [
      $excludeRoles
      ]
    }
  },
  "sessionControls": $sessionControls
}
"@
$json = $PSScriptRoot + "\json\" + "$($line.id)" + ".json"

Add-Content -Value $ConditionalAccessPolicies -PassThru $json

pause
}



