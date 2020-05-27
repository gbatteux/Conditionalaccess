###########################################################################################################################################################
##############   Version 1.2               ################################################################################################################
##############   Modification : 05/23/2020 ################################################################################################################
###########################################################################################################################################################

$AccessToken = Connect-toGraph  

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

#get all Roles
$apiUrl = "https://graph.microsoft.com/v1.0/directoryRoleTemplates"
$allrole = Get-MsGraph -AccessToken $AccessToken -Uri $apiUrl 

#get all groups
$apiUrl = "https://graph.microsoft.com/v1.0/groups"
$allgroups = Get-MsGraph -AccessToken $AccessToken -Uri $apiUrl 

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
    
    if ($line.grantbuiltInControls -eq "block")
    {
        $grantbuiltInControls = "`r`n" + '"' + "$($line.grantbuiltInControls)" + '"' + "`r`n" 
    }
    elseif ((($line.grantbuiltInControls).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.grantbuiltInControls).Split("`r"))
        {
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif( $count -eq ((($line.grantbuiltInControls).Split("`r")).count-1))
            {
                $grantbuiltInControls+= "`r`n" + '"' + "$i" + '"' + "`r`n" 
            }
            else 
            {
                $grantbuiltInControls+= "`r`n" + '"' + "$i" + '"' + ","

                
            }
        }
    }

    # define customAuthenticationFactors
    $customAuthenticationFactors = $null
    if ($line.customAuthenticationFactors -eq "")
    {
        $customAuthenticationFactors = $null
    }
    else
    {
        $customAuthenticationFactors = "`r`n" + '"' + "$($line.customAuthenticationFactors)" + '"' + "`r`n"
    }
    # define termofuse 
    $termsOfUse = $null
    if ( $line.termsOfUse -eq "")
    {
        $termsOfUse = $null
    }
    else
    {
        $termsOfUse = "`r`n" + '"' + "$($line.termsOfUse)" + '"' + "`r`n"
    }

    # define  $signInRiskLevels
    $signInRiskLevels = $null
    if ($line.signInRiskLevels -eq "")
    {
        $signInRiskLevels = $null
    }
    elseif ((($line.signInRiskLevels).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.signInRiskLevels).Split("`r"))
        {
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif( $count -eq ((($line.signInRiskLevels).Split("`r")).count-1))
            {
                $signInRiskLevels+= "`r`n" + '"' + "$i" + '"' + "`r`n"
            }
            else 
            {
                $signInRiskLevels+= "`r`n" + '"' +"$i" + '"' + ","
            }
        }
    }


    # define  $clientAppTypes
    $clientAppTypes = $null
    if ($line.clientAppTypes -eq "")
    {
        $clientAppTypes = $null
    }
    elseif ((($line.clientAppTypes).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.clientAppTypes).Split("`r"))
        {
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif( $count -eq ((($line.clientAppTypes).Split("`r")).count-1))
            {
                $clientAppTypes+= "`r`n" + '"' + $i + '"' + "`r`n"
            }
            else 
            {
                $clientAppTypes+= "`r`n" + '"' + $i + '"' + ","
            }
        }
    }

    # define  $Includeplatforms
    $Includeplatforms = $null
    if ($line.Includeplatforms -eq "")
    {
        $Includeplatforms = $null
    }
    elseif ((($line.Includeplatforms).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.Includeplatforms).Split("`r"))
        {
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif( $count -eq ((($line.Includeplatforms).Split("`r")).count-1))
            {
                $Includeplatforms+= "`r`n" + '"' + "$i" + '"' + "`r`n"
            }
            else 
            {
                $Includeplatforms+= "`r`n" + '"' +"$i" + '"' + ","
            }
        }
    }

    # define  $excludeplatforms
    $excludeplatforms = $null
    if ($line.excludeplatforms -eq "")
    {
        $excludeplatforms = $null
    }
    elseif ((($line.excludeplatforms).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeplatforms).Split("`r"))
        {
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif( $count -eq ((($line.excludeplatforms).Split("`r")).count-1))
            {
                $excludeplatforms+= "`r`n" + '"' + "$i" + '"' + "`r`n"
            }
            else 
            {
                $excludeplatforms+= "`r`n" + '"' +"$i" + '"' + ","
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
        $includeLocations = "`r`n" + '"' + "$($line.includeLocations)" + '"' + "`r`n"
    }
    elseif ((($line.includeLocations).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeLocations).Split("`r"))
        {
            $Locations= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif( $count -eq ((($line.includeLocations).Split("`r")).count-1))
            {
                $apiUrl = "https://graph.microsoft.com/beta/conditionalaccess/namedlocations?`$filter=displayname eq '$i'"
                $locations = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $includeLocations+= "`r`n" + '"' + "$locations" + '"' + "`r`n"
            }
            else 
            {
                $apiUrl = "https://graph.microsoft.com/beta/conditionalaccess/namedlocations?`$filter=displayname eq '$i'"
                $locations = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $includeLocations+= "`r`n" + '"' + "$locations" + '"' + ","
            }
        }
    }
    # define  $excludeLocations
    $excludeLocations = $null
    if ($line.excludeLocations -eq "")
    {
        $excludeLocations = $null
    }
    elseif ($line.excludeLocations -eq "All")
    {
        $excludeLocations = "`r`n" + '"' + "$($line.excludeLocations)" + '"' + "`r`n"
    }
    elseif ((($line.excludeLocations).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeLocations).Split("`r"))
        {
            $Locations= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif($count -eq ((($line.excludeLocations).Split("`r")).count-1))
            {
                $apiUrl = "https://graph.microsoft.com/beta/conditionalaccess/namedlocations?`$filter=displayname eq '$i'"
                $locations = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $excludeLocations+= "`r`n" + '"' + "$locations" + '"' + "`r`n"
                $count
            }
            else 
            {
                $apiUrl = "https://graph.microsoft.com/beta/conditionalaccess/namedlocations?`$filter=displayname eq '$i'"
                $locations = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $excludeLocations+= "`r`n" + '"' + "$locations" + '"' + ","
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
    elseif ((($line.includeDeviceStates).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeDeviceStates).Split("`r"))
        {
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif($count -eq ((($line.includeDeviceStates).Split("`r")).count-1))
            {
                $includeDeviceStates+= "`r`n" + '"' + "$i" + '"' + "`r`n"
            }
            else 
            {
                $includeDeviceStates+= "`r`n" + '"' + "$i" + '"' + ","
            }
        }
    }

    # define  $excludeDeviceStates
    $excludeDeviceStates = $null
    if ($line.excludeDeviceStates -eq "")
    {
        $excludeDeviceStates = $null
    }
    elseif ((($line.excludeDeviceStates).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeDeviceStates).Split("`r"))
        {
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif( $count -eq ((($line.excludeDeviceStates).Split("`r")).count-1))
            {
                $excludeDeviceStates+= "`r`n" + '"' + "$i" + '"' + "`r`n"
            }
            else 
            {
                $excludeDeviceStates+= "`r`n" + '"' +"$i" + '"' + ","
            }
        }
    }

    # define  $includeApplications
    $includeApplications = $null
    if ($line.includeApplications -eq "")
    {
        $includeApplications = $null
    }
    elseif ($line.includeApplications -eq "All")
    {
        $includeApplications = "`r`n" + '"' + "$($line.includeApplications)" + '"' + "`r`n"
    }
    elseif ($line.includeApplications -eq "None")
    {
        $includeApplications = "`r`n" + '"' + "$($line.includeApplications)" + '"' + "`r`n"
    }
    elseif ($line.includeApplications -eq "Office365")
    {
        $includeApplications = "`r`n" + '"' + "$($line.includeApplications)" + '"' + "`r`n"
    }
    elseif ((($line.includeApplications).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeApplications).Split("`r"))
        {
            $applications= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif($i -eq "Office365" -and $count -ne ((($line.includeApplications).Split("`r")).count-1))
            {
                $includeUsers+= "`r`n" + '"' + "$i" + '"' + ","
            }
            elseif( $count -eq ((($line.includeApplications).Split("`r")).count-1))
            {
                if($i -eq "Office365")
                {
                    $includeApplications+= "`r`n" + '"' + "$i" + '"' + "`r`n"
                }
                else
                {
                    $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=displayname eq '$i'"
                    $applications = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                    $includeApplications+= "`r`n" + '"' + "$applications" + '"' + "`r`n"
                }
            }
            else 
            {
                $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=displayname eq '$i'"
                $applications = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $includeApplications+= "`r`n" + '"' + "$applications" + '"' + ","
            }
        }
    } 

    # define  $excludeApplications
    $excludeApplications = $null
    if ($line.excludeApplications -eq "")
    {
        $excludeApplications = $null
    }
    elseif ($line.excludeApplications -eq "All")
    {
        $excludeApplications = "`r`n" + '"' + "$($line.excludeApplications)" + '"' + "`r`n"
    }
    elseif ($line.excludeApplications -eq "None")
    {
        $excludeApplications = "`r`n" + '"' + "$($line.excludeApplications)" + '"' + "`r`n"
    }
    elseif ($line.excludeApplications -eq "Office365")
    {
        $excludeApplications = "`r`n" + '"' + "$($line.excludeApplications)" + '"' + "`r`n"
    }
    elseif ((($line.excludeApplications).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeApplications).Split("`r"))
        {
            $applications= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif($i -eq "Office365" -and $count -ne ((($line.excludeApplications).Split("`r")).count-1))
            {
                $excludeApplications+= "`r`n" + '"' + "$i" + '"' + ","
            }
            elseif( $count -eq ((($line.excludeApplications).Split("`r")).count-1))
            {
                if($i -eq "Office365")
                {
                    $excludeApplications+= "`r`n" + '"' + "$i" + '"' + "`r`n"
                }
                else
                {
                    $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=displayname eq '$i'"
                    $applications = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                    $excludeApplications+= "`r`n" + '"' + "$applications" + '"' + "`r`n"
                }
            }
            else 
            {
                $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=displayname eq '$i'"
                $applications = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $excludeApplications+= "`r`n" + '"' + "$applications" + '"' + ","
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
        $includeUserActions = "`r`n" + "$($line.includeUserActions)" + "`r`n"
    }

    # define  $includeUsers
    $includeUsers = $null
    if ($line.includeUsers -eq "")
    {
        $includeUsers = $null
    }
    elseif ($line.includeUsers -eq "All")
    {
        $includeUsers = "`r`n" + '"' + "$($line.includeUsers)" + '"' + "`r`n"
    }
    elseif ($line.includeUsers -eq "GuestsOrExternalUsers ")
    {
        $includeUsers = "`r`n" + '"' + "$($line.includeUsers)" + '"' + "`r`n"
    }
    elseif ((($line.includeUsers).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeUsers).Split("`r"))
        {
            $users= $null
            $count++
            if($1 -eq "")
            {
                # do nothing last empty 
            }
            elseif($i -eq "GuestsOrExternalUsers" -and $count -ne ((($line.includeUsers).Split("`r")).count-1))
            {
                $includeUsers+= "`r`n" + '"' + "$i" + '"' + ","
            }
            elseif ( ($count-1) -eq ((($line.includeUsers).Split("`r")).count-1))
            {
                if($i -eq "GuestsOrExternalUsers")
                {
                    $excludeUsers+= "`r`n" + '"' + "$i" + '"' + "`r`n"
                }
                else
                {
                    $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=displayname eq '$i'"
                    $users  = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                    $includeUsers+= "`r`n" + '"' + "$users" + '"' + "`r`n"
                }
            }
            else 
            {
                $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=displayname eq '$i'"
                $users  = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $includeUsers+= "`r`n" + '"' + "$users" + '"' + ","
            }
        }
    }


# define  $excludeUsers
    $excludeUsers = $null
    if ($line.excludeUsers -eq "")
    {
        $excludeUsers = $null
    }
    elseif ($line.excludeUsers -eq "All")
    {
        $excludeUsers = "`r`n" + '"' + "$($line.excludeUsers)" + '"' + "`r`n"
    }
    elseif ($line.excludeUsers -eq "GuestsOrExternalUsers ")
    {
        $excludeUsers = "`r`n" + '"' + "$($line.excludeUsers)" + '"' + "`r`n"
    }
    elseif ((($line.excludeUsers).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeUsers).Split("`r"))
        {
            $users= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif($i -eq "GuestsOrExternalUsers" -and $count -ne ((($line.excludeUsers).Split("`r")).count-1))
            {
                $excludeUsers+= "`r`n" + '"' + "$i" + '"' + ","
            }
            elseif ( $count -eq ((($line.excludeUsers).Split("`r")).count-1))
            {
                if($i -eq "GuestsOrExternalUsers")
                {
                    $excludeUsers+= "`r`n" + '"' + "$i" + '"' + "`r`n"
                }
                else
                {
                    $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=displayname eq '$i'"
                    $users  = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                    $excludeUsers+= "`r`n" + '"' + "$users" + '"' + "`r`n"
                }
            }
            else 
            {
                
                $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=displayname eq '$i'"
                $users  = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $excludeUsers+= "`r`n" + '"' + "$users" + '"' + ","
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
        $includeGroups = "`r`n" + '"' + "$($line.includeGroups)" + '"' + "`r`n"
    }
    elseif ($line.includeGroups -eq "GuestsOrExternalusers")
    {
        $includeGroups = "`r`n" + '"' + "$($line.includeGroups)" + '"' + "`r`n"
    }
    elseif ((($line.includeGroups).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeGroups).Split("`r"))
        {
            $groups= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif($count -eq ((($line.includeGroups).Split("`r")).count-1))
            {
                $groups = ($allgroups | Where-Object { $_.displayname -eq "$i"}).id
                #$apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=displayname eq '$i'"
                #$groups  = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $includeGroups+= "`r`n" + '"' + "$groups" + '"' + "`r`n"
            }
            else 
            {
                $groups = ($allgroups | Where-Object { $_.displayname -eq "$i"}).id
                #$apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=displayname eq '$i'"
                #$groups  = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $includeGroups+= "`r`n" + '"' + "$groups" + '"' + ","
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
        $excludeGroups = "`r`n" + '"' + "$($line.excludeGroups)" + '"' + "`r`n"
    }
    elseif ($line.excludeGroups -eq "GuestsOrExternalusers")
    {
        $excludeGroups = "`r`n" + '"' + "$($line.excludeGroups)" + '"' + "`r`n"
    }
    elseif ((($line.excludeGroups).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeGroups).Split("`r"))
        {
            $groups= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif($count -eq ((($line.excludeGroups).Split("`r")).count-1))
            {
                $groups = ($allgroups | Where-Object { $_.displayname -eq "$i"}).id
                #$apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=displayname eq '$i'"
                #$groups  = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $excludeGroups+= "`r`n" + '"' + "$groups" + '"' + "`r`n"
            }
            else 
            {
                $groups = ($allgroups | Where-Object { $_.displayname -eq "$i"}).id
                #$apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=displayname eq '$i'"
                #$groups  = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).id
                $excludeGroups+= "`r`n" + '"' + "$groups" + '"' + ","
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
        $includeRoles = "`r`n" + '"' + "$($line.includeroles)" + '"' + "`r`n"
    }
    elseif ((($line.includeroles).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.includeroles).Split("`r"))
        {
            $roles= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif ( $count -eq ((($line.includeroles).Split("`r")).count-1))
            {
                $roles = ($allrole | Where-Object { $_.displayname -eq "$i"}).id
                $includeRoles+= "`r`n" + '"' + "$roles" + '"' + "`r`n"
            }
            else 
            {
                $roles = ($allrole | Where-Object { $_.displayname -eq "$i"}).id
                $includeRoles+= "`r`n" + '"' + "$roles" + '"' + ","
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
        $excludeRoles = "`r`n" + '"' + "$($line.excludeRoles)" + '"' + "`r`n"
    }
    elseif ((($line.excludeRoles).Split("`r")).count -ge 0)
    {
        $count = 0
        foreach ($i in ($line.excludeRoles).Split("`r"))
        {
            $roles= $null
            $count++
            if($i -eq "")
            {
                #do nothing last empty 
            }
            elseif ($count -eq (((($line.excludeRoles).Split("`r")).count)-1))
            {
                $roles = ($allrole | Where-Object { $_.displayname -eq "$i"}).id
                $excludeRoles+= "`r`n" + '"' + "$roles" + '"' + "`r`n"
            }
            else 
            {
                $roles = ($allrole | Where-Object { $_.displayname -eq "$i"}).id
                $excludeRoles+= "`r`n" + '"' + "$roles" + '"' + ","
            }
        }
    }

    # define $sessionControls
    $sessionControls = $null
    if($($line.sessionControls) -eq "")
    {
        $sessionControls = "null"
    }
    else
    {
        $sessionControls = "$($line.sessionControls)"
    }

    #################################################################
    # define .json file 
    $ConditionalAccessPolicies = @"
{
    "displayName": $displayname,
    "state": $state,
    "grantControls": {
    "operator": $operator,
    "builtInControls": [$grantbuiltInControls],
    "customAuthenticationFactors": [$customAuthenticationFactors],
    "termsOfUse": [$termsOfUse]
    },
    "conditions": {
    "signInRiskLevels": [$signInRiskLevels],
    "clientAppTypes": [$clientAppTypes],
    "platforms": {
      "includePlatforms": [$Includeplatforms],
      "excludePlatforms": [$excludeplatforms]
    },
    "locations": {
      "includeLocations": [$includeLocations],
      "excludeLocations": [$excludeLocations]
    },
    "devices": {
      "includeDeviceStates": [$includeDeviceStates],
      "excludeDeviceStates": [$excludeDeviceStates]
    },
    "applications": {
      "includeApplications": [$includeApplications],
      "excludeApplications": [$excludeApplications],
      "includeUserActions": [$includeUserActions]
    },
    "users": {
      "includeUsers": [$includeUsers],
      "excludeUsers": [$excludeUsers],
      "includeGroups": [$includeGroups],
      "excludeGroups": [$excludeGroups],
      "includeRoles": [$includeRoles],
      "excludeRoles": [$excludeRoles]
    }
  },
  "sessionControls": $sessionControls
}
"@
$json = $PSScriptRoot + ".\json\" + "$($line.id)" + ".json"

Add-Content -Value $ConditionalAccessPolicies -PassThru $json
#pause
}