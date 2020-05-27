$ModuleGraph = "C:\Users\GBATTEUX\Desktop\ConnectO365Services\ConnectO365Services.ps1"
& $ModuleGraph

$AccessToken = Connect-toGraph 


################################################################################
#################### Part for export 

$exportconditionalaccess = $PSScriptRoot + ".\Condtionalaccess.csv"


$apiUrl = 'https://graph.microsoft.com/beta/conditionalAccess/policies'
$conditionalaccess = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl



$allconditionalaccess = @()
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
    $ExcludeRole =$null
    $deviceStates = $null
    $clientAppTypes = $null
    $signInRiskLevels = $null
    $includeUserActions = $null
    $Includeplatforms = $null
    $excludeplatforms = $null 
    $includeLocations = $null
    $excludeLocations = $null
    $includeDeviceStates = $null
    $excludeDeviceStates = $null
    $grantbuiltInControls = $null


    ##################### part get global information 
    $output | Add-Member NoteProperty -Name "id" -Value "$($conditionalaccess.id)"
    $output | Add-Member NoteProperty -Name "displayname" -value "$($conditionalaccess.displayName)"
    $output | Add-Member NoteProperty -Name "createdatetime" -value "$($conditionalaccess.createdDateTime)"
    $output | Add-Member NoteProperty -Name "modifiedDateTime" -value "$($conditionalaccess.modifiedDateTime)"
    $output | Add-Member NoteProperty -Name "state" -value "$($conditionalaccess.state)"
    $output | Add-Member NoteProperty -Name "sessionControls" -value "$($conditionalaccess.sessionControls)"
    
    #################### part get Conditions
    
    # get Include App displayname 
    $($conditionalaccess.conditions.applications.includeApplications) | % {
        if ($_ -eq "ALL")
        {
            $includeApplications+= $_ 
        }
        elseif ($_ -eq "Office365")
        {
            $includeApplications+=  $_ +"`r" 
        }
        elseif ($_ -eq "none")
        {
            $includeApplications+=  $_ 
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=appid eq '$_'"
            $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
            $includeApplications+=  ($Data).displayName  + "`r"       
        }
    }
    $output | Add-Member NoteProperty -Name "includeApplications" -value  "$includeApplications"

    # get exclude App displayname 
    $($conditionalaccess.conditions.applications.excludeApplications) | % {
        if ($_ -eq "ALL")
        {
            $excludeApplications+=  $_ 
        }
        elseif ($_ -eq "Office365")
        {
            $excludeApplications+= $_ + "`r" 
        }
        elseif ($_ -eq "none")
        {
            $excludeApplications+=  $_ 
        }
        else
        {
            #$ExcludeGroupId = (Get-MsGraph -AccessToken $AccessToken -Uri $Uri | Where-Object { $_.displayName -eq $ExcludeGroup }).id
            $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=appid eq '$_'"
            $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
            $excludeApplications+=  ($Data).displayName + "`r"     
        }
    }
    $output | Add-Member NoteProperty -Name "excludeApplications" -value  "$excludeApplications"
    
    #get includeUserActions
    $($conditionalaccess.conditions.applications.includeUserActions) | % {
        $includeUserActions+= "$_" + "`r"
    }
    $output | Add-Member NoteProperty -Name "includeUserActions" -value "$includeUserActions"

    #Get signInRiskLevels
    $($conditionalaccess.conditions.signInRiskLevels) | % {
        $signInRiskLevels+= "$_" + "`r"
    }
    $output | Add-Member NoteProperty -Name "signInRiskLevels" -value "$signInRiskLevels"

    #get clientAppTypes
    $($conditionalaccess.conditions.clientAppTypes) | % {
        $clientAppTypes+= "$_" + "`r" 
    }
    $output | Add-Member NoteProperty -Name "clientAppTypes" -value "$clientAppTypes"

    #get deviceStates
    $($conditionalaccess.conditions.deviceStates) | % {
        $deviceStates+= "$_" + "`r" 
    }
    $output | Add-Member NoteProperty -Name "deviceStates" -value "$deviceStates"
    
    # get Include  user displayname 
    $($conditionalaccess.conditions.users.includeusers) | % {
        if ($_ -eq "ALL")
        {
            $userInclude+=  $_ 
        }
        elseif ($_ -eq "GuestsOrExternalUsers")
        {
            $userInclude+=  $_ + "`r" 
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$_'"
            $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
            $userInclude+= ($Data).displayName + "`r" 
        
        }
    }
    $output | Add-Member NoteProperty -Name "Includeusers" -value  "$userinclude"

    # get Exclude user displayname 

    $($conditionalaccess.conditions.users.excludeusers) | % {
        if ($_ -eq "ALL")
        {
            $userexclude+= $_ 
        }
        elseif ($_ -eq "GuestsOrExternalUsers")
        {
            $userexclude+=  $_ + "`r" 
        }
        else
        {    
            $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$_'"
            $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
            $userexclude+=  ($Data).displayName + "`r" 
        }
     }
     $output | Add-Member NoteProperty -Name "excludeUsers" -value "$userexclude"


    #get Include group displayname  
    $($conditionalaccess.conditions.users.includeGroups) | % {
        $apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=id eq '$_'"
        $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
        $groupInclude+= ($Data).displayName + "`r" 
    }
    $output | Add-Member NoteProperty -Name "includeGroups" -value "$groupInclude"

    #get Exclude group displayname
    $($conditionalaccess.conditions.users.excludeGroups) | % {
        $apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=id eq '$_'"
        $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
        $groupexclude+= ($Data).displayName + "`r"  
    }
    $output | Add-Member NoteProperty -Name "excludeGroups" -value "$groupexclude"

    #get Include groupManagement  displayname
    $($conditionalaccess.conditions.users.includeRoles) | % {
        $apiUrl = "https://graph.microsoft.com/v1.0/directoryRoleTemplates?`$filter=id eq '$_'"
        $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
        $IncludeRole+= ($Data).displayName + "`r" 
    }
    $output | Add-Member NoteProperty -Name "includeRoles" -value "$IncludeRole"
    
    #get exclude Roles  displayname
    $($conditionalaccess.conditions.users.excludeRoles) | % {
        $apiUrl = "https://graph.microsoft.com/v1.0/directoryRoleTemplates?`$filter=id eq '$_'"
        $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
        $excludeRole+= ($Data).displayName + "`r"
    }
    $output | Add-Member NoteProperty -Name "excludeRoles" -value "$excludeRole"

    # get includePlatforms
    $($conditionalaccess.conditions.platforms.includePlatforms) | % {
        $Includeplatforms+= "$_" + "`r" 
    }
    $output | Add-Member NoteProperty -Name "Includeplatforms" -value "$Includeplatforms"


    # get ExcludePlatforms
    $($conditionalaccess.conditions.platforms.excludePlatforms) | % {
        $excludeplatforms+= "$_" + "`r" 
    }
    $output | Add-Member NoteProperty -Name "excludeplatforms" -value "$excludeplatforms"


    # get includeLocations
    $($conditionalaccess.conditions.locations.includeLocations) | % {

        if ($_ -eq "ALL")
        {
            $includeLocations+= $_ 
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/beta/conditionalaccess/namedlocations?`$filter=id eq '$_'"
            $Locations = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).displayname
            $includeLocations+= "$locations" + "`r"
        }
    }
    $output | Add-Member NoteProperty -Name "includeLocations" -value "$includeLocations"
    
    # get excludeLocations
    $($conditionalaccess.conditions.locations.excludeLocations) | % {
        if ($_ -eq "ALL")
        {
            $excludeLocations+= $_ 
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/beta/conditionalaccess/namedlocations?`$filter=id eq '$_'"
            $Locations = (Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl).displayname
            $excludeLocations+= "$locations" + "`r"  
        }
    }
    $output | Add-Member NoteProperty -Name "excludeLocations" -value "$excludeLocations"

    # get includeDeviceStates
    $($conditionalaccess.conditions.devices.includeDeviceStates) | % {
        $includeDeviceStates+= "$_" + "`r" 
    }
    $output | Add-Member NoteProperty -Name "includeDeviceStates" -value "$includeDeviceStates"

    # get excludeDeviceStates
    $($conditionalaccess.conditions.devices.excludeDeviceStates) | % {
        $excludeDeviceStates+= "$_" + "`r" 
    }
    $output | Add-Member NoteProperty -Name "excludeDeviceStates" -value "$excludeDeviceStates"

    # part get grant Controls 
    $output | Add-Member NoteProperty -Name "grantoperator" -value "$($conditionalaccess.grantControls.operator)"

    #get grantControls
    $($conditionalaccess.grantControls.builtInControls) | % {
        $grantbuiltInControls+= "$_" + "`r"
    }
    $output | Add-Member NoteProperty -Name "grantbuiltInControls" -value "$grantbuiltInControls"

    $output | Add-Member NoteProperty -Name "customAuthenticationFactors" -value "$($conditionalaccess.grantControls.customAuthenticationFactors)"
    $output | Add-Member NoteProperty -Name "termsOfUse" -value "$($conditionalaccess.grantControls.termsOfUse)"

    $allconditionalaccess+= $Output

$($conditionalaccess.displayName)
#pause
}

$allconditionalaccess | Export-Csv -Path $exportconditionalaccess -NoTypeInformation -Encoding ASCII


