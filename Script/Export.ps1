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
            $includeApplications+= $_ +";"
        }
        elseif ($_ -eq "Office365")
        {
            $includeApplications+= $_ +";"
        }
        elseif ($_ -eq "none")
        {
            $includeApplications+= $_ +";"
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=appid eq '$_'"
            $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
            $includeApplications+= ($Data).displayName + ";"        
        }
    }
    $output | Add-Member NoteProperty -Name "includeApplications" -value  "$includeApplications"

    # get exclude App displayname 
    $($conditionalaccess.conditions.applications.excludeApplications) | % {
        if ($_ -eq "ALL")
        {
            $excludeApplications+= $_ +";"
        }
        elseif ($_ -eq "Office365")
        {
            $excludeApplications+= $_ +";"
        }
        else
        {
            #$ExcludeGroupId = (Get-MsGraph -AccessToken $AccessToken -Uri $Uri | Where-Object { $_.displayName -eq $ExcludeGroup }).id
            $apiUrl = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=appid eq '$_'"
            $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
            $excludeApplications+= ($Data).displayName + ";"        
        }
    }
    $output | Add-Member NoteProperty -Name "excludeApplications" -value  "$excludeApplications"
    $output | Add-Member NoteProperty -Name "includeUserActions" -value "$($conditionalaccess.conditions.applications.includeUserActions)"
    $output | Add-Member NoteProperty -Name "signInRiskLevels" -value "$($conditionalaccess.conditions.signInRiskLevels)"
    $output | Add-Member NoteProperty -Name "clientAppTypes" -value "$($conditionalaccess.conditions.clientAppTypes)"
    $output | Add-Member NoteProperty -Name "deviceStates" -value "$($conditionalaccess.conditions.deviceStates)"
    
    # get Include  user displayname 
    $($conditionalaccess.conditions.users.includeusers) | % {
        if ($_ -eq "ALL")
        {
            $userInclude+= $_ +";"
        }
        elseif ($_ -eq "GuestsOrExternalUsers")
        {
            $userInclude+= $_ +";"
        }
        else
        {
            $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$_'"
            $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
            $userInclude+= ($Data).displayName + ";" 
        
        }
    }
    $output | Add-Member NoteProperty -Name "Includeusers" -value  "$userinclude"

    # get Exclude user displayname 

    $($conditionalaccess.conditions.users.excludeusers) | % {
        if ($_ -eq "ALL")
        {
            $userexclude+= $_ +";"
        }
        elseif ($_ -eq "GuestsOrExternalUsers")
        {
            $userexclude+= $_ +";"
        }
        else
        {    
            $apiUrl = "https://graph.microsoft.com/v1.0/users?`$filter=id eq '$_'"
            $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
            $userexclude+= ($Data).displayName + ";"
        }
     }
     $output | Add-Member NoteProperty -Name "excludeUsers" -value "$userexclude"


    #get Include group displayname  
    $($conditionalaccess.conditions.users.includeGroups) | % {
    $apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=id eq '$_'"
    $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
    $groupInclude+= ($Data).displayName + ";" }
    $output | Add-Member NoteProperty -Name "includeGroups" -value "$groupInclude"

    #get Exclude group displayname
    $($conditionalaccess.conditions.users.excludeGroups) | % {
    $apiUrl = "https://graph.microsoft.com/v1.0/groups?`$filter=id eq '$_'"
    $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
    $groupexclude+= ($Data).displayName + ";" }
    $output | Add-Member NoteProperty -Name "excludeGroups" -value "$groupexclude"

    #get Include groupManagement  displayname
    $($conditionalaccess.conditions.users.includeRoles) | % {
    $apiUrl = "https://graph.microsoft.com/v1.0/directoryRoleTemplates?`$filter=id eq '$_'"
    $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
    $IncludeRole+= ($Data).displayName + ";" }
    $output | Add-Member NoteProperty -Name "includeRoles" -value "$IncludeRole"
    
    #get exclude groupManagement  displayname
    $($conditionalaccess.conditions.users.excludeRoles) | % {
    $apiUrl = "https://graph.microsoft.com/v1.0/directoryRoleTemplates?`$filter=id eq '$_'"
    $Data = Get-MsGraph -AccessToken $AccessToken -Uri $ApiUrl
    $excludeRole+= ($Data).displayName + ";" }
    $output | Add-Member NoteProperty -Name "excludeRoles" -value "$excludeRole"

    $output | Add-Member NoteProperty -Name "Includeplatforms" -value "$($conditionalaccess.conditions.platforms.includePlatforms)"
    $output | Add-Member NoteProperty -Name "excludeplatforms" -value "$($conditionalaccess.conditions.platforms.excludePlatforms)"
    $output | Add-Member NoteProperty -Name "includeLocations" -value "$($conditionalaccess.conditions.locations.includeLocations)"
    $output | Add-Member NoteProperty -Name "excludeLocations" -value "$($conditionalaccess.conditions.locations.excludeLocations)"
    $output | Add-Member NoteProperty -Name "includeDeviceStates" -value "$($conditionalaccess.conditions.devices.includeDeviceStates)"
    $output | Add-Member NoteProperty -Name "excludeDeviceStates" -value "$($conditionalaccess.conditions.devices.excludeDeviceStates)"
    # part get grant Controls 
    $output | Add-Member NoteProperty -Name "grantoperator" -value "$($conditionalaccess.grantControls.operator)"
    $output | Add-Member NoteProperty -Name "grantbuiltInControls" -value "$($conditionalaccess.grantControls.builtInControls)"
    $output | Add-Member NoteProperty -Name "customAuthenticationFactors" -value "$($conditionalaccess.grantControls.customAuthenticationFactors)"
    $output | Add-Member NoteProperty -Name "termsOfUse" -value "$($conditionalaccess.grantControls.termsOfUse)"

    $allconditionalaccess+= $Output

$($conditionalaccess.displayName)
}

$allconditionalaccess | Export-Csv -Path $exportconditionalaccess -Encoding UTF8 -NoTypeInformation 


