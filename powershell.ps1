$Version = $Host | select version
if ($Version.Version.Major -gt 1) {$Host.Runspace.ThreadOptions = 'ReuseThread'}

if ((Get-PSSnapin 'Microsoft.SharePoint.PowerShell' -ErrorAction SilentlyContinue) -eq $null)
{
   Write-Prgress -Activity 'Loading Modules' -Status 'Loading Microsoft.SharePoint.PowerShell'
   Add-PSSnapin Microsoft.SharePOint.PowerShell
}

$ErrorActionPreference = 'SilentlyContinue'

#Get_Set_Web_App_Policy -Option "Get" -Path "C:\DATA\CPS\bharath\Logs\"

#Get_Set_Web_App_Policy -Option "Set" -Path "C:\DATA\CPS\bharath\Logs\" -InputCSV "C:\DATA\CPS\bharath\InputFiles\SPWebAppPolicyReport.csv"

function Get_Set_Web_App_Policy() {
   
    Param 
    ( 
        [Parameter(Mandatory=$false)] 
        [Alias('GetSet')] 
        [ValidateSet("Get","Set")] 
        [string]$Option, 
         
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path,

        [Parameter(Mandatory=$false)] 
        [Alias('InputForUpdatePolicy')] 
        [string]$InputCSV
    ) 

        try
        {

            $sScriptVersion = '1.0'
            
   
            if ($Option -eq "Get")
            {
               
                $sLogPath = $Path #'C:\DATA\CPS\bharath\Logs\'

                $sLogName = "LogFile GetWebAppPolicy Report File $(Get-Date -Format "MMddyyyy_HHmmss").log"
                $sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

                Write-Log -Message $sScriptVersion -Path $sLogFile
                Write-Log -Message "Start Get Web App Policy" -Path $sLogFile

                $sFileName = "SPWebAppPolicyReport $(Get-Date -Format 'MMddyyyy_HHmmss').csv"
                $sFilePath = Join-Path -Path $sLogPath -ChildPath $sFileName

                GetWebAppPolicy $sFilePath $sLogFile

                Write-Log -Message "End Get Web App Policy" -Path $sLogFile
            }
            if ($Option -eq "Set")
            {
              
               $sLogPath = $Path #'C:\DATA\CPS\bharath\Logs\'
               $sLogName = "LogFile SetWebAppPolicy Report File $(Get-Date -Format "MMddyyyy_HHmmss").log"
               $sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

               Write-Log -Message $sScriptVersion -Path $sLogFile
               Write-Log -Message "Start Set Web App Policy" -Path $sLogFile

               SetWebAppPolicy $InputCSV $sLogFile

               Write-Log -Message "End Set Web App Policy" -Path $sLogFile
            }
   
         }
        catch
        {
            Write-Log -Message $_.Exception.Message -Path $sLogFile -Level Error 
        }
     
}

function SetWebAppPolicy([string] $sInputCSV, [string] $sLogFile)
{
    try
    {
        $csv = Import-Csv $sInputCSV

        $webAppsfromcsv = $csv | select "Web Application Url" -Unique

        #Write-Log -Message $webAppsfromcsv -Path $sLogFile

        foreach($webAppfromcsv  in $webAppsfromcsv)
        {
           #$webAppfromcsv

        Write-Host "Updating Policy for $($webAppfromcsv.'Web Application Url')"
        Write-Log -Message "$($webAppfromcsv.'Web Application Url') *************************** Start ***********************************" -Path $sLogFile
        Write-Log -Message "Updating Policy for $($webAppfromcsv.'Web Application Url')" -Path $sLogFile

        $webApp = Get-SPWebApplication $($webAppfromcsv.'Web Application Url')

        if($webApp)
        { 
           Write-Host "Removing all policies for $(webApp.Url)"

           Write-Log -Message "Removing all policies for $(webApp.Url)" -Path $sLogFile

           $UserNames = $webApp.Policies | select UserName
             foreach($UserName in $UserNames)
             {
                $webApp.Policies.Remove($UserName.UserName)
                $WebApp.Update()
             }
               $policiesfromCSV = $csv | ?{$_."Web Application Url" -eq $webApp.Url -and $_.'Is Claims' -eq $true}
               if(($policiesfromCSV | Measure-Object).Count -gt 0)
               {
                  Write-Host "Adding  policies from csv for $(webApp.Url)"

                  Write-Log -Message "Adding  policies from csv for $(webApp.Url)" -Path $sLogFile

                    foreach($policiefromCSV in $policiesfromCSV)
                    {
                         $identifier = $policiefromCSV.Identifier
                         $DispalyName = $policiefromCSV.'Display Name'

                         if (! $displayname)
                         { 
                         $web = Get-SpWeb $webApp.Url
                         $displayname = $web.EnsureUser($identifier).Displayname
                         $web.Dispose()
                         }

                       $type = $policiefromCSV.Type
                       $role = $policiefromCSV.'Role Name'
                       [bool]$isSystemUser = [System.Convert]::ToBoolean($policiefromCSv.'Is System User')
                       Switch($role)
                         {
                             "None" {$role = "None"}
                             "Deny All" {$role = "Deny All"}
                             "Deny Write" {$role = "DenyWrite"}
                             "Full Read" {$role = "FullRead"}
                             "Full Control" {$role = "FullControl"}
                              default {$role = "None"}
                          }

                      Write-Host " Adding User : $displayName, with permission: $role"

                      Write-Log -Message " Adding User : $displayName, with permission: $role" -Path $sLogFile

                      SetPolicyForWebAppPermission $webApp $identifier $displayname $type $role $isSystemUser $sLogFile
                  }
        }
        }
        Write-Log -Message "$($webAppfromcsv.'Web Application Url') *************************** End ***********************************" -Path $sLogFile
        }
    }
    catch
    {
        Write-Log -Message $_.Exception.Message -Path $sLogFile -Level Error 
    }

}


function GetWebAppPolicy([String] $sFilePath, [String] $sLogFile)
{
    try
    {       Write-Host "Get All Web Applications"
            Write-Log -Message "Get All Web Applications" -Path $sLogFile

            $was = Get-SPWebApplication #http
            $objs = @()
                foreach($wy in $was) {
                     
                    Write-Log -Message "$wy ********************************** Start *********************************" -Path $sLogFile
                    $obj = GetAllWebPolicy $wy $sLogFile
                    $objs +=obj
                    Write-Log -Message "$wy ********************************** End *********************************" -Path $sLogFile
                }

            $objs | Export -Csv $sFilePath -NoTypeInformation
            
            Write-Host "Get All Web Applications Exported to CSV"
            Write-Log -Message "Get All Web Applications Exported to CSV" -Path $sLogFile
    }
    catch
    {
        Write-Log -Message $_.Exception.Message -Path $sLogFile -Level Error 
    }
}


function GetAllWebAppPolicy([Microsoft.SharePoint.Administration.SPWebApplication] $webApp, [String] $sLogFile)
{
        try
        {
               $reportObjs =@()
               $SwebAppUrl = $webApp.Url
               $ClaimMgr = Get-SpClaimProviderManager
	
                   Write-Host "Generating web aplication policy report" $webApp.url 
                   Write-Log -Message "Generating web aplication policy report" $webApp.url -Path $sLogFile

                   foreach($policy in $webApp.policies)
		            {

                       $UserInfo = ""

                     try {

	                    foreach($role in $policy.PolicyRoleBindings)
	                    {	
                                    $userType =""
                                    $userIdentifier = ""
                                    $isclaims =""
                                    if(($policy.username).indexof("|") -gt 0) 
                                    {
                                    $userIdentifier =$claimMgr.ConvertClaimToIdentifier($policy.username)
                                    $isClaim = $true
                                    $claimType= ($policy.username).chars(3)
                                    if($claimType -eq "#") {
                                    $userType ="User"}
                                    elseif($claimType -eq "+"){
                                    $userType ="Group"}
                                    }
                                    else{
                                    $userType =""
                                    $userIdentifier = $policy.username
                                     $isClaim = $false
                                      }
                                    $SPWebAppPolicyObj = New-Object -TypeName PSObject
                                    $SPWebAppPolicyObj | Add-Member -Name "Web Application Url" -MemberType Noteproperty -Value $sWebAppUrl
                                    $SPWebAppPolicyObj | Add-Member -Name "Web Application Name" -MemberType Noteproperty -Value $WebApp.Name
                                    $SPWebAppPolicyObj | Add-Member -Name "User Name" -MemberType Notproperty -Value $sUserName
	                                $SPWebAppPolicyObj | Add-Member -Name "Role Name" -MemberType Noteproperty -Value $role.name
                                    $SPWebAppPolicyObj | Add-Member -Name "Display Name" -MemberType Notproperty -Value $policy.Displayname
	                                $SPWebAppPolicyObj | Add-Member -Name "Is System User" -MemberType Noteproperty -Value $role.namepolicy.IsSystemUser
                                   
                                    $SPWebAppPolicyObj | Add-Member -Name "Principal Type" -MemberType Noteproperty -Value $userType
                                    $SPWebAppPolicyObj | Add-Member -Name "Identifier" -MemberType Noteproperty -Value $userIdentifier
                                    $SPWebAppPolicyObj | Add-Member -Name "Is Claims" -MemberType Noteproperty -Value $isClaim
                                    $reportObjs += $SPWebAppPolicyobj

                                    $UserInfo = "Web Application Url: $sWebAppUrl ; Web Application Name: $WebApp.Name ; User Name: $userIdentifier ;"
                                    
                                    Write-Log -Message "User Info: $UserInfo" -Path $sLogFile
			                    }
		                    }
                            catch
                            {
                                   Write-Log -Message "Error at User Info: $UserInfo" -Path $sLogFile -Level WARNING
                                   Write-Log -Message $_.Exception.Message -Path $sLogFile -Level Error      
                            }

	           }	
	            return $reportObjs
        }
        catch
        {
            Write-Log -Message $_.Exception.Message -Path $sLogFile -Level Error 
        }
}

function SetPolicyForWebPermission([Microsoft.SharePoint.Administration.SPWebApllication] $webApp,[String] $principal,[String] $principalDisplayName,[String] $principalType,[String] $permissionlevel,[bool] $isSysytemUser=$false, [String] $sLogFile)
{
        try
        {
            #ToDo: Add validaation for permission level parameter
            #$webapp =Get-SpWebApplication $webAppUrl

            Write-Log -Message "Start Set Policy for $webApp $principal $principalDisplayName $principalType $permissionlevel" -Path $sLogFile

            if($WebApp.UserClaimsAuthentication)
            {
               Switch($principalType)
              {
                "User" { $principalClaims = New-SPClaimsPrincipal - identityType WindowsSamAccountName }
                "Group" { $principalClaims = New-SPClaimsPrincipal - identityType WindowsSecurityGroupName }
               default { $principalClaims = New-SPClaimsPrincipal - identityType WindowsSamAccountName }
               }
                $principalClaims = $principalClaims.ToEncodedString()
                $policy= $Webapp.policies.Add($principalClaims, $principalDisplayName)
            }
            else
            {
                $ppolicy = $Webapp.policies.Add($principalClaims, $principalDisplayName)
            }
              switch ($permissionlevel)
            {
                 "None"{$policyRole =$webApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::None)}
                 "DenyAll"{$policyRole =$webApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::DenyAll)}
                 "DenyWrite"{$policyRole =$webApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::DenyWrite)}
                 "FullRead"{$policyRole =$webApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullRead)}
                 "FullControl"{$policyRole =$webApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::FullControl)}
                 default{$policyRole =$webApp.PolicyRoles.GetSpecialRole([Microsoft.SharePoint.Administration.SPPolicyRoleType]::None)}
            } 

            # Write host "Adding web application policy role binding for $accountDisplayName with access level $permissionlevel for web app $webAppUrl
            $policy.PolicyRoleBindings.Add($policyRole)
            $policy.IsSystemUser = $isSystemUser
            $webApp.update()

            Write-Log -Message "End Set Policy for $webApp $principal $principalDisplayName $principalType $permissionlevel" -Path $sLogFile
        }
        catch
        {
            Write-Log -Message $_.Exception.Message -Path $sLogFile -Level Error 
        }

}

function Write-Log 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path, 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info" 
        
    ) 
        
        
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                $LevelText = 'INFO:' 
                } 
            } 
         
         if ($Level -eq 'Error')
         { 
            Write-Host "Error Occuerd, Please look into log : $Message" -foreground "red"
         }

        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
}

