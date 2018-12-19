param
([object]$WebhookData)
$VerbosePreference = 'continue'

#region Verify if Runbook is started from Webhook.

# If runbook was called from Webhook, WebhookData will not be null.
if ($WebHookData){

    # Collect properties of WebhookData
    $WebhookName     =     $WebHookData.WebhookName
    $WebhookHeaders  =     $WebHookData.RequestHeader
    $WebhookBody     =     $WebHookData.RequestBody

    # Collect individual headers. Input converted from JSON.
    $From = $WebhookHeaders.From
    $Input = (ConvertFrom-Json -InputObject $WebhookBody)
    Write-Verbose "WebhookBody: $Input"
    Write-Output -InputObject ('Runbook started from webhook {0} by {1}.' -f $WebhookName, $From)
}
else
{
   Write-Error -Message 'Runbook was not started from Webhook' -ErrorAction stop
}
#endregion

$employment = $Input.EmploymentStatus
Write-Output $employment

#Check Employment status
if($employment -eq "New"){

    #Updating SharePoint list item status
    $SPListItemID = $Input.ListItemID

    $spoconn = Connect-PnPOnline –Url https://tenant.sharepoint.com/sites/site –Credentials (Get-AutomationPSCredential -Name 'AzureAdmin') -ReturnConnection -Verbose

    $itemupdate = Set-PnPListItem -List "Employee Information" -Identity $SPListItemID -Values @{"Status" = "In Progress"} -Connection $spoconn

    #Local AD OU
    $Path = "OU=Employees,OU=OU,DC=domain,DC=net"

    $TemPass = "TempPass" + "$SPListItemID"
    Write-Output $TemPass

    $ADsplat = @{
        SamAccountName = $Input.FirstName
        UserPrincipalName = "$($Input.FirstName)`@domain.net"
        DisplayName = "$($Input.FirstName) $($Input.LastName)"
        Name = "$($Input.FirstName) $($Input.LastName)"
        GivenName = $Input.FirstName
        SurName = $Input.LastName
        ChangePasswordAtLogon = $true
        Description = "Created by Azure Automation"
        MobilePhone = $Input.Mobilephone
        AccountPassword = ConvertTo-SecureString $TemPass -AsPlainText -Force
        Department = $Input.Department
        Enabled = $true
        Path = $Path
    }


    $SamAccountName = $ADsplat.SamAccountName
    $upn = $ADsplat.UserPrincipalName

    #Creating account in local AD
    #Requires delegated permission in AD, do not need to be Domain Admin
    New-ADUser @ADsplat

    #Updating List Item Title
    $itemupdate = Set-PnPListItem -List "Employee Information" -Identity $SPListItemID -Values @{"Title" = $upn} -Connection $spoconn

    #Enable Mailbox
    #Requires delegated exchange permissions
    asnp *exchange*
    Enable-Mailbox -Identity $SamAccountName

    #Group memmbership in local AD
    Add-AdGroupMember -Identity "All Users" -Members $SamAccountName
    sleep 90

    #Start AAD Connect Sync
    #If AAD server is different than hybrid worker, user remove powershell to invoke synch 
    Start-ADSyncSyncCycle -PolicyType Delta

    #Assigning Licenses
    Write-Output "Wait Azure AD Sync"
    sleep 100
    
    #Make sure relevant Powershell modules for connecting to Office 365 is installed on the hybrid worker
    Connect-MsolService -Credential (Get-AutomationPSCredential -Name 'AzureAdmin')
    Set-MsolUser -UserPrincipalName $upn -UsageLocation NO
    Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "jh365dev:DEVELOPERPACK" 
    #Remove SWAY from User license
    $LO = New-MsolLicenseOptions -AccountSkuId "jh365dev:DEVELOPERPACK" -DisabledPlans "SWAY"
    Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $LO

    #Update list item
    $itemupdate = Set-PnPListItem -List "Employee Information" -Identity $SPListItemID -Values @{"Status" = "Completed"; "EmploymentStatus" = "Current"} -Connection $spoconn
}

elseIf($employment -eq "Terminated"){
    $user = Get-ADUser -Identity $Input.FirstName
    Disable-ADAccount -Identity $user
    Move-ADObject -Identity $user.ObjectGUID -TargetPath "OU=Terminated,OU=OU,DC=domain,DC=net"
    Write-Output "User Account Disabled"

}

else{
    Write-Output "Nothing to do..moving on..."
}

