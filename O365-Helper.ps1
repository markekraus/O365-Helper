# OnPrem Active Director Domain
$ADDomainName = "contoso.com"
# Best Practice is to run all commands against a single DC so resultes can be verified more accurately. 
# This does not effect the PS session, which can use a different server if psremoting is not enabled on the Domain Controllers
# The server in $OnPremADPSURL must be able to reach the server in $ADDomainController
$ADDomainController = "dc01.contoso.com"
# Domain Admin account used for managing OnPrem and Office 365. Use the UPN and not SamAccount name. This help assumes they are the same for both.
$MGRUserName = "domain.admin@contoso.com"
# Optional Export-CliXml created Crednetial File. You can create a credntial file and use it to save from entering a user and password. If the file does not exist, you will be prompted for credentials.
# http://windowsitpro.com/development/save-password-securely-use-powershell
$CredFile = "C:\O365-Helper\O365Creds.xml"
# The progromatic SkuID in your office 365 tennat where the exhnage licenses reside
$MSOLSkuID = "contosos:ENTERPRISEPACK"
# The domain name for your Office 365 Tenant Mail
$O365TenantMailDomain = "contoso.mail.onmicorsoft.com"


# Powershell remoting session URL for Office 365:
$O365PSURL= "https://ps.outlook.com/powershell/"
# Powershell remoting session URL for OnPrem Hybrid Exchnage Server
$OnPremPSURL= "http://exch01.contoso.com/powershell/"
# Powershell remoting jumpbox for MietlAD
$OnPremADPSURL = "dc01.contoso.com"
# SharePoint Online Admin URL
$O365SPOADMINURL = "https://contoso-admin.sharepoint.com"


<#
.SYNOPSIS
Check if a sting is a valid email address
.DESCRIPTION
This checks a string to see if it would be a valid meail address. This does not
check wehther an email address actually exists or is usauble. It only determines
if it would be a valid email adddres or not. It retruns true or False.
.PARAMETER Email
A string to test for a valid email address
.EXAMPLE
if((Validate-Email -Email bob.testerton@contoso.com)){echo "this is a valid email address"}
else{echo "This is not a valid email address"}
#>
function Validate-Email{ 
    [CmdletBinding(SupportsShouldProcess=$true)]
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$Email
    ) 
    begin{}
    process{
        if($PSCmdlet.ShouldProcess($Email)){
            write-verbose "Validating $Email"
            return $Email -match "^(?("")("".+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$" 
        }
    }
}  

<#
.SYNOPSIS
Get crednetials for mail operations.
.DESCRIPTION
Get the username and password used to connect to OnPrem Mail and Office 365. 
The fucntion will use the global $MGRUserName variable to pre-fill the username if it is set.
.EXAMPLE
Get-GlobalMailCreds
#>
function Get-GlobalMailCreds {
    if(Test-Path $Global:CredFile){
        write-verbose "Getting credentials from $Global:CredFile"
        $Global:MailCreds = import-clixml -path $Global:CredFile
    }
    else{
        $Global:MailCreds = Get-Credential -UserName $Global:MGRUserName -Message "$Global:ADDomainName Domain Admin and Office 365 Account Password"
    }
}


<#
.SYNOPSIS
Start Office 365 Remoting Session
.DESCRIPTION
Connect to the Office 365 remoting session on the URL defined in the global $O365PSURL 
variable. This uses the global mail managment credentials and runs Get-GlobalMailCreds
if they are not set. The Office 365 remoting session is imported with the O365 prefix.
This will also importt he Microsoft Online module and connect to the MSOL Service
using the same credentials.
.EXAMPLE
Connect-O365
#>
function Connect-O365 {
    if (!$Global:MailCreds) { write-verbose "Getting credentials"; Get-GlobalMailCreds }
    Write-Verbose "Opening PSSession to $Global:O365PSURL"
    $o365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Global:O365PSURL -Credential $Global:MailCreds -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue -Name O365
    Write-Verbose "Importing PSSession o365Session with prefix O365"
    Import-PSSession $o365Session -Prefix O365 -AllowClobber | Out-Null
    Write-Verbose "Verifying MSOnline Module"
    if((Get-Module -ListAvailable -Name "MSOnline")){
        Write-Verbose "Importing MSOnline Module"
        Import-Module MSOnline
        Write-Verbose "Connecting to MSolService"
	    Connect-MsolService -Credential $Global:MailCreds 
    }
    else{
        Write-Warning @'
MSOnline PowerShell Module missing.
Please Install the Microsoft Online Services Sign-In Assistant from:
    http://www.microsoft.com/en-us/download/details.aspx?id=41950
Then install the Azure AD Module for Windows PowerShell:
    64-Bit: http://go.microsoft.com/fwlink/p/?linkid=236297
    32-Bit: http://go.microsoft.com/fwlink/p/?linkid=236298
The following commands will not work without this module:
    Disable-O365User
    Enable-O365User
'@
    }

}

function Connect-O365SharePoint{
    if (!$Global:MailCreds) { write-verbose "Getting credentials"; Get-GlobalMailCreds }
    $Global:SPOCreds = $Global:MailCreds
    Write-Verbose "Verifying SharePoint Online module"
    if((get-module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell")){
        Write-Verbose "Importing SharePoint Online module"
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
        Write-Verbose "Connect to SharePoint Online ad min URL $Global:O365SPOADMINURL"
        try{
            Write-Verbose "Trying with stored mail credentials"
            Connect-SPOService -Url $Global:O365SPOADMINURL -Credential $Global:SPOCreds -ErrorAction Stop
        }
        Catch{
            $Global:SPOCreds = Get-Credential -Message "SharePoint Online Global Admin Credentials:"
            Write-Verbose "Stored credentials failed. Prompting for SPO credentials."
            Connect-SPOService -Url $Global:O365SPOADMINURL -Credential $Global:SPOCreds
        }

    }
    else{
        Write-Warning @'
SharePoint Online Management PowerShell Module missing.
Please Install the SharePoint Online Management Shell:
    https://www.microsoft.com/en-us/download/details.aspx?id=35588
'@
    }
}

<#
.SYNOPSIS
Start OnPrem Mail Remoting Session
.DESCRIPTION
Connect to the OnPrem Mail remoting session on the URL defined in the global $OnPremPSURL 
variable. This uses the global mail managment credentials and runs Get-GlobalMailCreds
if they are not set. The OnPrem Mail remoting session is imported with the OnPrem prefix.
.EXAMPLE
Connect-OnPrem
#>
function Connect-OnPrem {
    if (!$Global:MailCreds) { Get-GlobalMailCreds }
    $OnPremMailSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Global:OnPremPSURL -Credential $Global:MailCreds -AllowRedirection -WarningAction SilentlyContinue -Name OnPremMail -Authentication Kerberos
    Import-PSSession $OnPremMailSession -Prefix OnPrem -AllowClobber | Out-Null
    
}

<#
.SYNOPSIS
Start OnPrem AD Remoting Session
.DESCRIPTION
Connect to the OnPrem AD domain remoting session on the FQDN defined in the global $OnPremADPSURL 
variable. This uses the global mail managment credentials and runs Get-GlobalMailCreds
if they are not set. The OnPremAD remoting session is imported with the OnPrem prefix.
.EXAMPLE
Connect-OnPrem
#>
function Connect-OnPremAD {
    if (!$Global:MailCreds) { Get-GlobalMailCreds }
    $OnPremADSession = new-pssession -computer $Global:OnPremADPSURL -Credential $Global:MailCreds -Name OnPremAD
    Invoke-Command -session $OnPremADSession -script { Import-Module ActiveDirectory | Out-Null } | Out-Null
    Import-PSSession -session $OnPremADSession -module ActiveDirectory -prefix OnPrem -AllowClobber | Out-Null
    Import-PSSession -session $OnPremADSession -CommandName Get-Acl -prefix OnPrem -AllowClobber | Out-Null
    Import-PSSession -session $OnPremADSession -CommandName Set-Acl -prefix OnPrem -AllowClobber | Out-Null
}

<#
.SYNOPSIS
End Office 365 Remoting Session
.DESCRIPTION
Remove to the Office 365 remoting session and remove the session imported commands
with the O365 prefix
.EXAMPLE
Disonnect-O365
#>
function Disconnect-O365 {
	Remove-PSSession -Name O365
    Remove-Module MSOnline
}	

<#
.SYNOPSIS
End OnPrem Mail Remoting Session
.DESCRIPTION
Remove to the OnPrem Mail remoting session
.EXAMPLE
Disonnect-OnPrem
#>
function Disconnect-OnPrem {
	Remove-PSSession -Name OnPremMail
}

<#
.SYNOPSIS
End OnPremAD Remoting Session
.DESCRIPTION
Remove to the OnPremAD remoting session 
.EXAMPLE
Disonnect-OnPremAD
#>
function Disconnect-OnPremAD {
	Remove-PSSession -Name OnPremAD
}

<#
.SYNOPSIS
Start All Mail Remoting Sessions
.DESCRIPTION
This runs Connect-O365 and Connect-OnPrem to establish all mail PowerShell
remoting sessions. This can be run as the first command before all others as it
will consequently prompt for crednetials and start all mail sessions.
.EXAMPLE
Connect-AllMail
#>
function Connect-AllMail {
    Connect-O365
    Connect-OnPrem
}

<#
.SYNOPSIS
End All Mail Remoting Sessions
.DESCRIPTION
This runs Disconnect-O365 and Disconnect-OnPrem to end all mail PowerShell
remoting sessions. This should be run before closing a PowerShell console
or ending a script as a cleanup measure.
.EXAMPLE
Disconnect-AllMail
#>
Function Disconnect-AllMail {
    Disconnect-O365
    Disconnect-OnPrem
}
    
<#
.SYNOPSIS
Disable an Offcie 365 User
.DESCRIPTION
Disables a user in Office 365, even if their AD account is active. This can also
be done to avoid waiting for the 30 minute ADFS sync to occur.
.PARAMETER UserPrincipalName
This is the OnPrem AD UserPrincipalName, not their primary email address or Office
365 Identity. If a person is contractor it will be <upn>@contractor.contoso.com.
.EXAMPLE
Disable-O365User -UserPrincipalName bob.testerton@contoso.com

Disable a regular employee.
.EXAMPLE
Disable-O365User -UserPrincipalName bob.testerton@contractor.contoso.com

Disable a contractor employee.
.EXAMPLE
Disable-O365User -UserPrincipalName $((Get-O365Recipient -Identity "Bob Testerton").WindowsLiveId)

If you don't know the UPN.
#>
function Disable-O365User {
	[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="High")]
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
		[string]$UserPrincipalName
	)
    begin{
        if(!(Get-Module -Name MSOnline)){Connect-O365}
    }
    process{
        if($PSCmdlet.ShouldProcess($UserPrincipalName)){
	        Set-MsolUser -UserPrincipalName $UserPrincipalName -blockcredential $true
        }
    }
}

<#
.SYNOPSIS
Enable an Offcie 365 User
.DESCRIPTION
Enables a user in Office 365. This does not override a disabled AD account. 
This is only usefule to re-enable a user disabled with Disable-O365User.
.PARAMETER UserPrincipalName
This is the OnPrem AD UserPrincipalName, not their primary email address or Office
365 Identity. If a person is contractor it will be *@contractor.contoso.com.
.EXAMPLE
Enable-O365User -UserPrincipalName bob.testerton@contoso.com

Enable a regular employee.
.EXAMPLE
Enable-O365User -UserPrincipalName bob.testerton@contractor.contoso.com

Enables a contractor employee.
.EXAMPLE
Enable-O365User -UserPrincipalName $((Get-O365Recipient -Identity "Bob Testerton").WindowsLiveId)

If you don't know the UPN for Bob Testerton.
#>
function Enable-O365User {
    [CmdletBinding(SupportsShouldProcess=$true)]
	param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
		[string]$UserPrincipalName
	)
    begin{
        if(!(Get-Module -Name MSOnline)){Connect-O365}
    }
    process{
        if($PSCmdlet.ShouldProcess($UserPrincipalName)){
	        Set-MsolUser -UserPrincipalName $UserPrincipalName -blockcredential $false
        }
    }
}

<#
.SYNOPSIS
Grant full permissions to a mailbox for a specific user
.DESCRIPTION
This will grant full permissions (with the exception of Send As) to mailbox
for a specific user. 
.PARAMETER Identity
A string containing an Office 365 mailbox Identity such as a displayname or email 
address. This is the mailbox that will be accessed.
.PARAMETER User
A string containing an Office 365 mailbox Identity such as a displayname or email
address. This is the User wich will be accessing the mailbox.
.PARAMETER Automap
This is a boolean to skip automapping mailboxes to Outlook clients. Use
-Automap:$false to disable automapping
The default is $true 
.EXAMPLE
Add-O365FullPermissions -Identity shared-mailbox@contoso.com -User bob.testerton@contoso.com

Grant bob.testerton@contoso.com full access to shared-mailbox@contoso.com.
.EXAMPLE
Add-O365FullPermissions -Identity shared-mailbox@contoso.com -User bob.testerton@contoso.com -Automap:$false

Grant bob.testerton@contoso.com full access to shared-mailbox@contoso.com and 
disable automapping.
.NOTES
See Add-O365SendAsPermissions for adding Send As permissions.
#>
function Add-O365FullPermissions {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
            [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$Identity,
            [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$User,
            [bool]$Automap=$true
    )
    begin{
        if(!$Identity -and !$User){
            write-error "Either Identity or User must be defined."
            continue
        }
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    Process{
        if($PSCmdlet.ShouldProcess($Identity,"Remove full permissons for $User")){
            add-O365mailboxPermission -Identity $Identity -user $User  -accessRights FullAccess -InheritanceType All -automapping $automap
        }
    }
}

<#
.SYNOPSIS
Remove full permissions to a mailbox for a specific user
.DESCRIPTION
This will remove full permissions (with the exception of SendAs) to mailbox
for a specific user. 
.PARAMETER Identity
A string containing an Office 365 mailbox Identity such as a displayname or email 
address. This is the mailbox that will no longer be accessed.
.PARAMETER User
A string containing an Office 365 mailbox Identity such as a displayname or email
address. This is the User wich will no longer be accessing the mailbox.
.EXAMPLE
Remove-O365FullPermissions -Identity shared-mailbox@contoso.com -User bob.testerton@contoso.com

Remove bob.testerton@contoso.com's full access to shared-mailbox@contoso.com.
.NOTES
See Remove-O365SendAsPermissions for removing Send As permissions.
#>
function Remove-O365FullPermissions {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="High")]
    param(
            [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$Identity,
            [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$User
    )
    begin{
        if(!$Identity -and !$User){
            write-error "Either Identity or User must be defined."
            continue
        }
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    Process{
        if($PSCmdlet.ShouldProcess($Identity,"Remove full permissons for $User")){
            Remove-O365mailboxPermission -Identity $Identity -user $User  -accessRights FullAccess -InheritanceType All -Confirm:$false
        }
    }
}

<#
.SYNOPSIS
Grant Send As permissions to a mailbox for a specific user
.DESCRIPTION
This will grant Send As permissions to mailbox for a specific user. 
.PARAMETER Identity
A string containing an Office 365 mailbox Identity such as a displayname or email 
address. This is the mailbox that will be used to send as.
.PARAMETER User
A string containing an Office 365 mailbox Identity such as a displayname or email
address. This is the User wich will send as the mailbox.
.EXAMPLE
Add-O365SendAsPermissions -Identity shared-mailbox@contoso.com -User bob.testerton@contoso.com

Grant bob.testerton@contoso.com rights to Send As to shared-mailbox@contoso.com.
.NOTES
See Add-O365FullPermissions for adding full permissions.
#>
function Add-O365SendAsPermissions {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
            [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$Identity,
            [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$User
    )
    begin{
        if(!$Identity -and !$User){
            write-error "Either Identity or User must be defined."
            continue
        }
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    Process{
        if($PSCmdlet.ShouldProcess($Identity,"Add ""Send As"" permissons for $User")){
            Add-O365RecipientPermission -Identity $Identity -Trustee $User -AccessRights SendAs -Confirm:$false
        }
    }
}

<#
.SYNOPSIS
Remove Send As permissions to a mailbox for a specific user
.DESCRIPTION
This will remove Send As permissions to mailbox for a specific user. 
.PARAMETER Identity
A string containing an Office 365 mailbox Identity such as a displayname or email 
address. This is the mailbox that will no longer be used to send as.
.PARAMETER User
A string containing an Office 365 mailbox Identity such as a displayname or email
address. This is the User wich will no longer send as the mailbox.
.EXAMPLE
Remove-O365SendAsPermissions -Identity shared-mailbox@contoso.com -User bob.testerton@contoso.com

Remove bob.testerton@contoso.com rights to Send As to shared-mailbox@contoso.com.
.NOTES
See Remove-O365FullPermissions for removing full permissions.
#>
function Remove-O365SendAsPermissions {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="High")]
    param(
            [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$Identity,
            [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$User
    )
    begin{
        if(!$Identity -and !$User){
            write-error "Either Identity or User must be defined."
            continue
        }
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    Process{
        if($PSCmdlet.ShouldProcess($Identity,"Remove ""Send As"" permissons for $User")){
            Remove-O365RecipientPermission -Identity $Identity -Trustee $User -AccessRights SendAs -Confirm:$false
        }
    }
}

<#
.SYNOPSIS
List all Office 365 Commands
.DESCRIPTION
This is a shortcut to list all the commands imported by the Office 365 connection and 
those added by this helper script.
.EXAMPLE
Get-O365Commands
.EXAMPLE
Get-O365Commands | slect-string "mailbox"

Find all the Mailbox related offcie 365 commands.
.NOTES
For OnPrem Mail commands see Get-OnPremCommands
#>
function Get-O365Commands {
    if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    get-command -Name *O365*
}

<#
.SYNOPSIS
List all OnPrem Mail Commands
.DESCRIPTION
This is a shortcut to list all the commands imported by the OnPrem Mail connection and 
those added by this helper script.
.EXAMPLE
Get-OnPremCommands
.EXAMPLE
Get-OnPremCommands | slect-string "mailbox"

Find all the Mialbox related OnPrem Mail commands.
.NOTES
For Office 365 see Get-O365Commands
#>
function Get-OnPremCommands {
    if(!(Get-PSSession -Name OnPremMail -ea SilentlyContinue)){Connect-O365}
    get-command -Name *OnPrem*
}


function Get-O365DistributionGroupMembership {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$Identity,
        [switch]$Raw
    )
    begin {
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    process {
        if($PSCmdlet.ShouldProcess($Identity)){
            $Recipients = (
                Get-O365Recipient -ResultSize Unlimited -filter "members -eq '$(
                    (Get-O365Recipient -Identity $Identity -ea SilentlyContinue).distinguishedname
                    )'" -ea SilentlyContinue 
                ).PrimarySmtpAddress
            if($Recipients -and !$Raw){ ($Recipients |  Get-O365DistributionGroup -ResultSize Unlimited) } 
            else{ ($Recipients) }
        } #End ShouldProcess
    } #End Process
} #End Function

<#
.SYNOPSIS
Get all the distribution lists a receipient is a effectively a member of
.DESCRIPTION
This will retrun all the distribution lists a recipient is a direct member of as
as all the groups parent distributions groups the recipient is effectively a member
of. This ultimatelte shows all the distribution lists a recipient will receive mail 
to even if they are not a direct member of that group.

For example, if Bob is a member of the "Mobile IT Dallas" group and the "Mobile IT 
Dallas" group is a member of the "Mobile IT" group, get Get-O365DistributionGroupMembership
would only return "Mobile IT Dallas." But Get-O365EffectiveDistributionGroupMembership would 
return both "Mobile IT Dallas" and "Mobile IT."
.PARAMETER Identity
A string containing an Office 365 mailbox Identity such as a displayname or email
address. This is the whose distributions lists will be returned.
.PARAMETER Raw
A boolean that if true will cause the fucntion to return a string list of 
distribution list email addresses. If it is false, the fucnction will return 
distribution list objects (This the default). 
.EXAMPLE
Get-O365EffectiveDistributionGroupMembership -Identity bob.testerton@contoso.com

Return distribution list objects for distribution lists bob.testerton@contoso.com 
is a effectively a member of.
.EXAMPLE
Get-O365EffectiveDistributionGroupMembership -Identity bob.testerton@contoso.com -Raw:$true

Retrun distribution lists as a list of email addresses for distribution lists 
bob.testerton@contoso.com is effectively a member of.
.NOTES
Se also Get-O365DistributionGroupMembership
#>
function Get-O365EffectiveDistributionGroupMembership {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$Identity,
        [switch]$Raw
    )
    begin {
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    process {
        $recurse = { 
            param($Myident)
            foreach ($CurEmail in (Get-O365DistributionGroupMembership -Identity $Myident -Raw))
            {
                $CurEmail
                $recurse.Invoke($CurEmail)
            }
        }
        if($PSCmdlet.ShouldProcess($Identity)){
                if(!$Raw){ (($recurse.Invoke($Identity)) | Get-O365DistributionGroup) }
                else{($recurse.Invoke($Identity))}
        } #End ShouldProcess
    } #End Process
} #End Function

function Get-O365EffectiveDistributionGroupMembers {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$Identity,
        [switch]$Raw
    )
    begin {
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    process {
        $recurse = { 
            param($Myident)
            foreach ($CurEmail in (Get-O365DistributionGroupMember -Identity $Myident))
            {
                if($CurEmail.ObjectClass -contains 'group')
                {
                    $recurse.Invoke($CurEmail.PrimarySMTPAddress)
                }
                else
                {
                   $CurEmail.primarySMTPAddress
                } 
            }
        }
        if($PSCmdlet.ShouldProcess($Identity)){
                if(!$Raw){ (($recurse.Invoke($Identity)) | Get-O365Mailbox) }
                else{($recurse.Invoke($Identity))}
        } #End ShouldProcess
    } #End Process
} #End Function

<#
For Future Development
function Convert-0365SecurityToDistribution{
     param(
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )
    $OldGroup = Get-O365DistributionGroup -Identity $Identity
    $Members = Get-O365DistributionGroupMember -Identity $Identity
    $MemberOf = Get-O365DistributionGroupMembership -Identity $Identity
    $EmailAddresses = $OldGroup.EmailAddresses
    $DisplayName = $OldGroup.DisplayName
    $ManagedBy = $OldGroup.ManagedBy
}
#>
 
<#
.SYNOPSIS
Forward mail for a mailbox
.DESCRIPTION
This will forward mail from a specificed mailbox to a specified recipient. The
recipient can be either an internal or external user.
.PARAMETER Identity
A string containing an Office 365 mailbox Identity such as a displayname or email
address. This is the mailbox or user whose email willbe forwarded.
.PARAMETER Recpient
A string containing either an external email address or an Office 365 mailbox 
Identity such as a displayname. This is the recpient of the forwarded email.
.PARAMETER SaveAndFoward
A boolean used to determin if email to the mailbox will be saved and forwarded
or forwaded and not saved. The default value of $false means mail will not be 
saved and will only be forwarded.
.EXAMPLE
Add-O365MailboxForward -Identity bob.testerton@contoso.com -Recipient jill.testertong@contoso.com -SaveAndFoward $true

Have all mail sent to bob.testerton@contoso.com forwarded to jill.testertong@contoso.com
And save a copy of the forwarded email in bob.testerton@contoso.com's mailbox.
.EXAMPLE
Add-O365MailboxForward -Identity bob.testerton@contoso.com -Recipient jill.testertong@contoso.com

Have all mail sent to bob.testerton@contoso.com forwarded to jill.testertong@contoso.com
And not save a copy of the forwarded email in bob.testerton@contoso.com's mailbox.
#>    
function Add-O365MailboxForward {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
    param(
            [Parameter(Mandatory=$true)]
            [string]$Identity,
            [Parameter(Mandatory=$true)]
            [string]$Recipient,
            [bool]$SaveAndFoward = $false
    )
    begin {
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    process{
        if($PSCmdlet.ShouldProcess($Identity,"Forwad mail to $Recipient")){
            if((Validate-Email -Email $Recipient) -and ((Get-O365AcceptedDomain).DomainName.ToLower() -notcontains $Recipient.Split('@')[1].ToLower())) {
                Set-O365Mailbox -Identity $Identity -ForwardingSmtpAddress $Recipient -DeliverToMailboxAndForward $SaveAndFoward
            }
            else {
                Set-O365Mailbox -Identity $Identity -ForwardingAddress $Recipient -DeliverToMailboxAndForward $SaveAndFoward
            }
        } #End ShouldProcess
    } #End Process
} #End Function

Function New-PasswordString {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="low")]
    param(
        [alias("Upper")]
        [Switch]$UpperCase,
        [alias("Lower")]
        [switch]$LowerCase,
        [alias("Num","Numeral")]
        [switch]$Number,
        [alias("Sym","Special","SpecialChar","SpecialCharacter")]
        [switch]$Symbol,
        [alias("Default")]
        [switch]$All,
        [switch]$AlphaNumeric,
        [alias("MaxCharacters","NumCharacters","Length")]
        [ValidateRange(4,99)]
        [int]$PasswordLength = 12,
        [alias("NumberPasswords","Repeat")]
        [ValidateRange(1,100)]
        [int]$Count = 1

    )
    begin { 
        if($All -or (!$UpperCase -and !$LowerCase -and !$Number -and !$Symbol -and !$AlphaNumeric)) {
            $UpperCase = $True
            $LowerCase = $True
            $Number = $True
            $Symbol = $True
        }
        if($AlphaNumeric) {
            $UpperCase = $True
            $LowerCase = $True
            $Number = $True
        }

    }
    process {
        1..$Count | ForEach-Object {
            if($PSCmdlet.ShouldProcess($_)) {
                $MyPasswordLength = $PasswordLength
                Remove-Variable -name MyInputArray -ea SilentlyContinue
                Remove-Variable -name MyPassword -ea SilentlyContinue
                if($UpperCase) {
                    $MyPassword = $MyPassword + ([char[]](Get-Random -Input $(65..90) -Count 1))
                    $MyInputArray = $MyInputArray + $(65..90)
                    $MyPasswordLength -= 1
                }
                if($LowerCase) {
                    $MyPassword = $MyPassword + ([char[]](Get-Random -Input $(97..122) -Count 1))
                    $MyInputArray = $MyInputArray + $(97..122)
                    $MyPasswordLength -= 1
                }
                if($Number) {
                    $MyPassword = $MyPassword + ([char[]](Get-Random -Input $(48..57) -Count 1))
                    $MyInputArray = $MyInputArray + $(48..57)
                    $MyPasswordLength -= 1 
                }
                if($Symbol) {
                    $MyPassword = $MyPassword + ([char[]](Get-Random -Input $(33..38) -Count 1))
                    $MyInputArray = $MyInputArray + $(33..38)
                    $MyPasswordLength -= 1
                }
                $MyPassword = $MyPassword + ([char[]](Get-Random -Input $MyInputArray -Count $MyPasswordLength))
                $MyPassword = $MyPassword -join ""
                $MyPassword = [String]::Join("",(Get-Random -InputObject $MyPassword.ToCharArray() -Count ([int]::MaxValue)))
                $MyPassword
            }
        }
    }
}

function Get-O365OWAURL {
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
    param(
            [Parameter(Mandatory=$true)]
            [string]$Identity
    )
    begin {
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    process{
        if($PSCmdlet.ShouldProcess($Identity)){
            $mailbox = Get-O365Mailbox -Identity $Identity 
            if($mailbox){
                "https://outlook.office365.com/owa/$($mailbox.PrimarySmtpAddress)"
            }
        }
    }
}

function Clear-O365MailboxContents{
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="High")]
    param(
            [Parameter(Mandatory=$true)]
            [string]$Identity
    )
    begin {
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    process{
        if($PSCmdlet.ShouldProcess($Identity, "Delete all emails")){
            $OldPerf = $Global:ProgressPreference
            $Global:ProgressPreference = ’SilentlyContinue’
            $mailbox = Get-O365Mailbox -Identity $Identity 
            if($mailbox){
                while((Search-O365Mailbox -Identity $mailbox.ExchangeGuid.ToString() `
                    -SearchDumpster:$false `
                    -SearchQuery "Kind:email" `
                    -EstimateResultOnly `
                    -DoNotIncludeArchive `
                    -WarningAction SilentlyContinue).ResultItemsCount -gt 0 )
                {
                    Search-O365Mailbox -Identity $mailbox.ExchangeGuid.ToString() `
                        -SearchDumpster:$false `
                        -SearchQuery "Kind:email" `
                        -DeleteContent `
                        -Force `
                        -WarningAction SilentlyContinue
                }
            }
            $Global:ProgressPreference = $OldPerf
        }
    }
}


function Get-O365DynamicDistributionGroupMembers {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$Identity
    )
    begin {
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
    }
    process{
        if($PSCmdlet.ShouldProcess($Identity)){
            Get-O365Recipient -Filter (Get-O365DynamicDistributionGroup -Identity $Identity).RecipientFilter
        }
    }
}

function Get-O365DynamicDistributionGroupMembership {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$Identity
    )
    begin {
        if(!(Get-PSSession -Name O365 -ea SilentlyContinue)){Connect-O365}
        $DynamicGroups = Get-O365DynamicDistributionGroup
    }
    process{
        if($PSCmdlet.ShouldProcess($Identity)){
            if($Recipient = Get-O365Recipient -Identity $Identity){
                $DynamicGroups | ForEach-Object{
                    if(Get-O365Recipient -Filter "$($_.RecipientFilter) -and DistinguishedName -eq '$($Recipient.DistinguishedName)'"){
                        $_
                    }
                }
            }
        }
    }
}
