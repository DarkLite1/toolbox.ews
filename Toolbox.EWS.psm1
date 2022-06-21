#Requires -Modules MSAL.PS

$ewsDLL = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
$ExchangeVersion = 'Exchange2013_SP1'
$azureClientId = $env:AZURE_CLIENT_ID
$azureTenantId = $env:AZURE_TENANT_ID

try {
    try {
        Import-Module -Name $ewsDLL -EA Stop
    }
    catch {
        throw "Failed loading the DLL file '$ewsDLL': $_"
    }
    if (-not $azureClientId) {
        throw 'Azure Client ID is required'
    }
    if (-not $azureTenantId) {
        throw 'Azure Tenant ID is required'
    }
}
catch {
    throw "Failed loading the Exchange Web Service module: $_"
}

Function Find-MailboxFolderIdHC {
    <# 
    .SYNOPSIS   
        Search for the folder ID of a specific folder in a mailbox.

    .DESCRIPTION
        Search for the folder ID of a specific folder in a mailbox.

    .PARAMETER Path
        The path to search for within the mailbox. When searching for an exact 
        path this should be
        split with '\'. Example: '\PowerShell\Tickets SENT' or 'PowerShell''

    .PARAMETER Mailbox
        E-mail address

    .PARAMETER Service
        The Exchange Web Service object used to authenticate ourselves

    .EXAMPLE
        $params = @{
            Path    = '\Inbox\PowerShell\Expiring users' 
            Mailbox = 'Jos@brink.nl'
            Service = $Service
        }
        Find-MailboxFolderIdHC @params

        Searches for the folder 'Expiring users' under the root 
        folder 'Inbox\PowerShell' in the mailbox of 'Jos@brink.nl'
    #>

    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]$Path,
        [Parameter(Mandatory)]
        [String]$Mailbox,
        [Parameter(Mandatory)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service
    )

    Try {
        $FolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $Mailbox)
        $DataFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $FolderID)
    }
    Catch {
        throw "Folder '$P' not found in mailbox '$Mailbox'. Has user '$env:USERNAME' 'Full mailbox control' permissions? Is this an Office 365 mailbox?: $_"
    }

    $Array = $Path.Split('\')

    for ($i = 1; $i -lt $Array.Length; $i++) {
        $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1) 
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $Array[$i]) 
        $Results = $service.FindFolders($DataFolder.Id, $SearchFilter, $FolderView) 

        if ($Results.TotalCount -gt 0) { 
            foreach ($R in $Results.Folders) { 
                $DataFolder = $R                
            }
        } 
        else {
            $NotFound = $true
        }     
    }

    if ($NotFound) {
        if ($ErrorActionPreference -eq 'Ignore') {
            Write-Warning "Folder path '$Path' not found in the mailbox '$Mailbox'"
        }
        else {
            throw "Folder path '$Path' not found in the mailbox '$Mailbox'"
        }
    }
    else {
        $DataFolder.Id
        Write-Debug "Found folder '$Path' in '$Mailbox' with ID '$($DataFolder.Id.UniqueId)'"
    }
}
Function New-EwsServiceHC {
    <# 
    .SYNOPSIS   
        Create a new EWS service object.

    .DESCRIPTION
        Create a new Exchange Webs Services object with the correct configuration.

    .PARAMETER ExchangeVersion
        The version of Exchange currently in use.
        Options: https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.exchangeversion(v=exchg.80).aspx
    #>

    Param (
        [String]$ExchangeVersion = $ExchangeVersion
    )

    Try {
        $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $ExchangeVersion
        $Service.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
        $Service.UseDefaultCredentials = $false
        try {
            $msalParams = @{
                ClientId              = $azureClientId
                TenantId              = $azureTenantId
                IntegratedWindowsAuth = $true
                Scopes                = "https://outlook.office.com/EWS.AccessAsUser.All"
            }
            $token = Get-MsalToken @msalParams 
            $Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$token.AccessToken
        }
        Catch {
            throw "Failed retrieving a valid oAuth token form Azure: $_"
        }
        
        return $Service
    }
    Catch {
        throw "Failed creating a new EWS service object: $_"
    }
}
Function New-MailboxFolderHC {
    <# 
    .SYNOPSIS   
        Create a new folder in a mailbox when it doesn't exist yet.

    .DESCRIPTION
        Create a new folder in a mailbox when it doesn't exist yet.

    .PARAMETER Path
        The path to search for within the mailbox. When searching for an exact 
        path this should be split with '\'.
        Example: '\PowerShell\Tickets SENT' or 'PowerShell''

    .PARAMETER Mailbox
        Smtp address of the mailbox.

    .PARAMETER Service
        The Exchange Service object, aka EWS.

    .EXAMPLE
        $params = @{
            Mailbox = 'SrvBatch@heidelbergcement.com'
            Service = $Service
        }
        '\color\green\dark', '\fruit\kiwi' | New-MailboxFolderHC @params

        Create 2 folders in the mailbox root folder when they don't exist 
        already
    #>

    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [String[]]$Path,
        [Parameter(Mandatory)]
        [String]$Mailbox,
        [Parameter(Mandatory)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service
    )

    Process {
        foreach ($P in $Path) {
            Try {
                $FolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $Mailbox)
                $DataFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $FolderID)
            }
            Catch {
                throw "Folder '$P' not found in mailbox '$Mailbox'. Has user '$env:USERNAME' 'Full mailbox control' permissions? Is this an Office 365 mailbox?: $_"
            }

            $Array = $P.Split('\')

            for ($i = 1; $i -lt $Array.Length; $i++) {
                $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1) 
                $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $Array[$i]) 
                $Results = $Service.FindFolders($DataFolder.Id, $SearchFilter, $FolderView) 

                if ($Results.TotalCount -gt 0) { 
                    foreach ($R in $Results.Folders) { 
                        $DataFolder = $R
                    }
                } 
                else {
                    Write-Verbose "Folder '$($Array[0..$i] -join '\')' not found in mailbox '$Mailbox'"
                    $NewFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
                    $NewFolder.DisplayName = $Array[$i]
                    $NewFolder.Save($DataFolder.Id)
                    $DataFolder = $NewFolder
                    Write-Verbose "Created folder '$($Array[0..$i] -join '\')' in mailbox '$Mailbox'"
                }     
            }
        }
    }
}
Function Send-MailAuthenticatedHC {
    <# 
    .SYNOPSIS   
        Send an e-mail message as an authenticated user.

    .DESCRIPTION
        Send an e-mail message as an authenticated user by providing the 
        correct credentials and 'From' e-mail address. The e-mail will be 
        available in the 'Sent items' folder to the sender.

    .PARAMETER To
        Recipient(s) you wish to e-mail.

    .PARAMETER Bcc
        The e-mail address of the recipient(s) you wish to e-mail in Blind 
        Carbon Copy. Other users will not see the e-mail address of users in 
        the 'Bcc' field.

    .PARAMETER Cc
        The e-mail address of the recipient(s) you wish to e-mail in Carbon 
        Copy.

    .PARAMETER From 
        The e-mail address from which the mail is sent. You have to have the 
        correct permissions to do this provided by a PS Credential Object.
    
    .PARAMETER Subject 
        The Subject-header used in the e-mail.

    .PARAMETER Body 
        The message within the e-mail.

    .PARAMETER Priority
        Specifies the priority of the e-mail message. Valid values are 
        'Normal', 'High', and 'Low'. If not specified, the default value is 
        'Normal'.

    .PARAMETER Attachments
        Specifies the full path name to the files you want to attach to the 
        e-mail.

    .PARAMETER ExchangeVersion
        The version of Exchange currently in use. If the wrong code is used we 
        throw an error to tell you.
        Options: https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.exchangeversion(v=exchg.80).aspx

    .PARAMETER API
        The API used to access the needed Exchange functionalities called 'EWS, 
        Exchange.WebServices'.
        Download: http://www.microsoft.com/en-us/download/details.aspx?id=42951

    .PARAMETER SentItemsPath
        The path in the mailbox where we will save the mail after sending it. 
        By default we save the mail in the folder 'Sent items'. This path is 
        expressed as '\Inbox\My folder\Path'

    .PARAMETER EventLogSource
        Log source where the event will be saved, By default we save all 
        actions of sending e-mails in the Windows Event Log.

    .EXAMPLE
        $params = @{
            Credential  = Get-Credential
            From        = 'Warner@Bross.com'
            To          = @('Chuck@Norris.com','BobLee@Swagger.com')
            Subject     = 'Update'
            Attachments = (Get-ChildItem -File -Path 'T:\Share').FullName
        }
        'Hello world' | Send-MailAuthenticatedHC @params
        
        Send an e-mail with attachment and store it in the standard folder 
        'Sent items' of the authenticated user.

    .EXAMPLE
        $params = @{
            Credential    = Get-Credential
            From          = $SrvBatchMailbox
            To            = @('Chuck@Norris.com')
            Subject       = 'Expiring user accounts'
            SentItemsPath = '\Inbox\PowerShell\Expiring users OUT'
            DraftPath     = '\Inbox\PowerShell\Draft'
        }
        'Hello world' | Send-MailAuthenticatedHC @params

        Send an e-mail to Chuck, save it in the folder 
        '\Inbox\PowerShell\Expiring users OUT' after sending, and use the 
        draft folder '\Inbox\PowerShell\Draft'
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [parameter(Mandatory)]
        [ValidateScript( { Get-ADObject -LDAPFilter "(|(mail=$_)(proxyAddresses=smtp:$_))" })]
        [String]$From,
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String[]]$To,
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$Subject,
        [parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [String]$Body,
        [parameter(Mandatory)]
        [String]$EventLogSource,
        [String[]]$Cc,
        [String[]]$Bcc,
        [ValidateScript( { Test-Path $_ -PathType Leaf })]
        [String[]]$Attachments,
        [ValidateSet('Low', 'Normal', 'High')]
        [String]$Priority = 'Normal',
        [String]$SentItemsPath
    )

    Process {
        Try {
            Set-EWScredentialsSilentlyHC -Service $Service # for long running scripts we need a fresh token

            $Message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $Service
            $Message.Subject = $Subject
            $Message.From = $From
            $Message.Importance = $Priority
            $Message.Body = $Body

            $To | ForEach-Object { $null = $Message.ToRecipients.Add($_) }

            if ($Cc) { $Cc | ForEach-Object { $null = $Message.CcRecipients.Add($_) } }

            if ($Bcc) { $Bcc | ForEach-Object { $null = $Message.BccRecipients.Add($_) } }

            if ($Attachments) {
                # Excel files that are opened can't be sent as attachment, so we copy them first
                $Attachment = New-Object System.Collections.ArrayList($null)
        
                $TmpFolder = "$env:TEMP\Send-MailAuthenticatedHC {0}" -f (Get-Random)
                foreach ($a in $Attachments) {
                    if ($a -like '*.xlsx') {
                        if (-not(Test-Path $TmpFolder)) {
                            $null = New-Item $TmpFolder -ItemType Directory
                        }
                        Copy-Item $a -Destination $TmpFolder

                        $Attachment.Add("$TmpFolder\$(Split-Path $a -Leaf)")
                    }
                    else {
                        $Attachment.Add($a) | Out-Null
                    }
                }        
                $Attachment | ForEach-Object {
                    $null = $Message.Attachments.AddFileAttachment($_)
                    Write-Verbose "Mail attachment added: '$(Split-Path $_ -Leaf)'"
                }
            }

            if ($SentItemsPath) {
                $SentItemsFolderID = Find-MailboxFolderIdHC -Path $SentItemsPath -Mailbox $From -Service $Service
                    
                # in case of attachments, we need to save the mail first before sending it
                $Message.Save($SentItemsFolderID)
                Write-Debug "Mail saved in '$SentItemsPath' of user '$From'"
                $Message.SendAndSaveCopy($SentItemsFolderID)
            }
            else {
                Write-Debug "Mail saved in 'Sent items' of '$($Credentials.UserName)'"
                $Message.SendAndSaveCopy()
            }

            Write-Verbose "Mail sent from '$From' to '$To' with subject '$Subject'"

            #region Save in event log
            Import-EventLogParamsHC -Source $EventLogSource

            Write-EventLog @EventOutParams -Message ($env:USERNAME + ' - ' + 'Mail sent authenticated' + "`n`n" + 
                "- Subject:`t" + $Message.Subject + "`n" +
                "- To:`t`t" + $Message.ToRecipients + "`n" +
                "- CC:`t`t" + $Message.CcRecipients + "`n" +
                "- BCC:`t`t" + $Message.BccRecipients + "`n" +
                "- Priority:`t" + $Message.Importance + "`n" +
                "- From:`t`t" + $Message.From + "`n" +
                "- Attachments:`t" + $Message.Attachments + "`n" +
                "- Script location:`t" + $global:PSCommandPath
            )
            #endregion
        }
        Catch {
            throw "Send-MailAuthenticatedHC: $_"
        }
    }
    End {
        if ($Attachments) {
            if (Test-Path -LiteralPath $TmpFolder) {
                Remove-Item -LiteralPath $TmpFolder -Recurse
            }
        }
    }
}
Function Set-EWScredentialsSilentlyHC {
    <# 
    .SYNOPSIS   
        Refresh the credentials.

    .DESCRIPTION
        Set the credentials on the Exchange WebsServices object. The credentials
        are set by the MSAL.PS library upon creating of the Exchange Web Service
        object with Integrated Windows Authentication. 

        Later on, for long running tasks, the credentials need to be refreshed
        silently. This function checks for a valid token in the cache, when no
        valid cached token is found a new token is requested.

    .PARAMETER Service
        The Exchange Service object, aka EWS.

    .NOTES
        Azure App Registration:
        - Treat application as a public client: Yes
        - RedirectUri for Mobile and Desktop apps: urn:ietf:wg:oauth:2.0:oob

        1. login with account srvbatch@grouphc.net on Windows
        - RDP session for testing
        - Run scheduled task as account

        2. Acquire a token with Integrated Windows Authentication

        $msalParams = @{
            ClientId              = $azureClientId
            TenantId              = $azureTenantId 
            Scopes                = "https://outlook.office.com/EWS.AccessAsUser.All"
            IntegratedWindowsAuth = $true
        }
        Get-MsalToken @msalParams 

        3. Refresh the token silently

        $msalParams = @{
            ClientId  = $azureClientId
            TenantId  = $azureTenantId 
            Scopes    = "https://outlook.office.com/EWS.AccessAsUser.All"
            Silent    = $true
        }
        Get-MsalToken @msalParams
    #>

    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [String[]]$ClientId = $azureClientId,
        [String]$TenantId = $azureTenantId 
    )
        
    Try {
        $msalParams = @{
            ClientId = $azureClientId
            TenantId = $azureTenantId
            Silent   = $true
            Scopes   = "https://outlook.office.com/EWS.AccessAsUser.All"
        }
        $token = Get-MsalToken @msalParams
        # Write-Verbose "Set access token '$($token.AccessToken)'"
        $Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$token.AccessToken
    }
    Catch {
        throw "Failed setting the EWS credentials with the oAuth access token: $_"
    }
}

Export-ModuleMember -Function * -Alias *