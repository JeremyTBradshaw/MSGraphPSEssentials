#Requires -Version 5.1
using namespace System
using namespace System.Management.Automation.Host
using namespace System.Runtime.InteropServices
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates

<# Release Notes for v0.5.5 (2021-09-08):

    - Added [-ExoEwsAppOnlyScope] switch parameter for New-MSGraphAccessToken's ClientCredentials parameter sets.
        --  This will change the scope to https://outlook.office365.com/.default instead of the typical
        https://graph.microsoft.com/.default, to enable OAuth app-only authentication with Exchange Online for EWS
        applications.
        -- Delegated permissions / user-present auth. flows for EWS are already covered in the DeviceCode and
        RefreshToken parameter sets (i.e., supply -Scopes Ews.AccessUser.All).
#>

function New-MSGraphAccessToken {
    [CmdletBinding(
        DefaultParameterSetName = 'DeviceCode_Endpoint'
    )]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ClientCredentials_Certificate')]
        [Parameter(Mandatory, ParameterSetName = 'ClientCredentials_CertificateStorePath')]
        [Parameter(Mandatory, ParameterSetName = 'DeviceCode_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'RefreshToken_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'RefreshTokenCredential_TenantId')]
        [string]$TenantId, # Guid / FQDN

        [Parameter(Mandatory, ParameterSetName = 'ClientCredentials_Certificate')]
        [Parameter(Mandatory, ParameterSetName = 'ClientCredentials_CertificateStorePath')]
        [Parameter(Mandatory, ParameterSetName = 'DeviceCode_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'DeviceCode_Endpoint')]
        [Parameter(Mandatory, ParameterSetName = 'RefreshToken_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'RefreshToken_Endpoint')]
        [Guid]$ApplicationId,

        [Parameter(
            Mandatory,
            ParameterSetName = 'ClientCredentials_Certificate',
            HelpMessage = 'E.g. Use $Certificate, where `$Certificate = Get-ChildItem cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [X509Certificate2]$Certificate,

        [Parameter(
            Mandatory,
            ParameterSetName = 'ClientCredentials_CertificateStorePath',
            HelpMessage = 'E.g. cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317; E.g. cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [ValidateScript(
            {
                if (Test-Path -Path $_) { $true } else {

                    throw "An example proper path would be 'cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'."
                }
            }
        )]
        [string]$CertificateStorePath,

        [Parameter(ParameterSetName = 'ClientCredentials_Certificate')]
        [Parameter(ParameterSetName = 'ClientCredentials_CertificateStorePath')]
        [switch]$ExoEwsAppOnlyScope,

        [Parameter(ParameterSetName = 'ClientCredentials_Certificate')]
        [Parameter(ParameterSetName = 'ClientCredentials_CertificateStorePath')]
        [ValidateRange(1, 10)]
        [int16]$JWTExpMinutes = 2,

        [Parameter(ParameterSetName = 'DeviceCode_Endpoint')]
        [Parameter(ParameterSetName = 'RefreshToken_Endpoint')]
        [Parameter(ParameterSetName = 'RefreshTokenCredential_Endpoint')]
        [ValidateSet('Common', 'Consumers', 'Organizations')]
        [string]$Endpoint = 'Common',

        [Parameter(Mandatory, ParameterSetName = 'DeviceCode_TenantId', HelpMessage = 'E.g. Mail.Send, Ews.AccessAsUser.All')]
        [Parameter(Mandatory, ParameterSetName = 'DeviceCode_Endpoint', HelpMessage = 'E.g. Mail.Send, Ews.AccessAsUser.All')]
        [Parameter(ParameterSetName = 'RefreshToken_TenantId')]
        [Parameter(ParameterSetName = 'RefreshToken_Endpoint')]
        [Parameter(ParameterSetName = 'RefreshTokenCredential_TenantId')]
        [Parameter(ParameterSetName = 'RefreshTokenCredential_Endpoint')]
        [string[]]$Scopes,

        [Parameter(Mandatory, ParameterSetName = 'RefreshToken_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'RefreshToken_Endpoint')]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.refresh_token) { $true } else {

                    throw 'Invalid access token.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken -Scopes offline_access...'
                }
            }
        )]
        [Object]$RefreshToken,

        [Parameter(Mandatory, ParameterSetName = 'RefreshTokenCredential_TenantId')]
        [Parameter(Mandatory, ParameterSetName = 'RefreshTokenCredential_Endpoint')]
        [PSCredential]$RefreshTokenCredential
    )

    #region Initialization
    if ($PSCmdlet.ParameterSetName -eq 'ClientCredentials_CertificateStorePath') {
        try {
            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        catch { throw }
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'ClientCredentials_Certificate') {

        $Script:Certificate = $Certificate
    }

    if ($PSCmdlet.ParameterSetName -like 'ClientCredentials_*') {

        if (-not (Test-SigningCertificate -Certificate $Script:Certificate)) {

            throw "The supplied certificate must use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider', " +
            'and the SHA-256 hashing algorithm.  ' +
            'For best luck, use a certificate generated using New-SelfSignedMSGraphApplicationCertificate.'
        }
    }

    if ($PSCmdlet.ParameterSetName -like '*_TenantId') { $Script:Endpoint = $TenantId } else { $Script:Endpoint = $Endpoint }

    if ($PSCmdlet.ParameterSetName -like 'RefreshTokenCredential_*') {

        try {
            $Script:ApplicationId = [Guid]$RefreshTokenCredential.UserName
            $Script:RefreshToken = ConvertFrom-Json (ConvertFrom-SecureStringToPlainText $RefreshTokenCredential.Password)
        }
        catch {
            'Failed to validate refresh token credential object. ' +
            'Supply $RefreshTokenObject where $RefreshTokenObject = New-RefreshTokenObject ...' |
            Write-Warning
            throw
        }
    }
    elseif ($PSCmdlet.ParameterSetName -like 'RefreshToken_*') {

        $Script:ApplicationId = $ApplicationId
        $Script:RefreshToken = $RefreshToken
    }
    #endregion Initialization

    #region Functions
    function New-AppOnlyAccessToken ($TenantId, $ApplicationId, $Certificate, $JWTExpMinutes) {
        try {
            $NowUTC = [datetime]::UtcNow

            $EncodedHeader = [Convert]::ToBase64String(
                [Text.Encoding]::UTF8.GetBytes(
                    (ConvertTo-Json -InputObject (
                            @{
                                alg = 'RS256'
                                typ = 'JWT'
                                x5t = ConvertTo-Base64Url -String ([Convert]::ToBase64String($Certificate.GetCertHash()))
                            }
                        )
                    )
                )
            )

            $EncodedPayload = [Convert]::ToBase64String(
                [Text.Encoding]::UTF8.GetBytes(
                    (ConvertTo-Json -InputObject (
                            @{
                                aud = "https://login.microsoftonline.com/$TenantId/oauth2/token"
                                exp = (Get-Date $NowUTC.AddMinutes($JWTExpMinutes) -UFormat '%s') -replace '\..*'
                                iss = $ApplicationId.Guid
                                jti = [Guid]::NewGuid()
                                nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
                                sub = $ApplicationId.Guid
                            }
                        )
                    )
                )
            )

            $JWT = (ConvertTo-Base64Url -String $EncodedHeader, $EncodedPayload) -join '.'

            $Signature = ConvertTo-Base64Url -String (
                [Convert]::ToBase64String(
                    $Script:Certificate.PrivateKey.SignData(
                        [Text.Encoding]::UTF8.GetBytes($JWT),
                        [HashAlgorithmName]::SHA256,
                        [RSASignaturePadding]::Pkcs1
                    )
                )
            )

            $JWT = $JWT + '.' + $Signature

            $trBody = @{

                client_id             = $ApplicationId
                client_assertion      = $JWT
                client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                scope                 = "$(if ($ExoEwsAppOnlyScope) {'https://outlook.office365.com' } else { 'https://graph.microsoft.com' })/.default"
                grant_type            = "client_credentials"
            }

            $trParams = @{

                Method      = 'POST'
                Uri         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
                Body        = $trBody
                Headers     = @{ Authorization = "Bearer $($JWT)" }
                ContentType = 'application/x-www-form-urlencoded'
                UserAgent   = "MSGraphPSEssentials/$($PSCmdlet.MyInvocation.MyCommand.Module.Version)"
                ErrorAction = 'Stop'
            }

            $trResponse = Invoke-RestMethod @trParams

            # Output the token request response:
            $trResponse
        }
        catch { throw }
    }

    function New-DeviceCodeAccessToken ($Endpoint, $ApplicationId, $Scopes) {
        try {
            $dcrBody = @(
                "client_id=$($ApplicationId)",
                "scope=$($Scopes -join ' ')"
            ) -join '&'

            $dcrParams = @{

                Method      = 'POST'
                Uri         = "https://login.microsoftonline.com/$($Endpoint)/oauth2/v2.0/devicecode"
                Body        = $dcrBody
                ContentType = 'application/x-www-form-urlencoded'
                UserAgent   = "MSGraphPSEssentials/$($PSCmdlet.MyInvocation.MyCommand.Module.Version)"
                ErrorAction = 'Stop'
            }
            $dcrResponse = Invoke-RestMethod @dcrParams

            $dtNow = [datetime]::Now
            $sw1 = [Diagnostics.Stopwatch]::StartNew()
            $dcExpiration = "$($dtNow.AddSeconds($dcrResponse.expires_in).ToString('yyyy-MM-dd hh:mm:ss tt'))"

            $trBody = @(
                "grant_type=urn:ietf:params:oauth:grant-type:device_code",
                "client_id=$($ApplicationId)",
                "device_code=$($dcrResponse.device_code)"
            ) -join '&'

            # Wait for user to enter code before starting to poll token endpoint:
            switch (
                $host.UI.PromptForChoice(

                    "Authorization started (expires at $($dcExpiration))",
                    "$($dcrResponse.message)",
                    [ChoiceDescription]('&Done'),
                    0
                )
            ) { 0 { <##> } }

            if ($sw1.Elapsed.Minutes -lt 15) {

                $sw2 = [Diagnostics.Stopwatch]::StartNew()
                $successfulResponse = $false
                $pollCount = 0
                do {
                    if ($sw2.Elapsed.Seconds -ge $dcrResponse.interval) {

                        $sw2.Restart()
                        $pollCount++

                        try {
                            $trParams = @{

                                Method      = 'POST'
                                Uri         = "https://login.microsoftonline.com/$($Endpoint)/oauth2/v2.0/token"
                                Body        = $trBody
                                ContentType = 'application/x-www-form-urlencoded'
                                UserAgent   = "MSGraphPSEssentials/$($PSCmdlet.MyInvocation.MyCommand.Module.Version)"
                                ErrorAction = 'Stop'
                            }
                            $trResponse = Invoke-RestMethod @trParams
                            $successfulResponse = $true
                        }
                        catch {
                            if ($_.ErrorDetails.Message) {

                                $badResponse = ConvertFrom-Json -InputObject $_.ErrorDetails.Message

                                if ($badResponse.error -eq 'authorization_pending') {

                                    if ($pollCount -eq 1) {

                                        "The user hasn't finished authenticating, but hasn't canceled the flow (error: authorization_pending).  " +
                                        "Continuing to poll the token endpoint at the requested interval ($($dcrResponse.interval) seconds)." |
                                        Write-Warning
                                    }
                                }
                                elseif ($badResponse.error -match '^(authorization_declined)|(bad_verification_code)|(expired_token)$') {

                                    # https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code#expected-errors
                                    throw "Authorization failed due to foreseeable error: $($badResponse.error)."
                                }
                                elseif ($_.errorDetails.message -match '(AADSTS7000218)') {

                                    "Authorization failed due to 'invalid_client' (AADSTS7000218). " +
                                    'Ensure the Application is enabled for public client flows in Azure AD. ' +
                                    'If this is app is intended for app-only/unattended use, you should instead use the -Certificate/-CertificateStorePath parameters.' |
                                    Write-Warning
                                    throw
                                }
                                else {
                                    Write-Warning 'Authorization failed due to an unexpected error.'
                                    throw $badResponse.error_description
                                }
                            }
                            else {
                                Write-Warning 'An error was encountered with the Invoke-RestMethod command.  Authorization request did not complete.'
                                throw
                            }
                        }
                    }
                    if (-not $successfulResponse) { Start-Sleep -Seconds 1 }
                }
                while ($sw1.Elapsed.Minutes -lt 15 -and -not $successfulResponse)

                # Output the token request response:
                $trResponse
            }
            else {
                throw "Authorization request expired at $($dcExpiration), please try again."
            }
        }
        catch {
            if ($_.errorDetails.message -match '(AADSTS50059)') {

                "Device code request failed due to 'invalid_request' (AADSTS50059). " +
                'This appears to be a single-tenant app. Retry the command, supplying -TenantId <tenant Domain|Guid>.' |
                Write-Warning
                throw
            }
            elseif ($_.errorDetails.message -match '(AADSTS70011)') {

                "Device code request failed due to 'invalid_scope' (AADSTS70011), which typically means the requested scope(s) do not work with the specified endpoint ('Common' by default). " +
                'Retry the command with either -Endpoint Organizations or -TenantId <tenant Domain|Guid>.' |
                Write-Warning
                throw
            }
            else { throw }
        }
    }

    function Get-RefreshedAcessToken ($Endpoint, $ApplicationId, $RefreshToken, $Scopes) {
        try {
            $trBody = @(
                "client_id=$($ApplicationId)",
                'grant_type=refresh_token',
                "refresh_token=$($RefreshToken.refresh_token)"
            )
            if ($Scopes) { $trBody += "scope=$($Scopes -join ' ')" }

            $trBody = $trBody -join '&'

            $trParams = @{

                Method      = 'POST'
                Uri         = "https://login.microsoftonline.com/$($Endpoint)/oauth2/v2.0/token"
                Body        = $trBody
                ContentType = 'application/x-www-form-urlencoded'
                UserAgent   = "MSGraphPSEssentials/$($PSCmdlet.MyInvocation.MyCommand.Module.Version)"
                ErrorAction = 'Stop'
            }
            $trResponse = Invoke-RestMethod @trParams

            # Output the token request response:
            $trResponse
        }
        catch { throw }
    }
    #endregion Functions

    #region Main
    try {
        switch -Wildcard ($PSCmdlet.ParameterSetName) {

            'ClientCredentials_*' {

                New-AppOnlyAccessToken $TenantId $ApplicationId $Script:Certificate $JWTExpMinutes
            }

            'DeviceCode_*' {

                New-DeviceCodeAccessToken $Script:Endpoint $ApplicationId $Scopes
            }

            'RefreshToken*' {

                Get-RefreshedAcessToken $Script:Endpoint $Script:ApplicationId $Script:RefreshToken $Scopes
            }
        }
    }
    catch {
        if ($_.errorDetails.message -match '(AADSTS50194)') {

            "Application $($ApplicationId) appears to be a single-tenant application.  " +
            'Please supply either -Endpoint:Organizations or -TenantId:<Tenant Id/Guid>' |
            Write-Warning
            throw "$((ConvertFrom-Json $_.errorDetails.message).error_description)"
        }
        else { throw }
    }
    #endregion Main
}

function New-MSGraphRequest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Request,

        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.access_token) { $true } else {

                    throw 'Invalid access token.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken ...'
                }
            }
        )]
        [Object]$AccessToken,

        [Alias('API', 'Version', 'Endpoint')]
        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0',

        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')]
        [string]$Method = 'GET',

        [string]$Body,

        [ValidateSet('Ignore', 'Warn', 'Inquire', 'Continue', 'SilentlyContinue')]
        [string]$nextLinkAction = 'Warn'
    )

    try {
        $RequestParams = @{

            Headers     = @{ Authorization = "Bearer $($AccessToken.access_token)" }
            Uri         = "https://graph.microsoft.com/$($ApiVersion)/$($Request)"
            Method      = $Method
            ContentType = 'application/json'
            UserAgent   = "MSGraphPSEssentials/$($PSCmdlet.MyInvocation.MyCommand.Module.Version)"
            ErrorAction = 'Stop'
        }

        if ($PSBoundParameters.ContainsKey('Body')) {

            if ($Method -notmatch '(POST)|(PATCH)') {

                throw "Body is not allowed when the method is $($Method), only POST or PATCH."
            }
            else { $RequestParams['Body'] = $Body }
        }

        Invoke-RestMethod @RequestParams -OutVariable requestResponse

        if ($requestResponse.'@odata.nextLink') {

            $Script:Continue = $true

            switch ($nextLinkAction) {

                Ignore { $Script:Continue = $false }

                Warn {
                    Write-Warning -Message "There are more results available. Next page: $($requestResponse.'@odata.nextLink')"
                    $Script:Continue = $false
                }
                'Continue' {
                    Write-Information -MessageData 'There are more results available.  Getting the next page...' -InformationAction Continue

                }
                Inquire {
                    switch (
                        $host.UI.PromptForChoice(

                            'There are more results available (i.e. response included @odata.nextLink).',
                            'Get more results?',
                            [ChoiceDescription[]]@('&Yes', 'Yes to &All', '&No'),
                            2
                        )
                    ) {
                        0 { <# Will prompt for choice again if the next response includes another @odata.nextLink.#> }
                        1 { $nextLinkAction = 'SilentlyContinue' }
                        2 { $Script:Continue = $false }
                    }
                }
            }

            if ($Script:Continue) {

                $nextLinkRequestParams = @{

                    AccessToken    = $AccessToken
                    ApiVersion     = $ApiVersion
                    Request        = "$($requestResponse.'@odata.nextLink' -replace 'https://graph.microsoft.com/(v1\.0|beta)/')"
                    nextLinkAction = $nextLinkAction
                    ErrorAction    = 'Stop'
                }

                New-MSGraphRequest @nextLinkRequestParams
            }
        }
    }
    catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 429) {

            "The request was throttled by Microsoft Graph/Azure AD. " +
            "Please wait $($_.Exception.Response.Headers['Retry-After']) seconds before retrying the request, per the Retry-After response header (see `$Error[0].Exception.Response.Headers['Retry-After'])." | Write-Warning
        }
        throw
    }
}

function New-SelfSignedMSGraphApplicationCertificate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Subject,

        [string]$FriendlyName,

        [ValidateScript(
            {
                if (Test-Path -Path $_) { $true } else {

                    throw "An example proper location would be 'cert:\CurrentUser\My'."
                }
            }
        )]
        [string]$CertStoreLocation = 'cert:\CurrentUser\My',

        [datetime]$NotAfter = [datetime]::Now.AddDays(90),

        [ValidateSet('Signature', 'KeyExchange')]
        [string]$KeySpec = 'Signature'
    )

    $NewCertParams = @{

        Subject           = $Subject
        FriendlyName      = $FriendlyName
        CertStoreLocation = $CertStoreLocation
        NotAfter          = $NotAfter
        KeySpec           = $KeySpec
        Provider          = 'Microsoft Enhanced RSA and AES Cryptographic Provider'
        HashAlgorithm     = 'SHA256'
        ErrorAction       = 'Stop'
    }

    try {
        if ($PSVersionTable.PSEdition -eq 'Desktop') { New-SelfSignedCertificate @NewCertParams }
        else {
            # PowerShell Core's PKI module has an issue with allowing the private key to be exportable:
            # https://github.com/PowerShell/PowerShell/issues/12081
            try {
                #Pre-import the PKI module to make the WinPSCompatSession available to Invoke-Command:
                Import-Module -Name PKI -UseWindowsPowerShell -WarningAction:SilentlyContinue

                $WinPSCompatSession = Get-PSSession -Name WinPSCompatSession -ErrorAction:Stop
                if ($WinPSCompatSession) {

                    Invoke-Command -Session $WinPSCompatSession -ScriptBlock { $Global:tmpCertificate = New-SelfSignedCertificate @using:NewCertParams } -ErrorAction:Stop
                    Get-ChildItem -Path "$($CertStoreLocation)\$((Invoke-Command -Session $WinPSCompatSession -ScriptBlock {$tmpCertificate}).Thumbprint)" -ErrorAction:Stop
                }
                else { throw }
            }
            catch { throw "Failed to use WinPSCompatSession to generate valid self-signed certificate. please use Windows PowerShell 5.1 instead." }
        }
    }
    catch { throw }
}

function New-MSGraphPoPToken {
    [CmdletBinding(
        DefaultParameterSetName = 'Certificate'
    )]
    param (
        [Parameter(Mandatory)]
        [Alias('ClientId')]
        [Guid]$ApplicationObjectId,

        [Parameter(
            Mandatory,
            ParameterSetName = 'Certificate',
            HelpMessage = 'E.g. Use $Certificate, where `$Certificate = Get-ChildItem cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [X509Certificate2]$Certificate,

        [Parameter(
            Mandatory,
            ParameterSetName = 'CertificateStorePath',
            HelpMessage = 'E.g. cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317; E.g. cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [ValidateScript(
            {
                if (Test-Path -Path $_) { $true } else {

                    throw "An example proper path would be 'cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'."
                }
            }
        )]
        [string]$CertificateStorePath,

        [ValidateRange(1, 10)]
        [int16]$JWTExpMinutes = 2
    )

    try {
        if ($PSCmdlet.ParameterSetName -eq 'CertificateStorePath') {

            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        else { $Script:Certificate = $Certificate }

        if (-not (Test-SigningCertificate -Certificate $Script:Certificate)) {

            throw "The supplied certificate must use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider', " +
            'and the SHA-256 hashing algorithm.  ' +
            'For best luck, use a certificate generated using New-SelfSignedMSGraphApplicationCertificate.'
        }

        $NowUTC = [datetime]::UtcNow

        $EncodedHeader = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes(
                (ConvertTo-Json -InputObject (
                        @{
                            alg = 'RS256'
                            typ = 'JWT'
                            x5t = ConvertTo-Base64Url -String ([Convert]::ToBase64String($Script:Certificate.GetCertHash()))
                        }
                    )
                )
            )
        )

        $EncodedPayload = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes(
                (ConvertTo-Json -InputObject (
                        @{
                            aud = '00000002-0000-0000-c000-000000000000'
                            iss = $ApplicationObjectId.Guid
                            exp = (Get-Date $NowUTC.AddMinutes($JWTExpMinutes)-UFormat '%s') -replace '\..*'
                            nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
                        }
                    )
                )
            )
        )

        $JWT = (ConvertTo-Base64Url -String $EncodedHeader, $EncodedPayload) -join '.'

        $Signature = ConvertTo-Base64Url -String (
            [Convert]::ToBase64String(
                $Script:Certificate.PrivateKey.SignData(
                    [Text.Encoding]::UTF8.GetBytes($JWT),
                    [HashAlgorithmName]::SHA256,
                    [RSASignaturePadding]::Pkcs1
                )
            )
        )

        $JWT = $JWT + '.' + $Signature

        # Output the token:
        $JWT
    }
    catch { throw }
}

function Add-MSGraphApplicationKeyCredential {
    [CmdletBinding(
        DefaultParameterSetName = 'Certificate'
    )]
    param (
        [Parameter(Mandatory)]
        [Guid]$ApplicationObjectId,

        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.access_token) { $true } else {

                    throw 'Invalid access token.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken ...'
                }
            }
        )]
        [Object]$AccessToken,

        [Parameter(Mandatory)]
        [ValidatePattern('^[-\w]+\.[-\w]+\.[-\w]+$')]
        [string]$PoPToken,

        [Parameter(
            Mandatory,
            ParameterSetName = 'Certificate',
            HelpMessage = 'E.g. Use $Certificate, where `$Certificate = Get-ChildItem cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [X509Certificate2]$Certificate,

        [Parameter(
            Mandatory,
            ParameterSetName = 'CertificateStorePath',
            HelpMessage = 'E.g. cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317; E.g. cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
        )]
        [ValidateScript(
            {
                if (Test-Path -Path $_) { $true } else {

                    throw "An example proper path would be 'cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'."
                }
            }
        )]
        [string]$CertificateStorePath
    )

    try {
        if ($PSCmdlet.ParameterSetName -eq 'CertificateStorePath') {

            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        else { $Script:Certificate = $Certificate }

        if (-not (Test-SigningCertificate -Certificate $Script:Certificate)) {

            throw "The supplied certificate must use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider', " +
            'and the SHA-256 hashing algorithm.  ' +
            'For best luck, use a certificate generated using New-SelfSignedMSGraphApplicationCertificate.'
        }

        $Body = @{

            proof         = $PoPToken
            keyCredential = @{

                type  = "AsymmetricX509Cert"
                usage = "Verify"
                key   = [Convert]::ToBase64String($Script:Certificate.GetRawCertData())
            }
        }

        $AddKeyParams = @{

            AccessToken = $AccessToken
            Method      = 'POST'
            Request     = "applications/$($ApplicationObjectId)/addKey"
            Body        = (ConvertTo-Json $Body)
            ErrorAction = 'Stop'
        }

        New-MSGraphRequest @AddKeyParams
    }
    catch { throw }
}

function Remove-MSGraphApplicationKeyCredential {
    [CmdletBinding(
        DefaultParameterSetName = 'CertificateThumbprint'
    )]
    param (
        [Parameter(Mandatory)]
        [Guid]$ApplicationObjectId,

        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.access_token) { $true } else {

                    throw 'Invalid access token.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken ...'
                }
            }
        )]
        [Object]$AccessToken,

        [Parameter(Mandatory)]
        [ValidatePattern('^[-\w]+\.[-\w]+\.[-\w]+$')]
        [string]$PoPToken,

        [Parameter(
            Mandatory,
            ParameterSetName = 'CertificateThumbprint'
        )]
        [ValidatePattern('^[a-fA-F0-9]{40,40}$')]
        [string]$CertificateThumbprint,

        [Parameter(
            Mandatory,
            ParameterSetName = 'KeyId'
        )]
        [Guid]$KeyId
    )

    try {
        if ($PSCmdlet.ParameterSetName -eq 'CertificateThumbprint') {

            $GetApplicationParams = @{

                AccessToken = $AccessToken
                Request     = "applications/$($ApplicationObjectId)"
                ErrorAction = 'Stop'
            }

            $Application = New-MSGraphRequest @GetApplicationParams

            $MatchingKeyCredentials = $Application.keyCredentials |
            Where-Object { $_.customKeyIdentifier -eq $CertificateThumbprint }

            if ($MatchingKeyCredentials.Count -gt 1) {

                "Multiple keyCredentials matching certificate thumbprint $($CertificateThumbprint) were found.  " +
                "List these with the command below, then re-run this command using -KeyId instead of -CertificateThumbprint:`n" +
                "New-MSGraphRequest -AccessToken <AccessTokenObject> -Request 'applications/$($ApplicationObjectId)' | select -expand keyCredentials" |
                Write-Warning
                break
            }
            elseif ($MatchingKeyCredentials.Count -lt 1) {

                throw "No KeyCredential was found with certificate thumbprint $($CertificateThumbprint)."
            }
            else { $Script:KeyId = $MatchingKeyCredentials.KeyId }
        }
        else { $Script:KeyId = $KeyId.Guid }

        $Body = @{

            keyId = $Script:KeyId
            proof = $PoPToken
        }

        $RemoveKeyParams = @{

            AccessToken = $AccessToken
            Request     = "applications/$($ApplicationObjectId)/removeKey"
            Method      = 'POST'
            Body        = (ConvertTo-Json $Body)
            ErrorAction = 'Stop'
        }

        New-MSGraphRequest @RemoveKeyParams
    }
    catch { throw }
}

function Test-SigningCertificate ([X509Certificate2]$Certificate) {

    if ($PSVersionTable.PSEdition -eq 'Desktop') {

        $Provider = $Certificate.PrivateKey.CspKeyContainerInfo.ProviderName
    }
    else { $Provider = $Certificate.PrivateKey.Key.Provider }

    if (
        $Provider -eq 'Microsoft Enhanced RSA and AES Cryptographic Provider' -and
        $Certificate.SignatureAlgorithm.FriendlyName -match '(sha256)'
    ) {
        $true
    }
    else { $false }
}

function ConvertTo-Base64Url {
    param (
        [ValidatePattern('^(?:[A-Za-z0-9+\/]{4})*(?:[A-Za-z0-9+\/]{2}==|[A-Za-z0-9+\/]{3}=|[A-Za-z0-9+\/]{4})$')]
        [string[]]$String
    )

    $String -replace '\+', '-' -replace '/', '_' -replace '='
}

function ConvertFrom-Base64Url ([string[]]$String) {

    foreach ($s in $String) {

        while ($s.Length % 4) { $s += '=' }
        $s -replace '-', '\+' -replace '_', '/'
    }
}

function ConvertFrom-JWTAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_ -match '^eyJ[-\w]+\.[-\w]+\.[-\w]+$') { $true } else { throw 'Invalid JWT.' }
            }
        )]
        [Object]$JWT
    )

    $Headers, $Payload = ($JWT -split '\.')[0, 1]

    [PSCustomObject]@{
        Headers = ConvertFrom-Json (
            [Text.Encoding]::ASCII.GetString(
                [Convert]::FromBase64String((ConvertFrom-Base64Url $Headers))
            )
        )
        Payload = ConvertFrom-Json(
            [Text.Encoding]::ASCII.GetString(
                [Convert]::FromBase64String((ConvertFrom-Base64Url $Payload))
            )
        )
    }
}

function ConvertFrom-SecureStringToPlainText ([SecureString]$SecureString) {

    [Marshal]::PtrToStringAuto(
        [Marshal]::SecureStringToBSTR($SecureString)
    )
}

function New-RefreshTokenCredential {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Guid]$ApplicationId,

        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.refresh_token) { $true } else {

                    throw 'Invalid token object.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken -Scopes offline_access...'
                }
            }
        )]
        [Object]$TokenObject
    )

    [PSCredential]::new(
        $ApplicationId,
        (ConvertTo-Json $TokenObject | ConvertTo-SecureString -AsPlainText -Force)
    )
}

function Get-AccessTokenExpiration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.access_token) { $true } else {

                    throw 'Invalid token object.  Supply $TokenObject where: $TokenObject = New-MSGraphAccessToken...'
                }
            }
        )]
        [Object]$TokenObject
    )

    $JWT = ConvertFrom-JWTAccessToken -JWT $TokenObject.access_token
    $Epoch = [datetime]"1970-01-01"
    $Now = [datetime]::Now
    $exp = $Epoch.AddSeconds($JWT.Payload.exp).ToLocalTime()

    [PSCustomObject]@{

        IssuedAt_LocalTime       = $Epoch.AddSeconds($JWT.Payload.iat).ToLocalTime()
        NotBefore_LocalTime      = $Epoch.AddSeconds($JWT.Payload.nbf).ToLocalTime()
        ExpirationTime_LocalTime = $exp
        TimeUntilExpiration      = $exp - $Now
        IsExpired                = $Now -gt $exp
    }
}
