#Requires -Version 5.1
using namespace System
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates
using namespace System.Management.Automation.Host

<#
    v0.0.0 - Pre-release / unpublished (2020-12-07):
    
      - This is a direct copy/paste from MSGraphAppOnlyEssentials at this time.  So it works all the same.
      - To see what I'm working on for the upgraded version of the New-MSGraphAccessToken function:
        https://github.com/JeremyTBradshaw/PowerShell/blob/main/Dev/MSGraphPlayground.psm1
#>

function New-MSGraphAccessToken {

    [CmdletBinding(
        DefaultParameterSetName = 'Certificate'
    )]
    param (
        [Parameter(Mandatory)]
        [Alias('Tenant', 'TenantName', 'TenantDomain', 'TenantDomainName')]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [Alias('ClientId')]
        [Guid]$ApplicationId,

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

    if ($PSCmdlet.ParameterSetName -eq 'CertificateStorePath') {

        try {
            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        catch { throw $_ }
    }
    else { $Script:Certificate = $Certificate }

    if (-not (Test-CertificateProvider -Certificate $Script:Certificate)) {

        $ErrorMessage = "The supplied certificate does not use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'.  " +
        "For best luck, use a certificate generated using New-SelfSignedMSGraphApplicationCertificate."

        throw $ErrorMessage
    }

    $NowUTC = [datetime]::UtcNow

    $JWTHeader = @{

        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-Base64UrlFriendly -String ([Convert]::ToBase64String($Script:Certificate.GetCertHash()))
    }

    $JWTClaims = @{

        aud = "https://login.microsoftonline.com/$TenantId/oauth2/token"
        exp = (Get-Date $NowUTC.AddMinutes($JWTExpMinutes) -UFormat '%s') -replace '\..*'
        iss = $ApplicationId.Guid
        jti = [Guid]::NewGuid()
        nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
        sub = $ApplicationId.Guid
    }

    $EncodedJWTHeader = [Convert]::ToBase64String(
        
        [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $JWTHeader))
    )
    
    $EncodedJWTClaims = [Convert]::ToBase64String(
        
        [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $JWTClaims))
    )

    $JWT = ConvertTo-Base64UrlFriendly -String ($EncodedJWTHeader + '.' + $EncodedJWTClaims)

    $Signature = ConvertTo-Base64UrlFriendly -String ([Convert]::ToBase64String(
        
            $Script:Certificate.PrivateKey.SignData(
            
                [Text.Encoding]::UTF8.GetBytes($JWT),
                [HashAlgorithmName]::SHA256,
                [RSASignaturePadding]::Pkcs1
            )
        )
    )

    $JWT = $JWT + '.' + $Signature

    $Body = @{

        client_id             = $ApplicationId
        client_assertion      = $JWT
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        scope                 = 'https://graph.microsoft.com/.default'
        grant_type            = "client_credentials"
    }

    $TokenRequestParams = @{

        Method      = 'POST'
        Uri         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        Body        = $Body
        Headers     = @{ Authorization = "Bearer $($JWT)" }
        ContentType = 'application/x-www-form-urlencoded'
        ErrorAction = 'Stop'
    }

    try {
        Invoke-RestMethod @TokenRequestParams
    }
    catch { throw $_ }
}

function New-MSGraphRequest {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Alias('Query')]
        [string]$Request,

        [Parameter(Mandatory)]
        [ValidateScript(
            {
                if ($_.token_type -eq 'Bearer' -and $_.access_token -match '^[-\w]+\.[-\w]+\.[-\w]+$') { $true } else {

                    throw 'Invalid access token.  For best results, supply $AccessToken where: $AccessToken = New-MSGraphAccessToken ...'
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

        [ValidateSet('Warn', 'Inquire', 'Continue', 'SilentlyContinue')]
        [string]$nextLinkAction = 'Warn'
    )

    $RequestParams = @{

        Headers     = @{ Authorization = "Bearer $($AccessToken.access_token)" }
        Uri         = "https://graph.microsoft.com/$($ApiVersion)/$($Request)"
        Method      = $Method
        ContentType = 'application/json'
        ErrorAction = 'Stop'
    }

    if ($PSBoundParameters.ContainsKey('Body')) {
        
        if ($Method -notmatch '(POST)|(PATCH)') {

            throw "Body is not allowed when the method is $($Method), only POST or PATCH."
        }
        else { $RequestParams['Body'] = $Body }
    }

    try {
        Invoke-RestMethod @RequestParams -OutVariable requestResponse
    }
    catch { throw $_ }

    if ($requestResponse.'@odata.nextLink') {

        $Script:Continue = $true

        switch ($nextLinkAction) {

            Warn {
                Write-Warning -Message "There are more results available. Next page: $($requestResponse.'@odata.nextLink')"
                $Script:Continue = $false
            }
            Continue {
                Write-Information -MessageData 'There are more results available.  Getting the next page' -InformationAction Continue

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
                    0 {} # Will prompt for choice again if the next response includes another @odata.nextLink.
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

            try {
                New-MSGraphRequest @nextLinkRequestParams
            }
            catch { throw $_ }
        }
    }
}
New-Alias -Name New-MSGraphQuery -Value New-MSGraphRequest

function New-SelfSignedMSGraphApplicationCertificate {
    [CmdletBinding()]
    param (
        [ValidatePattern('(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+[a-zA-Z]{2,63}$)')]
        [string]$DnsName,    

        [Parameter(Mandatory)]
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

        DnsName           = $DnsName
        FriendlyName      = $FriendlyName
        CertStoreLocation = $CertStoreLocation
        NotAfter          = $NotAfter
        KeyExportPolicy   = 'Exportable'
        KeySpec           = $KeySpec
        Provider          = 'Microsoft Enhanced RSA and AES Cryptographic Provider'
        HashAlgorithm     = 'SHA256'
        ErrorAction       = 'Stop'
    }

    try {
        New-SelfSignedCertificate @NewCertParams
    }
    catch { throw $_ }
}
New-Alias -Name 'New-SelfSignedAzureADAppRegistrationCertificate' -Value New-SelfSignedMSGraphApplicationCertificate

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

    if ($PSCmdlet.ParameterSetName -eq 'CertificateStorePath') {

        try {
            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        catch { throw $_ }
    }
    else { $Script:Certificate = $Certificate }

    if (-not (Test-CertificateProvider -Certificate $Script:Certificate)) {

        $ErrorMessage = "The supplied certificate does not use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'.  " +
        "For best luck, use a certificate generated using New-SelfSignedMSGraphApplicationCertificate."

        throw $ErrorMessage
    }

    $NowUTC = [datetime]::UtcNow

    $JWTHeader = @{

        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-Base64UrlFriendly -String ([Convert]::ToBase64String($Script:Certificate.GetCertHash()))
    }

    $JWTClaims = @{

        aud = '00000002-0000-0000-c000-000000000000'
        iss = $ApplicationObjectId.Guid
        exp = (Get-Date $NowUTC.AddMinutes($JWTExpMinutes) -UFormat '%s') -replace '\..*'
        nbf = (Get-Date $NowUTC -UFormat '%s') -replace '\..*'
    }

    $EncodedJWTHeader = [Convert]::ToBase64String(
        
        [Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
    )
    
    $EncodedJWTClaims = [Convert]::ToBase64String(
        
        [Text.Encoding]::UTF8.GetBytes(($JWTClaims | ConvertTo-Json))
    )

    $JWT = ConvertTo-Base64UrlFriendly ($EncodedJWTHeader + '.' + $EncodedJWTClaims)

    $Signature = ConvertTo-Base64UrlFriendly ([Convert]::ToBase64String(
        
            $Script:Certificate.PrivateKey.SignData(
            
                [Text.Encoding]::UTF8.GetBytes($JWT),
                [HashAlgorithmName]::SHA256,
                [RSASignaturePadding]::Pkcs1
            )
        )
    )

    $JWT = $JWT + '.' + $Signature

    $JWT
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
                if ($_.token_type -eq 'Bearer' -and $_.access_token -match '^[-\w]+\.[-\w]+\.[-\w]+$') { $true } else {

                    throw 'Invalid access token.  For best results, supply $AccessToken where: $AccessToken = New-MSGraphAccessToken ...'
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

    if ($PSCmdlet.ParameterSetName -eq 'CertificateStorePath') {

        try {
            $Script:Certificate = Get-ChildItem -Path $CertificateStorePath -ErrorAction Stop
        }
        catch { throw $_ }
    }
    else { $Script:Certificate = $Certificate }

    if (-not (Test-CertificateProvider -Certificate $Script:Certificate)) {

        $ErrorMessage = "The supplied certificate does not use the provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'.  " +
        "For best luck, use a certificate generated using New-SelfSignedMSGraphApplicationCertificate."

        throw $ErrorMessage
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
    
    try {
        New-MSGraphRequest @AddKeyParams
    }
    catch { throw $_ }
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
                if ($_.token_type -eq 'Bearer' -and $_.access_token -match '^[-\w]+\.[-\w]+\.[-\w]+$') { $true } else {

                    throw 'Invalid access token.  For best results, supply $AccessToken where: $AccessToken = New-MSGraphAccessToken ...'
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

    if ($PSCmdlet.ParameterSetName -eq 'CertificateThumbprint') {

        $ListKeysParams = @{

            AccessToken = $AccessToken
            Request     = "applications/$($ApplicationObjectId)"
            ErrorAction = 'Stop'
        }

        try {
            $KeyCredentials = New-MSGraphRequest @ListKeysParams
        }
        catch { throw $_ }

        $Script:KeyId = ($KeyCredentials.KeyCredentials |
            Where-Object { $_.customKeyIdentifier -eq $CertificateThumbprint }).KeyId

        if (-not $Script:KeyId) {

            throw "No KeyCredential was found with certificate thumbprint $($CertificateThumbprint)."
        }
    }
    else {
        $Script:KeyId = $KeyId.Guid
    }

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

    try {
        New-MSGraphRequest @RemoveKeyParams
    }
    catch { throw $_ }
}

function ConvertTo-Base64UrlFriendly ([string]$String) {

    $String -replace '\+', '-' -replace '/', '_' -replace '='
}

function Test-CertificateProvider ([X509Certificate2]$Certificate) {

    if ($PSVersionTable.PSEdition -eq 'Desktop') {

        $certProvider = $Certificate.PrivateKey.CspKeyContainerInfo.ProviderName
    }
    else { $certProvider = $Certificate.PrivateKey.Key.Provider }

    if ($certProvider -eq 'Microsoft Enhanced RSA and AES Cryptographic Provider') {

        $true
    }
    else { $false }
}
