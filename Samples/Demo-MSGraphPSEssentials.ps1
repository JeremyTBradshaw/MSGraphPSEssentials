<#
    .Synopsis
    This sample script demonstrates how to use nearly all of the functions that are included with MSGraphPSEssentials.

    .Description
    The MSGraphPSEssentials PS module offers several key functions for interacting with Microsoft Graph API:

        - New-MSGraphAccessToken*
        - New-MSGraphRequest*
        - New-SelfSignedMSGraphApplicationCertificate
        - New-MSGraphPoPToken*
        - Add-MSGraphApplicationKeyCredential*
        - Remove-MSGraphApplicationKeyCredential*
        - ConvertFrom-JWTAccessToken*
        - New-RefreshTokenCredential
        - Get-AccessTokenExpiration*

    (* indicates the function is used in this sample script.)

    The function names make it clear what they do, but it might not be obvious to some that these can be used together
    and orchestrated in such a way that enables your scripts can behave like self-sufficient applications, interacting
    with Microsoft Graph harmoniously.

    **Note: This script is designed to work with Application API permissions (vs Delegated), in app-only, fully-
            unattended fashion.  Another sample script will be added eventually to cover how to work in semi-
            unattended fashion using Refresh Tokens and Delegated permissions.

    At a high-level, this script does the following things:

        - Relies on a PSD1 file to be supplied via parameter entry, which contains the following items:
            -- ApplicationId (for the registered app in Azure AD.  See how to register an app in the
               MSGraphPSEssentials wiki.)
            -- TenantId (again for the registered app).
            -- CertificateStorePath (for a certificate whose public key has been exported and uploaded to the
               registered app in Azure AD.)
            -- AppKeyRolloverDays (the max age in days of the registered app's certificate before the script will
               replace it.)
            -- ThrottlingMaxRetries (how many retries to perform if/when throttled by MS Graph.  Throttling is avoided
               proactively by waiting a minimum of 200ms between MS Graph requests.)
            -- ApiPermissions (a list of API permissions that must be present for the script to work.  The script will
               verify these are present during initialization.)

        - Obtains an initial access token using *New-MSGraphAccessToken*.
            -- Uses *Get-AccessTokenExpiration* (via the script's checkMGTokenExpiringSoon function) before sending
               any other MS Graph requests, ensuring a new access token is obtained before the current one is expired.
               Get-AccessTokenExpiration uses the *ConvertFrom-JWTAccessToken* to determine the expiration time.

        - Uses *Add-MSGraphApplicationKeyCredential* and *Remove-MSGraphApplicationKeyCredential* to "roll the keys"
          (i.e., replace the current certificate) according to the supplied appKeyRolloverDays value.  This is handled
          via the script's rollAppKey function.  The rollAppKey function also makes use of both *New-MSGraphPoPToken*
          and *New-MSGraphRequest* (via newMGRequest) in its process.

        - Avoids throttling proactively, and handles being throttled gracefully if/when it does occur.  This is
          orchestrated in the script's newMGRequest function which makes use of *New-MSGraphRequest*, and the supplied
          value for ThrottlingMaxRetries.  newMGRequest ensures a minimum of 200ms delay between requests which should
          be adequate to avoid throttling based on the published service limits.  Note however, hrottling is dynamic and
          Microsoft may choose to throttle as much as they need, whenever necessary to protect their service during
          heavy usage from sources across the entire service, not just this app or tenant.

    Additional sample scripts will be added to the MSGraphPSEssentials repository, which use this script as the
    starting template and then demonstrate example use cases for the script (e.g. automations which interact with MS
    Graph).

    .Parameter MSGraphParamsPSD1
    Specifies the full/relative file path of the PSD1 file containing the various parameters for use with MS Graph/
    Azure AD.  The following example is what the PSD1 file contents should look like:

    # File .\Demo-MSGraphPSEssentials.MSGraphParams.psd1:
    -----------------------------------------------------
        @{
            tokenParams          = @{

                ApplicationId        = 'f5575a40-f043-4fe6-83db-1428f5df7212'
                TenantId             = 'de19b221-c000-487e-93e4-545b311cfe68'
                CertificateStorePath = 'Cert:\CurrentUser\My\B71F324602735236A538472914CC5DAEBE7D4DF8'
            }

            throttlingMaxRetries = 100

            apiPermissions                 = @{

                listUsers              = '(User\.Read\.All)|(User\.ReadWrite\.All)|(Directory\.Read\.All)|(Directory\.ReadWrite\.All)'
                listSignIns1           = '(AuditLog\.Read\.All)'
                listSignIns2           = '(Directory\.Read\.All)'
            }

            appKeyRolloverDays = 2
            appObjectId = '7767b2dd-a567-4869-a4bb-4686a65a2001' #<--: Needed for /addKey & /removeKey
        }
    -----------------------------------------------------------------------------------
    ** See Demo-mSGraphPSEssentials.MSGraphParams.psd1 for a more detailed example, with comments included.

    .Parameter DisableLogging
    By default, this script will create a folder in the same folder as the script itself, named after the script, but
    appended with "_Logs", and each time the script runs a new log file will be created and written to.  Add additional
    logging as desired following the examples.

    Use this switch to disable logging if you don't want to log.

    .Example
    .\Demo-MSGraphPSEssentials.ps1 -MSGraphParamsPSD1 .\Demo-MSGraphPSEssentials.MSGraphParams.psd1

    This example would roll the keys for the specified application in the PSD1 file, based on the appKeyRolloverDays
    parameter specified in the PSD1.

    .Notes
    (2021-05-31): This is the initial draft and this note will be lifted when this sample script is intended to be
    considered complete.
#>
#Requires -Version 5.1
#Requires -Modules @{ModuleName = 'MSGraphPSEssentials'; Guid = '7394f3f8-a172-4e18-8e40-e41295131e0b'; RequiredVersion = '0.6.0'}
using namespace System.Management.Automation

[CmdletBinding(
    SupportsShouldProcess,
    ConfirmImpact = 'High'
)]
param (
    [Parameter(Mandatory)]
    [ValidateScript(
        {
            if (Test-Path -Path $_ -PathType Leaf) { $true } else {

                throw "Cannot find MS Graph parameters PSD1 file '$($_)'."
            }
        }
    )]
    [System.IO.FileInfo]$MSGraphParamsPSD1,
    [switch]$DisableLogging
)

#======#-----------#
#region# Functions #
#======#-----------#
<#
    Functions are being defined as the first step, so that logging/alerting can be performed during the initialization
    step, in case of issues.
#>

function writeLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$LogName,
        [Parameter(Mandatory)][datetime]$LogDateTime,
        [Parameter(Mandatory)][System.IO.FileInfo]$Folder,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)][string]$Message,
        [ErrorRecord]$ErrorRecord,
        [switch]$DisableLogging,
        [switch]$PassThru
    )

    if (-not $DisableLogging -and -not $WhatIfPreference.IsPresent) {
        try {
            if (-not (Test-Path -Path $Folder)) {

                [void](New-Item -Path $Folder -ItemType Directory -ErrorAction Stop)
            }

            $LogFile = Join-Path -Path $Folder -ChildPath "$($LogName)_$($LogDateTime.ToString('yyyy-MM-dd_HH-mm-ss')).log"
            if (-not (Test-Path $LogFile)) {

                [void](New-Item -Path $LogFile -ItemType:File -ErrorAction Stop)
            }

            $Date = [datetime]::Now.ToString('yyyy-MM-dd hh:mm:ss tt')

            "[ $($Date) ] $($Message)" | Out-File -FilePath $LogFile -Append

            if ($PSBoundParameters.ErrorRecord) {

                # Format the error as it would be displayed in the PS console.
                "[ $($Date) ][Error] $($ErrorRecord.Exception.Message)`r`n" +
                "$($ErrorRecord.InvocationInfo.PositionMessage)`r`n" +
                "`t+ CategoryInfo: $($ErrorRecord.CategoryInfo.Category): " +
                "($($ErrorRecord.CategoryInfo.TargetName):$($ErrorRecord.CategoryInfo.TargetType))" +
                "[$($ErrorRecord.CategoryInfo.Activity)], $($ErrorRecord.CategoryInfo.Reason)`r`n" +
                "`t+ FullyQualifiedErrorId: $($ErrorRecord.FullyQualifiedErrorId)`r`n" |
                Out-File -FilePath $LogFile -Append -ErrorAction Stop
            }

            if ($PassThru) { $Message }
            else { Write-Verbose -Message $Message }
        }
        catch { throw }
    }
    Write-Verbose -Message $Message
}

function newMGAccessToken {
    try { $Script:mgAccessToken = New-MSGraphAccessToken @Script:mgTokenParams -ErrorAction Stop }
    catch { throw }
}

function checkMGTokenExpiringSoon {
    try {
        (Get-AccessTokenExpiration -TokenObject $Script:mgAccessToken).TimeUntilExpiration -le [timespan]"00:05:00"
    }
    catch { throw }
}

function checkMGPermissions {
    try {
        $missingPermissions = @()
        $JWT = ConvertFrom-JWTAccessToken -JWT $Script:mgAccessToken.access_token

        foreach ($ap in $Script:msGraphParams.apiPermissions.Keys) {

            if (-not ($JWT.Payload.Roles -match $Script:msGraphParams.apiPermissions[$ap])) {

                $missingPermissions += $ap
            }
        }
        if ($missingPermissions.Count -ge 1) { $missingPermissions }
    }
    catch { throw }
}

function newMGRequest {
    [CmdletBinding()]
    param(
        [string]$Request,
        [string]$Method = 'GET',
        [hashtable]$Body,
        [string]$ApiVersion = 'v0.1',
        [string]$nextLinkAction = 'SilentlyContinue'
    )

    try {
        # If this is a retry and max retries have been reached, reset counter and bailout:
        if ($Script:mgThrottledCounter -gt $Script:msGraphParams.throttlingMaxRetries) {

            $Script:mgThrottledCounter = 0
            throw "Maximum retries ($($Script:msGraphParams.throttlingMaxRetries)) reached, stopping retry sequence now."
        }

        # Proactively delay by 200ms to avoid being throttled in the first place:
        if ($Script:mgStopwatch.Elapsed.Milliseconds -lt 200) {

            Start-Sleep -Milliseconds 200
            $Script:mgStopwatch.Restart()
        }

        # If necessary, get a new access token:
        if (checkMGTokenExpiringSoon) { newMGAccessToken }

        $reqParams = @{
            AccessToken    = $Script:mgAccessToken
            Request        = $Request
            nextLinkAction = $nextLinkAction
            ErrorAction    = 'Stop'
        }
        if ($PSBoundParameters.ContainsKey('Method')) { $reqParams['Method'] = $Method }
        if ($PSBoundParameters.ContainsKey('Body')) { $reqParams['Body'] = ConvertTo-Json $Body -Depth 100 }
        if ($PSBoundParameters.ContainsKey('ApiVersion')) { $reqParams['ApiVersion'] = $ApiVersion }

        # Write-Debug -Message 'Inspect $reqParams'
        New-MSGraphRequest @reqParams

        # Reset throttled counter:
        $Script:mgThrottledCounter = 0
    }
    catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 429) {

            # Increment the counter.
            $Script:mgThrottledCounter++

            "Throttled by MS Graph during request: $($Request).  " +
            "Waiting greator of 10 or {$($_.Exception.Response.Headers['Retry-After'])} ('Retry-After' response header) " +
            "seconds before performing the next retry (max retries: $($Script:msGraphParams.throttlingMaxRetries))." |
            writeLog @writeLogParams -PassThru |
            Write-Warning
            Start-Sleep -Seconds ([math]::Max(10, [int]$_.Exception.Response.Headers['Retry-After']))
            newMGRequest @reqParams
        }
        else {
            # Reset throttled counter and bailout:
            $Script:mgThrottledCounter = 0
            throw
        }
    }
}

function rollAppKey {
    try {
        $cert = Get-ChildItem -Path $Script:mgTokenParams['CertificateStorePath'] -ErrorAction Stop

        if ($cert.NotBefore -lt [datetime]::Now.AddDays(-$Script:msGraphParams['appKeyRolloverDays'])) {

            writeLog @writeLogParams -Message "Certificate $($cert.Thumbprint) is $($Script:msGraphParams['appKeyRolloverDays']) or more days old so will be replaced both locally and in Azure AD via MS Graph /addKey and /removeKey." -PassThru |
            Write-Warning

            # Create and add a new self-signed cert to the app in Azure AD (stores locally in Cert:\CurrentUser\My):
            try {
                # Create a Proof of Possession JWT of the current certificate.
                $currentCertPoPToken = New-MSGraphPoPToken -ApplicationObjectId $Script:msGraphParams['appObjectId'] -Certificate $cert -JWTExpMinutes 5 -ErrorAction Stop

                # Create and add new certificate to the Azure AD application:
                $newCert = New-SelfSignedMSGraphApplicationCertificate -Subject $PSCmdlet.MyInvocation.MyCommand.Name -ErrorAction Stop

                if (checkMGTokenExpiringSoon) { newMGAccessToken }
                Add-MSGraphApplicationKeyCredential -AccessToken $Script:mgAccessToken -ApplicationObjectId $Script:msGraphParams['appObjectId'] -PoPToken $currentCertPoPToken -Certificate $newCert -ErrorAction Stop
            }
            catch {
                writeLog @writeLogParams -Message "Failed to upload new certificate ($($newCert.Thumbprint)) to Azure AD (via MS Graph: /addKey))." -PassThru |
                Write-Warning
                throw
            }

            # Verify the app is ready with the new certificate by requesting a new access token using it:
            $tmpSW = [System.Diagnostics.Stopwatch]::StartNew()
            do {
                try {
                    $newMGTokenParams = @{

                        ApplicationId = $Script:mgTokenParams['ApplicationId']
                        TenantId      = $Script:mgTokenParams['TenantId']
                        Certificate   = $newCert
                        ErrorAction   = 'Stop'
                    }
                    if ($tmpSW.Elapsed.Seconds -ge 30) { $newMGToken = New-MSGraphAccessToken @newMGTokenParams }
                }
                catch {
                    if ($_.ErrorDetails.Message) {
                        $errorMessage = ConvertFrom-Json $_.ErrorDetails.Message
                        if ($errorMessage.error_description -match '(AADSTS700027)') {

                            writeLog @writeLogParams -Message "New certificate $($newCert.Thumbprint) is still not available for use with the app in Azure AD.  Checking again in 30 seconds." -PassThru |
                            Write-Warning
                            Start-Sleep -Seconds 30
                        }
                        else { throw }
                    }
                    else { throw }
                }
            }
            until ($newMGToken -or ($tmpSW.Elapsed.Minutes -ge 10))

            # If successful to this point, update the Params.psd1 file with the new certificate's thumbprint:
            try {
                $msGraphParamsPSD1Content = Get-Content -Path $Script:msGraphParamsPSD1 -ErrorAction Stop
                Set-Content -Path $Script:msGraphParamsPSD1 -Value $msGraphParamsPSD1Content.Replace("$($cert.Thumbprint)", "$($newCert.Thumbprint)") -ErrorAction Stop

                # Re-import the updated Params.psd1 which now has the new thumbprint:
                $Script:msGraphParams = Import-PowerShellDataFile $Script:msGraphParamsPSD1 -ErrorAction Stop

                # Redefine $mgTokenParams:
                $Script:mgTokenParams = $Script:msGraphParams.tokenParams

                # Now get a new access token using the usual approach (to make sure it works):
                $tmpSW.Restart()
                $_success = $false
                do {
                    try {
                        if ($tmpSW.Elapsed.Seconds -ge 30) {

                            newMGAccessToken
                            $_success = $true
                        }
                    }
                    catch {
                        if ($_.ErrorDetails.Message) {
                            $errorMessage = ConvertFrom-Json $_.ErrorDetails.Message
                            if ($errorMessage.error_description -match '(AADSTS700027)') {

                                writeLog @writeLogParams -Message "New certificate is unavailable for use again (MS Graph / AAD replication issue) with the app in Azure AD.  Trying again in 30 seconds." -PassThru |
                                Write-Warning
                                Start-Sleep -Seconds 30
                            }
                            else { throw }
                        }
                        else { throw }
                    }
                }
                until ($_success -or ($tmpSW.Elapsed.Minutes -ge 10))
            }
            catch {
                writeLog @writeLogParams -Message "function rollAppKey: Failed to update Params PSD1 file with new certificate thumbprint, or to re-import the updated PSD1, or to successfully obtain an access token using the updated/re-imported PSD1." -ErrorRecord $_ -PassThru | Write-Warning

                Write-Debug -Message "rollAppKey: try {(update/re-import/test PSD1)} catch {*}"
                throw
            }

            # Finally, remove the old certificate, first from the local store, then from Azure AD:
            try {
                Remove-Item -Path "$($cert.PSParentPath)\$($cert.Thumbprint)" -ErrorAction Stop

                # Get a Proof of Possession of the new cert:
                $newCertPoPToken = New-MSGraphPoPToken -ApplicationObjectId $Script:msGraphParams['appObjectId'] -Certificate $newCert -JWTExpMinutes 5 -ErrorAction Stop

                # Foreach ensure's all instances of the old cert from Azure AD, in case the same cert had been uploaded manually outside of this script.
                $app = newMGRequest -Request "/applications/$($Script:msGraphParams['appObjectId'])"
                foreach ($kc in ($app.keyCredentials |
                        Where-Object { $_.customKeyIdentifier -eq $cert.Thumbprint })) {

                    Remove-MSGraphApplicationKeyCredential -AccessToken $Script:mgAccessToken -ApplicationObjectId $Script:msGraphParams['appObjectId'] -PoPToken $newCertPoPToken -KeyId $kc.keyId -ErrorAction Stop
                }
            }
            catch {
                writeLog @writeLogParams -Message "Failed to remove the old certificate ($($cert.Thumbprint)) from either or both the local store and Azure AD (via MS Graph: /removeKey))." -ErrorRecord $_ -PassThru |
                Write-Warning
                throw
            }

            writeLog @writeLogParams -Message "Replaced old certificate ($($cert.Thumbprint)) with new certificate ($($newCert.Thumbprint)) both locally and in the app in Azure AD."
        }
        else {
            writeLog @writeLogParams -Message "Certificate $($cert.Thumbprint) is not $($Script:msGraphParams['appKeyRolloverDays']) or more days old so will not be replaced at this time."
        }
    }
    catch {
        Write-Debug -Message 'function: rollAppKey { try {} catch {*} }'
        throw
    }
}

#=========#-----------#
#endregion# Functions #
#=========#-----------#

try {

    #======#-----------------------------------#
    #region# Initialization / Common Variables #
    #======#-----------------------------------#

    # 1. Setup writeLog splat and test writeLog:
    $Script:dtNow = [datetime]::Now
    $Script:writeLogParams = @{

        LogName        = "$($PSCmdlet.MyInvocation.MyCommand.Name -replace '\.ps1')"
        LogDateTime    = $dtNow
        Folder         = "$($PSCmdlet.MyInvocation.MyCommand.Source -replace '\.ps1')_Logs"
        DisableLogging = $DisableLogging
        ErrorAction    = 'Stop'
    }
    writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand.Name) - Start"
    writeLog @writeLogParams -Message "MyCommand: $($PSCmdlet.MyInvocation.Line)"

    # 2. Import MS Graph / AAD app parameters:
    try {
        $Script:msGraphParams = Import-PowerShellDataFile $MSGraphParamsPSD1 -ErrorAction Stop
        writeLog @writeLogParams -Message "Imported MS Graph parameters from $($MSGraphParamsPSD1)."
    }
    catch {
        writeLog @writeLogParams -Message 'Failed to import MS Graph parameters PSD1.' -PassThru | Write-Warning
        throw
    }

    # 3. Get the initial access token for use with MS Graph:
    try {
        $Script:mgTokenParams = $Script:MSGraphParams.TokenParams
        newMGAccessToken #<--: Stores it in $Script:mgAccessToken.
    }
    catch {
        writeLog @writeLogParams -Message 'Failed to get an access token for use with MS Graph.' -PassThru | Write-Warning
        throw
    }

    # 4. Verify AAD app's MS Graph permissions:
    try {
        $missingPermissions = checkMGPermissions
        if ($missingPermissions.Count -ge 1) {

            throw "The access token is missing one or more required permissions.  Refer to the 'apiPermissions' section of the MS Graph parameters PSD1 file."
        }
    }
    catch {
        writeLog @writeLogParams -Message 'Failed while checking the access token for required API permissions (via function checkMGPermissions).' -PassThru | Write-Warning
        throw
    }

    # 5. Start a stopwatch to control MS Graph request frequency to avoid throttling.  If throttled, newMGRequest has built-in retry logic as well.
    $Script:mgStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # 6. Initialize a counter for throttling occurrences:
    $Script:mgThrottledCounter = 0

    # 7. Roll the keyCredentials on the AAD app (if current cert is old enough to merit being replaced):
    try { rollAppKey }
    catch {
        writeLog @writeLogParams -Message "Failed to replace the AAD app's certificate." -PassThru | Write-Warning
        throw
    }

    #=========#-----------------------------------#
    #endregion# Initialization / Common Variables #
    #=========#-----------------------------------#



    #======#------------------------------#
    #region# Stay Organized with Regions! #
    #======#------------------------------#

    #=========#------------------------------#
    #endregion# Stay Organized with Regions! #
    #=========#------------------------------#
}
catch {
    writeLog @writeLogParams -Message 'Ending script prematurely.' -ErrorRecord $_ -PassThru | Write-Warning
    throw
}
finally {
    writeLog @writeLogParams -Message "$($PSCmdlet.MyInvocation.MyCommand.Name) - End"
}
