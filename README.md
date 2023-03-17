# ⚠ End-of-Life Notice (2022-04-20, 2023-03-17)

FYI, at this point in time, Microsoft have delivered a proper set of PowerShell modules for interacting with Microsoft Graph.  Thanks to that, I will not be working on MSGraphPSEssentials any longer, unless some valid reason to comes up.  I suggest moving over to the official modules considering the fact they are backed and supported by Microsoft.  Where they fall short in my opinion is that you need to learn the ins and outs of both MS Graph and their new modules, each having their own entirely different set of syntaxes.  So twice the learning time required to take part, whereas with MSGraphPSEssentials, you only need to learn the MS Graph reference material and plug your requests into the New-MSGraphRequest function.  At the same time, the [Invoke-MgGraphRequest](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0#using-invoke-mggraphrequest) function that comes with the official module is essentially a like-for-like replacement of MSGraphPSEssentials' New-MSGraphRequest, and then some.  So this module is officially obsoleted by Microsoft's official offering..

If you tried and/or enjoyed MSGraphPSEssentials, thank you very much, it was a fun side project for me.  I won't be de-listing MSGraphPSEssentials from Powershell Gallery for now, and will also leave this repo here too.  It still works just like always so definitely feel free to keep using it (or clone it and take it further as you see fit).

* [Microsoft Graph PowerShell overview](https://docs.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-1.0)
* [Install the Microsoft Graph PowerShell SDK](https://docs.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0)

---

# MSGraphPSEssentials
A collection of functions enabling easier consumption of Microsoft Graph using just PowerShell (Desktop/Core).

[Available on PowerShell Gallery](https://www.powershellgallery.com/packages/MSGraphPSEssentials)

This module provides some essential functions for working with Microsoft Graph using PowerShell 5.1 and newer, in App-Only (i.e. 100% unattended) fashion using certificate credentials, or in Delegated (i.e. interactive / partially unattended) fashion, using Device Code flow or Refresh Tokens.  Obtain access tokens, perform nearly any MS Graph request you can think of, and roll your own keys using the addKey and removeKey application resource type methods.  The only thing left to do manually is to setup an App Registration in Azure AD (configure API permissions and upload the initial certificate for App-Only scenarios).  Use the [MS Graph API (v1.0 / beta) reference material](https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0) to figure out your requests and query syntax, then start sending them with `New-MSGraphRequest`.  Options for handling paging are provided, allowing 'nextLink' responses to be handled however you prefer (automatic, inquire, warn, etc.).  Best of all - no DLL's; no big files at all.  In fact, only about 30-40KB.

For Delegated scenarios, if you'd rather forego the hassles of creating an App Registration in Azure AD, feel free to use my MSGraphPSEssentials app from my Azure AD tenant.  It is setup as a multi-tenant app with support for organizational (i.e. Work/School) accounts and personal Microsoft accounts (MSA's) alike.  Just supply `-ApplicationId 0c26a905-7c94-4296-aeb4-b8925cb7e036` when using `New-MSGraphAccessToken`.

📰**FYI (as of v0.5.5)** `New-MSGraphAccessToken` has been updated with a new parameter, `-ExoEwsAppOnlyScope`, which brings support for EWS app-only OAuth to this module.  Along with that came the deprecation of the **EwsOAuthAppOnlyEssentials** module.

💡 **In addition to to the documentation below, a growing list of samples can be found in the [wiki](https://github.com/JeremyTBradshaw/MSGraphPSEssentials/wiki).  A sample script has also be added (see /samples).**

---

## (Essential) Functions

### New-MSGraphAccessToken

This function is used to get an access token.  Access tokens generally last for 1 hour, so keep this in mind in long-running scripts, and also make use of `Get-AccessTokenExpiration` (see notes farther down the page below) to remove any guesswork.

ℹ **Note:** For a less-powerful (security-wise) alternative to app-only access tokens, you may find that you can accomplish enough unattended'ness by interactively requesting the initial access token using Device Code flow, and then programmatically refreshing the access token before the refresh token expires.  As long as the refresh token is never revoked, this could go on forever, unattended.

Parameters | Parameter Set | Description
---------: | :-----------: | :-----------
ApplicationId | All parameter sets | The app's ApplicationId (a.k.a. ClientId)
TenantId | All parameter sets | The directory/tenant ID (Guid or ***.onmicrosoft.com domain)
Certificate<br />(Option 1) | ClientCredentials_Certificate | Use `$Certificate`, where `$Certificate = Get-ChildItem Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317`
CertificateStorePath<br/>(Option 2) | ClientCredentials_CertificateStorePath | E.g. 'Cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
**_(New in v0.5.5)_**<br />ExoEwsAppOnlyScope | ClientCredentials_* | Uses the Exchange Online scope (https://outlook.office365.com/.default), for use with EWS OAuth app-only authentication.
JWTExpMinutes | ClientCredentials_* | In case of a poorly-synced clock, use this to adjust the expiry of the JWT that is the client assertion sent in the request.  Max. value is 10.
Endpoint | DeviceCode_\*, RefreshToken_* | **Common**, Consumers, Organizations.  Common (default) should suffice for most scenarios.  Scopes will dictate when either of the other options must be used instead.
Scopes | DeviceCode_\*, RefreshToken_* | One or more delegated permissions (e.g. Ews.AccessAsUser.All, mail.Send, offline_access).  Refer to MS Graph reference pages for required permissions.  Those permissions are exaclty what to specify here.
RefreshToken | RefreshToken_* | Use `$TokenObject` where `$TokenObject = New-MSGraphAccessToken -Scopes ..., offline_access`.  The 'offline_access' scope is what tells Azure AD to also hand back a refresh token when issuing an access token.
RefreshTokenCredential | RefreshTokenCredential_* | Supply a PSCredential object that was created by the function `New-RefreshTokenCredential`.  Refer to the example from `New-RefreshTokenCredential`'s section (further down in this readme) to see how to do this.

**Example 1**

Get an app-only access token (with `-Certificate`):

```powershell
$TokenObjectParams = @{

    ApplicationId = '4ba21eca-462c-46cd-b296-9467232638a4'
    TenantId      = 'c7bdcf5c-7a22-44f0-8240-146ababc5858'
    Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\F046351F8B17FA1755F4A567C175BEA1FC86A1EC'
}
$TokenObject = New-MSGraphAccessToken @TokenObjectParams
```
**Example 2**

Get an app-only access token (with `-CertificateStorePath`):

```powershell
$TokenObjectParams = @{

    ApplicationId        = '4ba21eca-462c-46cd-b296-9467232638a4'
    TenantId             = 'c7bdcf5c-7a22-44f0-8240-146ababc5858'
    CertificateStorePath = 'Cert:\CurrentUser\My\F046351F8B17FA1755F4A567C175BEA1FC86A1EC'
}
$TokenObject = New-MSGraphAccessToken @TokenObjectParams
```

**Example 3**

Get an access token using the Device Code flow:

```powershell
$TokenObject = Get-MSGraphAccessToken -ApplicationId '0c26a905-7c94-4296-aeb4-b8925cb7e036' -Scopes mail.send, offline_access

# Follow the instructions, which will look similar to this:
Authorization started (expires at 2020-12-19 03:01:44 PM)
To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code FB228GRH7 to authenticate.
[D] Done  [?] Help (default is "D"):
```

**Example 4**

Refresh an access token (which was obtained using `New-MSGraphAccessToken` and including 'offline_access' in the `-Scopes` parameter):

```powershell
# The original access token request needs to include the 'offline_access' scope to be given a refresh token along with the access token:
$TokenObject = New-MSGraphAccessToken -ApplicationId '0c26a905-7c94-4296-aeb4-b8925cb7e036' -Scopes mail.readwrite, mail.send, offline_access

# To refresh the above access token:
$NewTokenObject = New-MSGraphAccessToken -ApplicationId '0c26a905-7c94-4296-aeb4-b8925cb7e036' -RefreshToken $TokenObject
```

### New-MSGraphRequest

Use this function to send requests/queries to Microsoft Graph.

Parameters | Description
---------: | :-----------
ApiVersion | **v1.0**, beta.  Use either of these as needed for the task at hand.  Stick to v1.0 for production scripts!
Method | **GET**, POST, PATCH, PUT, DELETE.  Follow the reference articles' instructions and see the examples below.
AccessToken | Use `$TokenObject` where `$TokenObject = New-MSGraphAccessToken ...`
Request | Use the **HTTP request** shown in MS Graph reference articles.  Include the query parameters, E.g. ``"auditLogs/signIns?&`$filter=userId eq '86d691b0-8c2b-4adf-bdcd-d9f0ab7b6183' and status/errorCode eq 0"``.  The leading forward slash (e.g. "<b>/</b>users") can be include or omitted.
Body<br/>(Optional) | Use the **request body** as shown in MS Graph reference articles, supplied as a hashtable (the function converts it to JSON automatically).
nextLinkAction | **Warn**, Ignore, Inquire, Continue, SilentlyContinue.  When multiple pages of results are available (i.e. more than 100 users), use this to decide how and whether to keep getting the next page of results.  When possible, it is best to use the $top OData query option to set a larger page size to avoid an unnecessary abundance of nextLink requests.

**Example 1**

```powershell
# Get all users:
New-MSGraphRequest -AccessToken $TokenObject -Request users -nextLinkAction SilentlyContinue
```

**Example 2**

```powershell
# Add a new item to a SharePoint list:
$listItem = @{
    fields = @{
        Make   = 'Honda'
        Model  = 'Accord'
        Engine = 'V6'
        Trim   = 'EX'
        Year   = '2006'
    }
}

$requestParams = @{
    AccessToken = $TokenObject
    Method      = 'POST'
    Request     = 'sites/1098c395-c09a-4c84-8e5e-2f454f318667/lists/2d6ac42c-f4d9-4ba2-9897-220fc87bec8f/items'
    Body        = $listItem
}

New-MSGraphRequest @requestParams
```

**Example 3**

```powershell
# Update and existing item on a SharePoint list:
$listItem = @{
    Style      = 'Sedan'
    Condition  = 'Mint'
    Kilometers = 1000000
}

$requestParams = @{
    AccessToken = $TokenObject
    Method      = 'PATCH'
    Request     = 'sites/1098c395-c09a-4c84-8e5e-2f454f318667/lists/2d6ac42c-f4d9-4ba2-9897-220fc87bec8f/items/23'
    Body        = $listItem
}

New-MSGraphRequest @requestParams
```

### New-SelfSignedMSGraphApplicationCertificate

This function is simply a wrapper for New-SelfSignedCertificate with some base settings to ensure the certificate will work with the other functions and Microsoft Graph in general.  It uses the Microsoft Enhanced RSA and AES Cryptographic Provider ensuring the certificate will be able to do everything it might need with the other functions.  The other functions also insist on this provider.

Parameters | Description
---------: | :----------
Subject | String of choice, not tested thoroughly so don't get crazy with this.  E.g. 'jb365 MS Graph Automation'
FriendlyName | Another string.  Will match Subject if nothing is supplied.
CertStoreLocation | Maps directly to the same parameter of New-SelfSignedCertificate.  Default is 'cert:\CurrentUser\My'.  Any valid location where write access is available will work (e.g. 'cert:\LocalMachine\My', when scheduling tasks using local SYSTEM account).
NotAfter | Default is 90 days.  Supply a [datetime] like this `(Get-Date).AddDays(7)`.  The shorter the better, because applications can add new certificates for themselves using the addKey method, so we can easily roll these often, programmatically.
KeySpec | **Signature**, KeyExchange.  Recommendation: don't change this unless there is a reason.

**Example 1**

```powershell
New-SelfSignedMSGraphApplicationCertificate -Subject "jb365 MS Graph Automation" -NotAfter (Get-Date).AddDays(7)
```

**Example 2**

```powershell
$date = [datetime]::Now.ToString('yyyyMMdd')
$newCertParams = @{
    Subject           = "jb365 MS Graph Automation"
    FriendlyName      = "jb365 MS Graph Automation ($($date))"
    CertStoreLocation = 'Cert:\LocalMachine\My'
    NotAfter          = (Get-Date).AddDays(7)
}
New-SelfSignedMSGraphApplicationCertificate @newCertParams
```

### New-MSGraphPoPToken

This function generates a Proof of Possession JWT (JSON Web Token) for use with the addKey/removeKey MS Graph application resource type methods (in this case the `Add-MSGraphApplicationKeyCredential` and `Remove-MSGraphApplicationKeyCredential` functions).  When adding a new key (i.e., certificate) to itself via /addKey, the application needs to prove it has a copy of one of the existing certificates (via a PoP token).  When removing a key (i.e., certificate) from itself via /removeKey, the application needs to prove it has a copy of one of the other existing certificates (in order to not cutoff its own access accidentally).

Parameters | Description
---------: | :-----------
ApplicationObjectId | The app's Azure AD ObjectId (NOT the ApplicationId/ClientId).
Certificate<br />(Option 1) | Use `$Certificate`, where `$Certificate = Get-ChildItem Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317`
CertificateStorePath<br/>(Option 2) | E.g. 'Cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
JWTExpMinutes | In case of a poorly-synced clock, use this to adjust the expiry of the JWT that is the client assertion sent in the request.  Max. value is 10.

**Example 1**

```powershell
$popTKParams = @{

    ApplicationObjectId = '309f2789-a5ea-4b3a-92ea-687d54249613'
    Certificate         = Get-ChildItem -Path 'Cert:\CurrentUser\My\E25ADCA7817A40C986EB46D5827D804CD805E9B9'
}
$popTK = New-MSGraphPoPToken @popTKParams
```

**Example 2**

```powershell
$popTKParams = @{

    ApplicationObjectId  = '309f2789-a5ea-4b3a-92ea-687d54249613'
    CertificateStorePath = 'Cert:\CurrentUser\My\E25ADCA7817A40C986EB46D5827D804CD805E9B9'
}
$popTK = New-MSGraphPoPToken @popTKParams
```

### Add-MSGraphApplicationKeyCredential

Did you know Azure AD applications can roll their own keys!  No special permissions required, they by default have permission to add new keyCredentials (i.e. public keys) as well as remove previous ones, for themselves.  This means we don't have to worry about long-lasting certificates, or laborious processes to ensure regular certificate/key turnover.  A great security feature of Azure AD applications.

Parameters | Description
---------: | :-----------
ApplicationObjectId | The app's Azure AD ObjectId (NOT the ApplicationId/ClientId).
AccessToken | Use `$TokenObject` where `$TokenObject = New-MSGraphAccessToken ...`.
PoPToken | User `$popTK` where `$popTK = New-MSGraphPoPToken ...`.
Certificate<br />(Option 1) | Use `$Certificate`, where `$Certificate = Get-ChildItem Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317`
CertificateStorePath<br/>(Option 2) | E.g. 'Cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'

**Example 1**

```powershell
$TokenObjectParams = @{

    ApplicationId = '309f2789-a5ea-4b3a-92ea-687d54249613'
    TenantId      = '1098c395-c09a-4c84-8e5e-2f454f318667'
    Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
}
$TokenObject = New-MSGraphAccessToken @TokenObjectParams

$popTKParams = @{

    ApplicationObjectId = '2d6ac42c-f4d9-4ba2-9897-220fc87bec8f'
    Certificate         = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
}
$popTK = New-MSGraphPoPToken @popTKParams

$dateString = [datetime]::Now.ToString('yyyyMMdd')
$newCert = New-SelfSignedMSGraphApplicationCertificate -Subject "$($dateString).jb365.ca"

$addKeyParams = @{

    AccessToken         = $TokenObject
    PoPToken            = $popTK
    ApplicationObjectId = $popTKParams['ApplicationObjectId']
    Certificate         = $newCert
}
$addKeyResponse = Add-MSGraphApplicationKeyCredential @addKeyParams
```

### Remove-MSGraphApplicationKeyCredential

This function is used to remove a keyCredential (i.e. certificate public key) from the Azure AD application.  The access token and PoP token should obtained/created using a different certificate than then one to be removed, otherwise MS Graph will hand back an error that doesn't make any sense, although it is doing so to protect the application from rendering itself unable to get authorization again.

Parameters | Description
---------: | :-----------
ApplicationObjectId | The app's Azure AD ObjectId (NOT the ApplicationId/ClientId).
AccessToken | Use `$TokenObject` where `$TokenObject = New-MSGraphAccessToken ...`.
PoPToken | User `$popTK` where `$popTK = New-MSGraphPoPToken ...`.
CertificateThumbprint<br/>(Option 1) | The thumbprint of the certificate/keyCredential to be removed.  The function will send a request to MS Graph to get the list of current keyCredentials, to find the keyId value for the one with the matching thumbprint.
KeyId<br/>(Option 2) | The keyId (Guid) of the keyCredential to be removed.  This value can be found via MS Graph (i.e. request to applications/{id}) or in the Azure AD Portal.

**Example 1**

```powershell
$TokenObjectParams = @{

    ApplicationId = '309f2789-a5ea-4b3a-92ea-687d54249613'
    TenantId      = '1098c395-c09a-4c84-8e5e-2f454f318667'
    Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
}
$TokenObject = New-MSGraphAccessToken @TokenObjectParams

$popTKParams = @{

    ApplicationObjectId = '2d6ac42c-f4d9-4ba2-9897-220fc87bec8f'
    Certificate         = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
}
$popTK = New-MSGraphPoPToken @popTKParams

$removeKeyParams = @{

    AccessToken           = $TokenObject
    PoPToken              = $popTK
    ApplicationObjectId   = $popTKParams['ApplicationObjectId']
    CertificateThumbprint = 'D7BC8D4E6A10053C29B62D2E2978CDDDEC5526AE'
}
Remove-MSGraphApplicationKeyCredential @removeKeyParams
```

**Example 2**

```powershell
$TokenObjectParams = @{

    ApplicationId = '309f2789-a5ea-4b3a-92ea-687d54249613'
    TenantId      = '1098c395-c09a-4c84-8e5e-2f454f318667'
    Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
}
$TokenObject = New-MSGraphAccessToken @TokenObjectParams

$popTKParams = @{

    ApplicationObjectId = '2d6ac42c-f4d9-4ba2-9897-220fc87bec8f'
    Certificate         = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
}
$popTK = New-MSGraphPoPToken @popTKParams

$removeKeyParams = @{

    AccessToken         = $TokenObject
    PoPToken            = $popTK
    ApplicationObjectId = $popTKParams['ApplicationObjectId']
    KeyId               = '046c2a7c-8b6f-4221-a5b5-4a4e8da12ce3'
}
Remove-MSGraphApplicationKeyCredential @removeKeyParams
```

## Additional Functions

### ConvertFrom-JWTAccessToken

⚠ **Warning:** Microsoft officially recommends that we NEVER look inside the access_token, and claims they will eventually encrypt all issued access tokens.  When that happens, this function will be obsolete.  And until then, it's still a sort of 'faux pas'.

This utility function is included to give an easy way to inspect access tokens and see what headers and claims they contain.  Note that MSA (Microsoft Account) access tokens are different and won't work with this function.  Nor will refresh tokens.  Just access tokens from Azure AD for use with organizations.

Parameters | Description
---------: | :-----------
JWT | Use `$TokenObject.access_token` where `$TokenObject = New-MSGraphAccessToken ...`.

**Example 1**

```powershell
$TokenObjectParams = @{

    ApplicationId = '309f2789-a5ea-4b3a-92ea-687d54249613'
    TenantId      = '1098c395-c09a-4c84-8e5e-2f454f318667'
    Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
}
$TokenObject = New-MSGraphAccessToken @TokenObjectParams

ConvertFrom-JWTAccessToken -JWT $TokenObject.access_token | Select-Object -ExpandProperty Headers
ConvertFrom-JWTAccessToken -JWT $TokenObject.access_token | Select-Object -ExpandProperty Payload
```

### New-RefreshTokenCredential

This function makes it easy to store a token object from `New-MSGraphAccessToken` (and ApplicationId Guid) as a PSCredential object which can be exported securely with Export-Clixml for easy reuse later on.  `New-MSGraphAccessToken` also has a parameter - `-RefreshTokenCredential` - which takes a PSCredential object that has been prepared by this function.

Parameters | Description
---------: | :-----------
ApplicationId | The app's ApplicationId (a.k.a. ClientId)<br />*(This parameter will go away in a future version, as 'appid' is available from the Payload of the access_token itself, so no need to require it via a parameter.)*
TokenObject | Use `$TokenObject` where `$TokenObject = New-MSGraphAccessToken -Scopes offline_access, ...` (must have the refresh_token (hence `-Scopes offline_access`)).

**Example 1**

```powershell
$TokenObject = Get-MSGraphAccessToken -ApplicationId '0c26a905-7c94-4296-aeb4-b8925cb7e036' -Scopes mail.send, offline_access

$RTCredential = New-RefreshTokenCredential -ApplicationId '0c26a905-7c94-4296-aeb4-b8925cb7e036' -TokenObject $TokenObject
$RTCredential | Export-Clixml .\RefreshToken.xml

# ...then later on :

$RTCredential = Import-Clixml .\RefreshToken.xml
$NewTokenObject = New-MSGraphAccessToken -RefreshTokenCredential $RTCredential
```

### Get-AccessTokenExpiration

This functions makes use of `New-MSGraphAccessToken`'s shimmed-in `issued_at` property to determine when your access_token is due to expire.

Parameters | Description
---------: | :-----------
TokenObject | Use `$TokenObject` where `$TokenObject = New-MSGraphAccessToken ...`.

**Example 1**

```powershell
$TokenObjectParams = @{

    ApplicationId = '309f2789-a5ea-4b3a-92ea-687d54249613'
    TenantId      = '1098c395-c09a-4c84-8e5e-2f454f318667'
    Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
}
$TokenObject = New-MSGraphAccessToken @TokenObjectParams

Get-AccessTokenExpiration -TokenObject $TokenObject

# Output:

IssuedAt_LocalTime       : 9/13/2021 8:02:26 PM
ExpirationTime_LocalTime : 9/13/2021 9:07:26 PM
TimeUntilExpiration      : 00:59:50.9092678
IsExpired                : False
```

## Samples

Check out the [sample script](https://github.com/JeremyTBradshaw/MSGraphPSEssentials/tree/main/Samples).  It demonstrates a few concepts for how to use this module with scripts/functions.  Some example topics covered are add/remove keyCredentials (certificates), how to detect and respond to MS Graph throttling responses, and more.

## References

A lot of time and effort went into researching, development, and then testing to get these functions working as well as they do.  A few key blog posts, and otherwise lots and lots of Microsoft Docs articles, were my source for figuring this stuff out.  I'm listing links here that I have found particularly useful in the process.  I will add others if they come to mind as well.

**OAuth2 / JWT IETF's**

https://tools.ietf.org/html/rfc6749<br/>
https://tools.ietf.org/html/rfc7519<br/>

**Client Credentials flow (certificate credentials) / Device Code flow**

https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow<br/>
https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-certificate-credentials<br/>
https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-net-client-assertions<br/>
https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/Client-Assertions<br/>
https://docs.microsoft.com/en-us/azure/architecture/multitenant-identity/client-assertion<br/>
https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code<br/>
https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#refresh-the-access-token<br/>
https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth<br/>

**Application KeyCredential Management**

https://docs.microsoft.com/en-us/graph/api/application-addkey?view=graph-rest-1.0&tabs=http<br/>
https://docs.microsoft.com/en-us/graph/api/application-removekey?view=graph-rest-1.0&tabs=http<br/>
https://docs.microsoft.com/en-us/graph/application-rollkey-prooftoken<br/>
https://docs.microsoft.com/en-us/powershell/module/azuread/new-azureadapplicationkeycredential?view=azureadps-2.0#example-2--use-a-certificate-to-add-an-application-key-credential<br/>

**Microsoft Graph Reference / Examples**

https://docs.microsoft.com/en-us/graph/use-the-api<br/>
https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0<br/>
https://docs.microsoft.com/en-us/graph/paging?context=graph/api/1.0<br/>
https://docs.microsoft.com/en-us/graph/query-parameters?context=graph/api/1.0<br/>

**Microsoft Graph Batching / Service Limits / Throttling**

https://docs.microsoft.com/en-us/graph/json-batching?context=graph/api/1.0<br/>
https://docs.microsoft.com/en-us/graph/throttling?context=graph/api/1.0<br/>
https://docs.microsoft.com/en-us/graph/throttling?context=graph/api/1.0#throttling-and-batching<br/>
https://docs.microsoft.com/en-us/graph/throttling?context=graph/api/1.0#service-specific-limits<br/>
