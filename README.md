# MSGraphPSEssentials
A collection of functions enabling easier consumption of Microsoft Graph using just PowerShell (Desktop/Core).

[Available on PowerShell Gallery](https://www.powershellgallery.com/packages/MSGraphPSEssentials)

This module is the successor to [MSGraphAppOnlyEssentials](https://github.com/JeremyTBradshaw/MSGraphAppOnlyEssentials), which was geared specifically towards App-Only use cases.  With this module, I've broadened the scope to support additional authentication flows which use delegated permissions rather than application permissions.  This allows me to accommodate scripters who plan to write scripts/modules which can do things for users in organizational tenants (i.e. Work/School accounts) as well as personal Microsoft accounts.  Aside from the broader target audience, delegated permissions can be requested on the fly, and they're commonly limited in scope to the user who is delegating the access, making this approach a lot more accessible in terms of effort with the pre-setup of the App Registration, and palatability from a security standpoint.

The module provides six essential functions for working with Microsoft Graph using PowerShell 5.1 and newer, in App-Only fashion using certificate credentials, _**or**_ in Delegated fashion, using Device Code flow (and/or Refresh Tokens).  There are a few other short and sweet utility functions included as well.  Obtain access tokens, perform nearly any MS Graph request you can think of, and roll your own keys using the addKey and removeKey application resource type methods.  The only thing left to do manually is to setup an App Registration in Azure AD and upload the initial certificate (for App-Only scenarios).  Simply use the [MS Graph API (v1.0 / beta) reference material](https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0) to figure out your requests and query syntax, then start sending them with `New-MSGraphRequest`.  Options for handling paging are provided, allowing 'nextLink' responses to be handled however you prefer (automatic, inquire, warn, etc.).  Best of all - no DLL's; no big files at all.  In fact, only about 30-40KB.

💡 **[Samples are now live in the wiki](https://github.com/JeremyTBradshaw/MSGraphPSEssentials/wiki)**

---

## (Essential) Functions

### New-MSGraphAccessToken

This function is used to get an access token.  Access tokens last for 1 hour, so keep this in mind in long-running scripts.

ℹ **Note:** For a less-powerful (security-wise) alternative to app-only access tokens, you may find that you can accomplish enough unattended'ness by interatively requesting the initial access token using Device Code flow, and then programmatically refreshing the access token before the refresh token expires.  As long as the refresh token is never revoked, this could go on forever, unattended.

Parameters | Parameter Set | Description
---------: | :-----------: | :-----------
ApplicationId | All parameter sets | The app's ApplicationId (a.k.a. ClientId)
TenantId | All parameter sets | The directory/tenant ID (Guid)
Certificate<br />(Option 1) | ClientCredentials_Certificate | Use `$Certificate`, where `$Certificate = Get-ChildItem Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317`
CertificateStorePath<br/>(Option 2) | ClientCredentials_CertificateStorePath | E.g. 'Cert:\LocalMachine\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
JWTExpMinutes | ClientCredentials_* | In case of a poorly-synced clock, use this to adjust the expiry of the JWT that is the client assertion sent in the request.  Max. value is 10.
Endpoint | DeviceCode_\*, RefreshToken_* | **Common**, Consumers, Organizations.  Common (default) should suffice for most scenarios.  Scopes will dictate when either of the other options must be used instead.
Scopes | DeviceCode_\*, RefreshToken_* | One or more delegated permissions (e.g. Ews.AccessAsUser.All, mail.Send, offline_access).  Refer to MS Graph reference pages for required permissions.  Those permissions are exaclty what to specify here.
RefreshToken | RefreshToken_* | Use `$TokenObject` where `$TokenObject = New-MSGraphAccessToken -Scopes ..., offline_access`.  The 'offline_access' scope is what tells Azure AD to also hand back a refresh token when issuing an access token.

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

(New) Get an access token using the Device Code flow:

```powershell
$TokenObject = Get-MSGraphAccessToken -ApplicationId '0c26a905-7c94-4296-aeb4-b8925cb7e036' -Scopes mail.send, offline_access

# Follow the instructions, which will look similar to this:
Authorization started (expires at 2020-12-19 03:01:44 PM)
To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code FB228GRH7 to authenticate.
[D] Done  [?] Help (default is "D"):
```

**Example 4**

(New) Refresh an access token (which was obtained using `New-MSGraphAccessToken` and including 'offline_access' in the `-Scopes` parameter):

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
nextLinkAction | **Warn**, Inquire, Continue, SilentlyContinue.  When multiple pages of results are available (i.e. more than 100 users), use this to decide how and whether to keep getting the next page of results.

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
DnsName | Any FQDN of choice.  E.g. 20201116.jb365.ca
FriendlyName | "jb365 automation 2020-11-16"
CertStoreLocation | Maps directly to the same parameter of New-SelfSignedCertificate.  Default is 'cert:\CurrentUser\My'.  Any valid location where write access is available will work (e.g. 'cert:\LocalMachine\My', when scheduling tasks using local SYSTEM account).
NotAfter | Default is 90 days.  Supply a [datetime] like this `(Get-Date).AddDays(7)`.  The shorter the better, because applications can add new certificates for themselves using the addKey method, so we can easily roll these often, programmatically.
KeySpec | **Signature**, KeyExchange.  Recommendation: don't change this unless there is a reason.

**Example 1**

```powershell
New-SelfSignedMSGraphApplicationCertificate -DnsName "signInsMonitor.jb365.ca" -FriendlyName "jb365 signIns monitor ($($date))"
```

**Example 2**

```powershell
$date = [datetime]::Now.ToString('yyyyMMdd')
$newCertParams = @{
    DnsName           = "$($date).jb365.ca"
    FriendlyName      = "jb365 automation ($($date))"
    CertStoreLocation = 'Cert:\LocalMachine\My'
    NotAfter          = (Get-Date).AddDays(7)
}
New-SelfSignedMSGraphApplicationCertificate @newCertParams
```

### New-MSGraphPoPToken

This function generates a Proof of Possession JWT (JSON Web Token) for use with the addKey/removeKey MS Graph application resource type methods (in this case the `Add-MSGraphApplicationKeyCredential` and `Remove-MSGraphApplicationKeyCredential` functions).

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
$newCert = New-SelfSignedMSGraphApplicationCertificate -DnsName "$($dateString).jb365.ca" -FriendlyName "$($dateString).jb365.ca"

$addKeyParams = @{

    AccessToken         = $TokenObject
    PoPToken            = $popTK
    ApplicationObjectId = $popTKParams['ApplicationObjectId']
    Certificate         = $newCert
}
$addKeyResponse = Add-MSGraphApplicationKeyCredential @addKeyParams
```

### Remove-MSGraphApplicationKeyCredential

This is the sixth and last function in the bunch to cover.  It is used to remove a keyCredential (i.e. certificate public key) from the Azure AD application.  The access token and PoP token should obtained/created using a different certificate than then one to be removed, otherwise MS Graph will hand back an error that doesn't make any sense, although it is doing so to protect the application from rendering itself unable to get authorization again.

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

## Utility Functions

### Test-SigningCertificate

This is a simple utilty function used by some of the essential functions to verify the certificate is using the 'Microsoft Enhanced RSA and AES Cryptographic Provider' and SHA-256 hashing algorigthm.  These two things ensure everything will work as planned with the certificate and these functions.

Parameters | Description
---------: | :-----------
Certificate | Use `$Certificate`, where `$Certificate = Get-ChildItem Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317`

**Example 1**

```powershell
$Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
Test-SigningCertificate $Certificate
# Output will be True or False
```

### ConvertTo-Base64Url

This utility function is used every in this module that deals with JWT's.  Many things inside a JWT (for these purposes anyway) need to be encoded as Base64.  Since Base64 contains some characters that aren't web-friendly, Base64Url to the rescue.

Parameters | Description
---------: | :-----------
String | One or more Base64 encoded strings.

**Example 1**

```powershell
$Certificate   = Get-ChildItem -Path 'Cert:\CurrentUser\My\C3E7F30B9DD50B8B09B9B539BC41F8157642D317'
ConvertTo-Base64Url -String ([Convert]::ToBase64String($Script:Certificate.GetCertHash()))
# Output will be a Base64Url representation of the certificate's hash.
```

### ConvertFrom-Base64Url

Much like ConvertTo-Base64Url, just in the opposite direction.

Parameters | Description
---------: | :-----------
String | One or more Base64Url encoded strings.

**Example 1**

```powershell
$Headers, $Payload = ($TokenObject.access_token -split '\.')[0,1] |
Foreach-Object { ConvertFrom-Base64Url -String $_ }
# Output will be two Base64 strings, which are the first 2 of 3 parts of the JWT that is the access token.
```

### ConvertFrom-JWTAccessToken

This utility function is not actually used by any other functions in this module.  It is simply included to give an easy way to inspect access tokens and see what headers and claims they contain.  Note that MSA (Microsoft Account) access tokens are different and won't work with this function.  Nor will refresh tokens.  Just access tokens from Azure AD for use with organizations.

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

## Next Steps

ℹ**Up next:** Now that the initial version of this module is published to the PowerShell Gallery, I will be focusing more on producing scripts that actually make use of this stuff, to drive some visibility into how/when/why to use this module.  Samples will be included in this repo as well as documented in the wiki.  Example ideas are:

- (Done already) [Get-MailboxLargeItems.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/Get-MailboxLargeItems.ps1) / [New-LargeItemsSearchFolder.ps1](https://github.com/JeremyTBradshaw/PowerShell/blob/main/New-LargeItemsSearchFolder.ps1)
- Script to reorganize large mailboxes into folders by year.  This could be used by Outlook.com users and Exchange Online users alike, and for the latter, either admins for bulk-usage, or individual users.
- Similar scripts, but for OneDrive / OneDrive for Business.
- ...and more, along these lines.

The next planned addition to the module itself will be a function to help with [batching bulk requests to the $batch MS Graph endpoint](https://docs.microsoft.com/en-us/graph/json-batching).  I'm still deciding if this task would be better-suited to a script rather than a function in the module.  Either way, I will be getting it done, have done so already in my own scripts, so it's just a matter of time before it's ready to be shared in its final form.

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
