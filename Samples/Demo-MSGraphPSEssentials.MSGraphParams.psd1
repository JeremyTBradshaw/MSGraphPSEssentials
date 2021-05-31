@{
    tokenParams          = @{

        ApplicationId        = 'f5575a40-f043-4fe6-83db-1428f5df7212'
        TenantId             = 'de19b221-c000-487e-93e4-545b311cfe68' #<--: <TenantName>.onmicrosoft.com is OK too.
        CertificateStorePath = 'Cert:\CurrentUser\My\B71F324602735236A538472914CC5DAEBE7D4DF8'
    }

    throttlingMaxRetries = 100

    apiPermissions                 = @{
        <#
            List all MS Graph API *application* permissions which are required for the script to work.
            The examples below show all possible permissions for some example tasks, in order of least to most privilege.

            - List AAD users (https://docs.microsoft.com/en-us/graph/api/user-list).  Requires one of:
                -- User.Read.All
                -- User.ReadWrite.All
                -- Directory.Read.All
                -- Directory.ReadWrite.All

            - List AAD users' last sign-in (https://docs.microsoft.com/en-us/graph/api/signin-list):
                -- AuditLog.Read.All and Directory.Read.All

            The script will report all actions descriptions where none of the possible corresponding permissions are
            granted/consented to the registered app in Azure AD.
        #>

        # Action descriptions  = # List of adequate API permissions (per related MS Docs reference page)
        listUsers              = '(User\.Read\.All)|(User\.ReadWrite\.All)|(Directory\.Read\.All)|(Directory\.ReadWrite\.All)'
        listSignIns1           = '(AuditLog\.Read\.All)'
        listSignIns2           = '(Directory\.Read\.All)'
    }

    # How often to replace the certificate both locally (in the certificate store, and in this PSD1 file) and in Azure AD.
    appKeyRolloverDays = 2
    appObjectId = '7767b2dd-a567-4869-a4bb-4686a65a2001' #<--: Needed for /addKey & /removeKey
}
