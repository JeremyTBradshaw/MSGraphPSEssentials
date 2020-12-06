# MSGraphPSEssentials
A collection of functions enabling easier consumption of Microsoft Graph using just PowerShell (Desktop/Core).

This module is the successor of [MSGraphAppOnlyEssentials](https://github.com/JeremyTBradshaw/MSGraphAppOnlyEssentials), which is geared specifically towards App-Only use cases.  With this module, I intend to broaden the scope and support additional authentication flows which use delegated permissions rather than application permissions.  This allows me to accommodate scripters who plan to write scripts/modules which can do things for users in organizational tenants (i.e. Work/School accounts) as well as personal Microsoft accounts.  Aside from the broader target audience, delegated permissions can be requested on the fly, and they're limited in scope to the user who is delegating the access, making this approach a lot more accessible in terms of effort with the pre-setup of the App Registration, and palatability from a security standpoint.

I'll be porting the `New-MSGraphAccessToken` function over to this module first, then the rest of the original functions after that (which will for the most part be a copy paste).  For the former, I'll be adding in support for the device code flow right out of the gate.  It's actually quite easy to accomplish the device code flow in PowerShell, certainly easier than the client/certicate credentials (signed assertion) flow.  Once the module is up and running, I hope to then start focusing on producing scripts actually make use of this stuff.  Example ideas are:

- (Done already) [Get-MailboxLargeItems.ps1](https://github.com/JeremyTBradshaw/blob/main/Get-MailboxLargeItems.ps1) / [New-LargeItemsSearchFolder.ps1](https://github.com/JeremyTBradshaw/blob/main/New-LargeItemsSearchFolder.ps1)
- Script to reorganize large mailboxes into folders by year.  This could be used by Outlook.com users and Exchange Online users alike, and for the latter, either admins for bulk-usage, or individual users.
- Similar scripts, but for OneDrive / OneDrive for Business.
- ...and more, along these lines.

For the immediate term, I've published a placeholder module to the PowerShell Gallery to hold down the name - **MSGraphPSEssentials**.
