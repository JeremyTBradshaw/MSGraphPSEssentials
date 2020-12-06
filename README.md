# MSGraphPSEssentials
A collection of functions enabling easier consumption of Microsoft Graph using just PowerShell (Desktop/Core).

This module is the successor of, or maybe sibling to, [MSGraphAppOnlyEssentials](https://github.com/JeremyTBradshaw/MSGraphAppOnlyEssentials), which is geared specifically towards App-Only use cases.  With this module, I intend to broaden the scope and support additional authentication flows which use delegated permissions rather than application permissions.  This allows me to accommodate scripters who plan to write scripts/modules which can do things for users in oranizational tenants (i.e. Work/School accounts) as well as personal Microsoft accounts.

Initially, I'll be porting the `New-MSGraphAccessToken` function over to this module, or at least it's functionality, and I'll be adding in support for the device code flow.  It's actually quite easy to accomplish the device code flow in PowerShell, certainly easier than the client/certicate credentials (signed assertion) flow.  Once the module is up and running, I will (or would like to) then start producing scripts actually make use of this stuff.  Example ideas are:

- Script to reorganize large mailboxes into folders by year.  This could be used by Outlook.com users and Exchange Online users alike, and for the latter, either admins for bulk-usage, or individual users.

- Similar scripts, but for OneDrive / OneDrive for Business.

- ... and more, along these lines.

For the immediate term, I'll be publishing a placeholder module to the PowerShell Gallery to hold down the name - **MSGraphPSEssentials**.
