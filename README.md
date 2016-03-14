# OWA-MailboxSearch
A PowerShell tool that leverages the EWS API to search and pull content from an MS Exchange user's Mailbox. In all cmdlets the email subject and body are returned for viewing.

Note: Requires the Microsoft.Exchange.WebServices.dll. See script for link to the 2.2 EWS API
Note2: I've uploaded the v2.2 API's DLL. Feel free to use or not use as preferred.

v1.1 - Latest Version


## Current Functions:
    Invoke-EmailSubjectSearch  -   Searches Email Subjects for matches against provided terms and returns email content.
    Invoke-EmailBodySearch     -   Searches Email bodys for matches against provided terms and returns email content.
    Get-FolderItems            -   Returns all emails within the user provided Mailbox folder.
    
    



