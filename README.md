# OWA-MailboxSearch
A PowerShell tool that leverages the EWS API to search Email Subject and Body content against user provided terms. For any given match, both Email Subject and Body are returned for user viewing.

Note: Requires the Microsoft.Exchange.WebServices.dll. See script for link to the 2.2 EWS API
Note2: I've uploaded the v2.2 API's DLL. Feel free to use or not use as preferred.

v1.0 - Latest Version


## Current Functions:
    Invoke-EmailSubjectSearch  -   Searches Email Subjects for matches against provided terms and returns email content.
    Invoke-EmailBodySearch     -   Searches Email bodys for matches against provided terms and returns email content.
    Get-FolderItems            -   Returns all emails within the user provided Mailbox folder.
    
    



