# OWA-MailboxSearch
A PowerShell tool that leverages the EWS API to search and pull content from an MS Exchange user's Mailbox. In all cmdlets the email subject and body are returned for viewing.

Note: The Microsoft.Exchange.WebServices.dll is now embedded within the script to make it more portable.


## Current Functions:
    Invoke-SearchEmailSubject  -   Searches Email Subjects for matches against provided terms and returns email content.
    Invoke-SearchEmailBody     -   Searches Email Bodys for matches against provided terms and returns email content.
    Get-FolderContents         -   Returns all emails within the user provided folder id, acquired from Get-Folders.
    Get-Folders                -   Returns hash table containing all folder names and associated ids.
    

    



