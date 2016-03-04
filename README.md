# OWA-MailboxSearch
A PowerShell tool that leverages the EWS API to search Email Subject and Body content against user provided terms. For any given match, both Email Subject and Body are returned for user viewing.

Note: Requires the Microsoft.Exchange.WebServices.dll. See script for link to the 2.2 EWS API

v1.0 - Latest Version


## Current Functions:
    Invoke-EmailSubjectSearch  -   Searches Email Subjects for matches against provided terms and returns email content.
    Invoke-EmailBodySearch     -   Searches Email bodys for matches against provided terms and returns email content.
    
## Example Run:  
Invoke-EmailBodySearch -UserEmail administrator@ch33z.local -ExchangeVersion Exchange2013 -Credential "CH33KZ\administrator" -SearchTerms "password|Password|dawg" -DLLPath "C:\Users\Administrator\Documents\Microsoft.Exchange.WebServices.dll"



