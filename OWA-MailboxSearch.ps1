<#
.SYNOPSIS

Extract MailBox contents based upon user provided searchterms.

Author: Andrew '@ch33kyf3ll0w' Bonstrom
License: BSD 3-Clause
Required Dependencies: Microsoft.Exchange.WebServices.dll
Optional Dependencies: None

.VERSION

1.0

.LINK

EWS API 2.2: https://www.microsoft.com/en-us/download/details.aspx?id=42951

#>

<#
Helper Functions Begin
#>

function Get-ExchangeVersion ($ExchangeVersion){


    #Set exchange version based upon user provided input
    switch ($ExchangeVersion){
        Exchange2007SP1{
		    $exchVUserProv = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
		    break
        }
        Exchange2010{
		    $exchVUserProv = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010
		    break
        }
        Exchange2010SP1{
		    $exchVUserProv = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
		    break
        }
        Exchange2010SP2{
		    $exchVUserProv = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
		    break
        }
        Exchange2013{
		    $exchVUserProv = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
		    break
        }
        Exchange2013SP1{
		    $exchVUserProv = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
		    break
        }

    }

    return $exchVUserProv

}

function Get-ExchServiceObject ($UserEmail, $exchVUserProv, $UserName, $UserPassword, $UserDomain){
    
    #Create a new object containing an EWS instance
    #Also feeds new object the user provided credentials and the autodiscover url
    $exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchVUserProv) 
    $exchService.Credentials = New-Object System.Net.NetworkCredential -ArgumentList $UserName, $UserPassword, $UserDomain 
    $exchService.AutodiscoverUrl($UserEmail)


    return $exchService
}

function Get-MailboxFolderIDs ($exchService){

    #Section to find all folders within a user's mailbox
    #Credits to http://gsexdev.blogspot.com/2012/01/ews-managed-api-and-powershell-how-to_23.html
    #Define Extended properties  
    $PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    $folderidcnt = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
  
    #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
    $fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000) 
 
    #Deep Transval will ensure all folders in the search path are returned  
    $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
 
    $psPropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $PR_MESSAGE_SIZE_EXTENDED = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long) 
    $PR_DELETED_MESSAGE_SIZE_EXTENDED = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26267,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long)
    $PR_Folder_Path = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)

    #Add Properties to the  Property Set  
    $psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED) 
    $psPropertySet.Add($PR_Folder_Path);  
    $fvFolderView.PropertySet = $psPropertySet

    #The Search filter will exclude any Search Folders  
    $sfSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  
    $fiResult = $null 

    #Initialize array to store folder names
    $folderList = @()

    #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
    do {  
        $fiResult = $exchService.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)  
        foreach($ffFolder in $fiResult.Folders){  
            #Append folder ID to array
            $folderList += $ffFolder.Id 
        } 

    }while($fiResult.MoreAvailable -eq $true)

    return $FolderList

}
<#
Helper Functions End
#>
<#
CMDLETS Begin
#>
function Invoke-EmailSubjectSearch{

<#
.SYNOPSIS

Extract MailBox contents based upon user provided searchterms.

.DESCRIPTION

Invoke-EmailSubjectSearch is designed to search the subjects of any and all stored emails for references the user provides and to then present the email's contents.

.LINK

EWS API 2.2: https://www.microsoft.com/en-us/download/details.aspx?id=42951

.PARAMETER UserEmail

Email of the user whos mailbox you wish to search.

.PARAMETER Credential

Username in "DOMAIN\USERNAME" formation to present a credential entry pop up.

.PARAMETER ExchangeVersion

The Exchange version the MailBox you are targeting uses.

Accepted Versions:

Exchange2007_SP1
Exchange2010
Exchange2010_SP1
Exchange2010_SP2
Exchange2013
Exchange2013_SP1

.PARAMETER SearchTerms

Terms you wish to search for seperated by a |.

.PARAMETER DLLPath

Path to the Microsoft.Exchange.WebServices.dll.

.EXAMPLE

Invoke-EmailSubjectSearch -UserEmail administrator@ch33z.local -ExchangeVersion Exchange2013 -Credential "CH33KZ\administrator" -SearchTerms "password|Password|dawg" -DLLPath "C:\Users\Administrator\Documents\Microsoft.Exchange.WebServices.dll"

#>
    [CmdletBinding()]
    Param(
	
    [Parameter(Mandatory = $True)]
    [string]
    $UserEmail,

    [Parameter(Mandatory = $False)]
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $True)]
    [string]
    $ExchangeVersion,

    [Parameter(Mandatory = $True)]
    [string]
    $SearchTerms,

    [Parameter(Mandatory = $True)]
    [string]
    $DLLPath

    )

try{
    #Add the assembly type for the API access
    Add-Type -Path $DLLPath

    #Assigns appropriate version of Exchangeversion based on user provided value
    $exchVUserProv = Get-ExchangeVersion $ExchangeVersion

    #Create a new object containing an EWS instance
    #Also feeds new object the user provided credentials and the autodiscover url
    $exchService = Get-ExchServiceObject $UserEmail $exchVUserProv $Credential.UserName $Credential.Password $Credential.Domain

    #Section to search through items in identfied folders
    #Credits to https://social.technet.microsoft.com/Forums/scriptcenter/en-US/335a888b-bf85-4a36-a555-71cc84608960/download-email-content-text-from-exchange-ews-with-powershell?forum=ITCG
    #Define the Item view
    $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)

    #Define PropertySet so we receive the contents of the email without HTML
    $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text


    #Call helper function to provide us with all folder ids
    $FolderList = Get-MailboxFolderIDs $exchService

}
catch{

    Write-Output "[-] Please review the script's Get-Help output to ensure parameters are correct,i.e. Get-Help Invoke-EmailSubjectSearch -Detailed. Error: $_"

}

#Search items within every identified folder
ForEach($folder in $FolderList ){
    #Search the items for every previously identfied folder
    $Mailitems = $exchService.FindItems($folder,$itemView)

    ForEach($Mailitem in $Mailitems){
        
        if($MailItem.Subject -match $SearchTerms){             
            #Load the email content with out predefined property set
            $MailItem.Load($PropertySet)
            $MailBody = $MailItems.Body.Text

            #Present user with item subject and content
            Write-Output "Sender: " $MailItem.Sender.Address
            Write-Output "___________________________________"
            Write-Output "Email Subject:" $Mailitem.Subject
            Write-Output "___________________________________"
            Write-Output "Email Content:" $MailBody "`n"
            }   
        } 
      } 
      
}

function Invoke-EmailBodySearch{
<#

.SYNOPSIS

Extract MailBox contents based upon user provided searchterms.

.DESCRIPTION

Invoke-OWAEmailBodySearch is designed to search the body of any and all stored emails for references the user provides and to then present the email's contents in HTML format.

.LINK

EWS API 2.2: https://www.microsoft.com/en-us/download/details.aspx?id=42951

.PARAMETER UserEmail

Email of the user whos mailbox you wish to search.

.PARAMETER Credential

Username in "DOMAIN\USERNAME" formation to present a credential entry pop up.

.PARAMETER ExchangeVersion

The Exchange version the MailBox you are targeting uses.

Accetped Versions:

Exchange2007_SP1
Exchange2010
Exchange2010_SP1
Exchange2010_SP2
Exchange2013
Exchange2013_SP1

.PARAMETER SearchTerms

Terms you wish to search for seperated by a |.

.PARAMETER DLLPath

Path to the Microsoft.Exchange.WebServices.dll.

.EXAMPLE

Invoke-EmailBodySearch -UserEmail administrator@ch33z.local -ExchangeVersion Exchange2013 -Credential "CH33KZ\administrator" -SearchTerms "password|Password|dawg" -DLLPath "C:\Users\Administrator\Documents\Microsoft.Exchange.WebServices.dll"

#>
    [CmdletBinding()]
    Param(
	
    [Parameter(Mandatory = $True)]
    [string[]]
    $UserEmail,

    [Parameter(Mandatory = $False)]
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $True)]
    [string]
    $ExchangeVersion,

    [Parameter(Mandatory = $True)]
    [string]
    $SearchTerms,

    [Parameter(Mandatory = $True)]
    [string]
    $DLLPath

    )
try{
    #Load the EWS API DLL
    Add-Type -Path $DLLPath

    #Assigns appropriate version of Exchangeversion based on user provided value
    $exchVUserProv = Get-ExchangeVersion $ExchangeVersion

    #Create a new object containing an EWS instance
    #Also feeds new object the user provided credentials and the autodiscover url
    $exchService = Get-ExchServiceObject $UserEmail $exchVUserProv $Credential.UserName $Credential.Password $Credential.Domain

    #Section to search through items in identfied folders
    #Credits to https://social.technet.microsoft.com/Forums/scriptcenter/en-US/335a888b-bf85-4a36-a555-71cc84608960/download-email-content-text-from-exchange-ews-with-powershell?forum=ITCG
    #Create view for items
    $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)

    #Define PropertySet so we receive the contents of the email without HTML
    $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

    #Call helper function to provide us with all folder ids
    $FolderList = Get-MailboxFolderIDs $exchService

}
catch{

    Write-Output "[-] Please review the script's Get-Help output to ensure parameters are correct, i.e. Get-Help Invoke-EmailBodySearch -Detailed. Error: $_"

}

#Search items within every identified folder
ForEach($folder in $FolderList ){

    $Mailitems = $exchService.FindItems($folder,$itemView)

    ForEach($Mailitem in $Mailitems){
        $MailItem.Load($PropertySet)
        if($MailItem.Body.Text -match $SearchTerms){ 
            #Strip out newline chars      
            $MailBody = $MailItems.Body.Text

            #Present user with item subject and content
            Write-Output "Sender: " $MailItem.Sender.Address
            Write-Output "___________________________________"
            Write-Output "Email Subject:" $Mailitem.Subject
            Write-Output "___________________________________"
            Write-Output "Email Content:" $MailBody "`n"

            }
        }  
      } 
      
}
