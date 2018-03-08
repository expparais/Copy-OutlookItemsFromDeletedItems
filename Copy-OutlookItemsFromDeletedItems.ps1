<#
.SYNOPSIS
Copy-OutlookItemsFromDeletedItems copies any deleted items to an Inbox subfolder.

.DESCRIPTION
Copy-OutlookItemsFromDeletedItem Copies any deleted items to an Inbox subfolder using the -ComObject Outlook.application in MS Outlook 2010.

.NOTES
Beside copying items from Deleted Items, you can modify $deletedItemsFolder = $outlookAccount.Folders | ? {$_.Name -match 'Deleted Items' } to $anyOtherItemsFolder =  $inboxFolder.Folders | ? {$_.Name -match 'NAME-OF-FOLDER' }
to copy or move ({$_.moveTo($powershellScriptFolder)}).

$outlookAccount | Get-Member --->>> to get the available members at your disposal.

.REFERENCE
If you need more information about COM Object, visit:
https://docs.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/creating-.net-and-com-objects--new-object-?view=powershell-6

If you need more information about MAPI, visit:
https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/namespace-object-outlook

Author:  Omar Rosa
 
#>


$OutlookApp = New-Object -ComObject outlook.application # We need to create a new object that will contain the COM Object Outlook.application
$outlookNameSpace = $OutlookApp.GetNameSpace("MAPI") # Load the MAPI NameSpace
$outlookAccount = $outlookNameSpace.folders | ? {$_.Name -eq 'username@domain.com' }
$inboxFolder = $outlookAccount.Folders | ? {$_.Name -match 'Inbox' }

$deletedItemsFolder = $outlookAccount.Folders | ? {$_.Name -match 'Deleted Items' } # From folder
$powershellScriptFolder = $inboxFolder.Folders | ? {$_.Name -match 'PowershellScript'} # To folder

$deletedItemsFolder | ForEach-Object {$_.copyTo($powershellScriptFolder)}



