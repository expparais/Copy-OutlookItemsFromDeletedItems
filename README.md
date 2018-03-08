# Copy-OutlookItemsFromDeletedItems

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
