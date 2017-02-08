---
title: Store Object (Outlook)
keywords: vbaol11.chm3155
f1_keywords:
- vbaol11.chm3155
ms.prod: OUTLOOK
api_name:
- Outlook.Store
ms.assetid: 1eb22fe9-8849-7476-5388-2515b48591b9
---


# Store Object (Outlook)

Represents a file on the local computer or a network drive that stores e-mail messages and other items for an account in the current profile.


## Remarks

A profile defines one or more e-mail accounts, and each e-mail account is associated with a server of a specific type. For an Exchange server, a store can be on the server, in an Exchange Public folder, or in a local Personal Folders File (.pst) or Offline Folder File (.ost). For a POP3, IMAP, or HTTP e-mail server, a store is a .pst file.

You can use the  **[Stores](stores-object-outlook.md)** and **Store** objects to enumerate all folders and search folders on all stores in the current session. Since getting the root folder or search folders in a store requires the store to be open and opening a store imposes an overhead on performance, you can check the **[Store.IsOpen](http://msdn.microsoft.com/library/store-isopen-property-outlook%28Office.15%29.aspx)** property before you decide to pursue the operation.

If you use an Exchange server, you can access other explicit built-in  **Store** properties for store characteristics such as **[ExchangeStoreType](http://msdn.microsoft.com/library/store-exchangestoretype-property-outlook%28Office.15%29.aspx)**, **[IsCachedExchange](http://msdn.microsoft.com/library/store-iscachedexchange-property-outlook%28Office.15%29.aspx)**, and **[IsDataFileStore](http://msdn.microsoft.com/library/store-isdatafilestore-property-outlook%28Office.15%29.aspx)**. Use the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object returned by **[Store.PropertyAccessor](http://msdn.microsoft.com/library/store-propertyaccessor-property-outlook%28Office.15%29.aspx)** to access other store properties that are not exposed in the Outlook object model.

For more information on storing Outlook items in folders and stores, see [Storing Outlook Items](http://msdn.microsoft.com/library/storing-outlook-items%28Office.15%29.aspx).


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) enumerates all folders on all stores for a session:


```
Sub EnumerateFoldersInStores() 
 
 Dim colStores As Outlook.Stores 
 
 Dim oStore As Outlook.Store 
 
 Dim oRoot As Outlook.Folder 
 
 
 
 On Error Resume Next 
 
 Set colStores = Application.Session.Stores 
 
 For Each oStore In colStores 
 
 Set oRoot = oStore.GetRootFolder 
 
 Debug.Print (oRoot.FolderPath) 
 
 EnumerateFolders oRoot 
 
 Next 
 
End Sub 
 
 
 
Private Sub EnumerateFolders(ByVal oFolder As Outlook.Folder) 
 
 Dim folders As Outlook.folders 
 
 Dim Folder As Outlook.Folder 
 
 Dim foldercount As Integer 
 
 
 
 On Error Resume Next 
 
 Set folders = oFolder.folders 
 
 foldercount = folders.Count 
 
 'Check if there are any folders below oFolder 
 
 If foldercount Then 
 
 For Each Folder In folders 
 
 Debug.Print (Folder.FolderPath) 
 
 EnumerateFolders Folder 
 
 Next 
 
 End If 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[GetDefaultFolder](http://msdn.microsoft.com/library/store-getdefaultfolder-method-outlook%28Office.15%29.aspx)|
|[GetRootFolder](http://msdn.microsoft.com/library/store-getrootfolder-method-outlook%28Office.15%29.aspx)|
|[GetRules](http://msdn.microsoft.com/library/store-getrules-method-outlook%28Office.15%29.aspx)|
|[GetSearchFolders](http://msdn.microsoft.com/library/store-getsearchfolders-method-outlook%28Office.15%29.aspx)|
|[GetSpecialFolder](http://msdn.microsoft.com/library/store-getspecialfolder-method-outlook%28Office.15%29.aspx)|
|[RefreshQuotaDisplay](http://msdn.microsoft.com/library/store-refreshquotadisplay-method-outlook%28Office.15%29.aspx)|
|[CreateUnifiedGroup](http://msdn.microsoft.com/library/store-createunifiedgroup-method-outlook%28Office.15%29.aspx)|
|[DeleteUnifiedGroup](http://msdn.microsoft.com/library/store-deleteunifiedgroup-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/store-application-property-outlook%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/store-categories-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/store-class-property-outlook%28Office.15%29.aspx)|
|[DisplayName](http://msdn.microsoft.com/library/store-displayname-property-outlook%28Office.15%29.aspx)|
|[ExchangeStoreType](http://msdn.microsoft.com/library/store-exchangestoretype-property-outlook%28Office.15%29.aspx)|
|[FilePath](http://msdn.microsoft.com/library/store-filepath-property-outlook%28Office.15%29.aspx)|
|[IsCachedExchange](http://msdn.microsoft.com/library/store-iscachedexchange-property-outlook%28Office.15%29.aspx)|
|[IsConversationEnabled](http://msdn.microsoft.com/library/store-isconversationenabled-property-outlook%28Office.15%29.aspx)|
|[IsDataFileStore](http://msdn.microsoft.com/library/store-isdatafilestore-property-outlook%28Office.15%29.aspx)|
|[IsInstantSearchEnabled](http://msdn.microsoft.com/library/store-isinstantsearchenabled-property-outlook%28Office.15%29.aspx)|
|[IsOpen](http://msdn.microsoft.com/library/store-isopen-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/store-parent-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/store-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/store-session-property-outlook%28Office.15%29.aspx)|
|[StoreID](http://msdn.microsoft.com/library/store-storeid-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[Store Object Members](http://msdn.microsoft.com/library/store-members-outlook%28Office.15%29.aspx)
