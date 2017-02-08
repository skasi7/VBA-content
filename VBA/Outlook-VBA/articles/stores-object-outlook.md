---
title: Stores Object (Outlook)
keywords: vbaol11.chm3019
f1_keywords:
- vbaol11.chm3019
ms.prod: OUTLOOK
api_name:
- Outlook.Stores
ms.assetid: 8915a8e4-9c22-21d5-c492-051d393ce5f7
---


# Stores Object (Outlook)

A set of  **[Store](store-object-outlook.md)** objects representing all the stores available in the current profile.


## Remarks

You can use the  **Stores** and **Store** objects to enumerate all folders and search folders on all stores in the current session. For more information on storing Outlook items in folders and stores, see[Storing Outlook Items](http://msdn.microsoft.com/library/storing-outlook-items%28Office.15%29.aspx).


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


## Events



|**Name**|
|:-----|
|[BeforeStoreRemove](http://msdn.microsoft.com/library/stores-beforestoreremove-event-outlook%28Office.15%29.aspx)|
|[StoreAdd](http://msdn.microsoft.com/library/stores-storeadd-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Item](http://msdn.microsoft.com/library/stores-item-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/stores-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/stores-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/stores-count-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/stores-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/stores-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[Stores Object Members](http://msdn.microsoft.com/library/stores-members-outlook%28Office.15%29.aspx)
