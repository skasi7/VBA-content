---
title: Folder.Folders Property (Outlook)
keywords: vbaol11.chm1989
f1_keywords:
- vbaol11.chm1989
ms.prod: OUTLOOK
api_name:
- Outlook.Folder.Folders
ms.assetid: 41464c32-023e-9079-4f24-51586305325c
---


# Folder.Folders Property (Outlook)

Returns the  **[Folders](http://msdn.microsoft.com/library/folders-object-outlook%28Office.15%29.aspx)** collection that represents all the folders contained in the specified **[Folder](folder-object-outlook.md)**. Read-only.


## Syntax

 _expression_. **Folders**

 _expression_ A variable that represents a **Folder** object.


## Remarks

The  **[NameSpace](namespace-object-outlook.md)** object is the root of all the folders for the given name space.


## Example

This Visual Basic for Applications (VBA) example uses the  **[Folders.Add](http://msdn.microsoft.com/library/folders-add-method-outlook%28Office.15%29.aspx)** method to add the new folder named "My Personal Contacts" to the default **Contacts** folder.


```
Sub CreatePersonalContacts() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myNewFolder As Outlook.Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderContacts) 
 
 Set myNewFolder = myFolder.Folders.Add("My Personal Contacts") 
 
End Sub
```

This VBA example uses the  **Folders.Add** method to add two new folders in the **Tasks** folder. The first folder, "My Notes Folder", will contain note items. The second folder, "My Contacts Folder", will contain contact items. If the folders already exist, a message box will inform the user.




```
Sub CreateFolders() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myNotesFolder As Outlook.Folder 
 
 Dim myContactFolder As Outlook.Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderTasks) 
 
 On Error GoTo ErrorHandler 
 
 Set myNotesFolder = _ 
 
 myFolder.Folders.Add("My Notes Folder", olFolderNotes) 
 
 Set myContactFolder = _ 
 
 myFolder.Folders.Add("My Contacts Folder", olFolderContacts) 
 
 Exit Sub 
 
ErrorHandler: 
 
 MsgBox "Error creating the folder. The folder may already exist." 
 
 Resume Next 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)
#### Other resources


[Folder Object Members](http://msdn.microsoft.com/library/folder-members-outlook%28Office.15%29.aspx)
