---
title: FileDialog Object (Office)
keywords: vbaof11.chm256000
f1_keywords:
- vbaof11.chm256000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.FileDialog
ms.assetid: 71a030f2-3b02-21e1-c156-0514ff5eddb7
---


# FileDialog Object (Office)

Provides file dialog box functionality similar to the functionality of the standard  **Open** and **Save** dialog boxes found in Microsoft Office applications.


## Remarks

Use the  **FileDialog** property to return a **FileDialog** object. The **FileDialog** property is located in each individual Office application's **Application** object. The property takes a single argument, _DialogType_, that determines the type of **FileDialog** object that the property returns. There are four types of **FileDialog** object:


-  **Open** dialog box - lets users select one or more files that you can then open in the host application using the **Execute** method.
    
-  **SaveAs** dialog box - lets users select a single file that you can then save the current file as using the **Execute** method.
    
-  **File Picker** dialog box - lets users select one or more files. The file paths that the user selects are captured in the **FileDialogSelectedItems** collection.
    
-  **Folder Picker** dialog box - lets users select a path. The path that the user selects is captured in the **FileDialogSelectedItems** collection.
    
Each host application can only create a single instance of the  **FileDialog** object. Therefore, many of the properties of the **FileDialog** object persist even when you create multiple **FileDialog** objects. Therefore, make sure that you set all of the properties appropriately for your purpose before you display the dialog box.


## Example

To display a file dialog box using the  **FileDialog** object, you must use the **Show** method. Once a dialog box is displayed, no code executes until the user dismisses the dialog box. The following example creates and displays a **File Picker** dialog box and then displays each selected file in a message box.


```
Sub Main() 
 
 'Declare a variable as a FileDialog object. 
 Dim fd As FileDialog 
 
 'Create a FileDialog object as a File Picker dialog box. 
 Set fd = Application.FileDialog(msoFileDialogFilePicker) 
 
 'Declare a variable to contain the path 
 'of each selected item. Even though the path is aString, 
 'the variable must be a Variant because For Each...Next 
 'routines only work with Variants and Objects. 
 Dim vrtSelectedItem As Variant 
 
 'Use a With...End With block to reference the FileDialog object. 
 With fd 
 
 'Use the Show method to display the File Picker dialog box and return the user's action. 
 'The user pressed the button. 
 If .Show = -1 Then 
 
 'Step through each string in the FileDialogSelectedItems collection. 
 For Each vrtSelectedItem In .SelectedItems 
 
 'vrtSelectedItem is aString that contains the path of each selected item. 
 'You can use any file I/O functions that you want to work with this path. 
 'This example displays the path in a message box. 
 MsgBox "The path is: " &amp; vrtSelectedItem 
 
 Next vrtSelectedItem 
 'The user pressed Cancel. 
 Else 
 End If 
 End With 
 
 'Set the object variable to Nothing. 
 Set fd = Nothing 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Execute](http://msdn.microsoft.com/library/filedialog-execute-method-office%28Office.15%29.aspx)|
|[Show](http://msdn.microsoft.com/library/filedialog-show-method-office%28Office.15%29.aspx)|

## Properties


<<<<<<< HEAD

|**Name**|
|:-----|
|[AllowMultiSelect](http://msdn.microsoft.com/library/filedialog-allowmultiselect-property-office%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/filedialog-application-property-office%28Office.15%29.aspx)|
|[ButtonName](http://msdn.microsoft.com/library/filedialog-buttonname-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/filedialog-creator-property-office%28Office.15%29.aspx)|
|[DialogType](http://msdn.microsoft.com/library/filedialog-dialogtype-property-office%28Office.15%29.aspx)|
|[FilterIndex](http://msdn.microsoft.com/library/filedialog-filterindex-property-office%28Office.15%29.aspx)|
|[Filters](http://msdn.microsoft.com/library/filedialog-filters-property-office%28Office.15%29.aspx)|
|[InitialFileName](http://msdn.microsoft.com/library/filedialog-initialfilename-property-office%28Office.15%29.aspx)|
|[InitialView](http://msdn.microsoft.com/library/filedialog-initialview-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/filedialog-item-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/filedialog-parent-property-office%28Office.15%29.aspx)|
|[SelectedItems](http://msdn.microsoft.com/library/filedialog-selecteditems-property-office%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/filedialog-title-property-office%28Office.15%29.aspx)|

## See also


#### Other resources

[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
=======
[Object Model Reference](reference-object-library-reference-for-office.md)
>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

