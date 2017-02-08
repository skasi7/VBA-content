---
title: Categories Object (Outlook)
keywords: vbaol11.chm3178
f1_keywords:
- vbaol11.chm3178
ms.prod: OUTLOOK
api_name:
- Outlook.Categories
ms.assetid: 319efa26-269d-9f2f-c8ec-33082e80a9e2
---


# Categories Object (Outlook)

Represents the collection of  **[Category](http://msdn.microsoft.com/library/category-object-outlook%28Office.15%29.aspx)** objects that define the Master Category List for a namespace.


## Remarks

Microsoft Outlook provides a categorization system by which Outlook items can be easily identified and grouped into user-defined categories. The  **Categories** object represents the set of user-defined categories available to the user of a given mailbox.

Use the  **[Categories](http://msdn.microsoft.com/library/namespace-categories-property-outlook%28Office.15%29.aspx)** property of the **[NameSpace](namespace-object-outlook.md)** object to obtain a **Categories** object reference, representing the Master Category List for that namespace.

Use the  **[Add](http://msdn.microsoft.com/library/categories-add-method-outlook%28Office.15%29.aspx)** method to create a new **Category** object and append it to the collection. Use the **[Item](http://msdn.microsoft.com/library/categories-item-method-outlook%28Office.15%29.aspx)** method to obtain a **Category** object reference for an existing category, and the **[Remove](http://msdn.microsoft.com/library/categories-remove-method-outlook%28Office.15%29.aspx)** method to remove a **Category** object from the collection. Use the **[Count](http://msdn.microsoft.com/library/categories-count-property-outlook%28Office.15%29.aspx)** property to return the number of categories contained in the collection.


## Example

The following Visual Basic for Applications (VBA) example displays a dialog box containing the names and identifiers for each  **Category** object contained in the **Categories** collection associated with the default **[NameSpace](namespace-object-outlook.md)** object.


```
Private Sub ListCategoryIDs() 
 Dim objNameSpace As NameSpace 
 Dim objCategory As Category 
 Dim strOutput As String 
 
 ' Obtain a NameSpace object reference. 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 ' Check if the Categories collection for the Namespace 
 ' contains one or more Category objects. 
 If objNameSpace.Categories.Count > 0 Then 
 
 ' Enumerate the Categories collection. 
 For Each objCategory In objNameSpace.Categories 
 
 ' Add the name and ID of the Category object to 
 ' the output string. 
 strOutput = strOutput &amp; objCategory.Name &amp; _ 
 ": " &amp; objCategory.CategoryID &amp; vbCrLf 
 Next 
 End If 
 
 ' Display the output string. 
 MsgBox strOutput 
 
 ' Clean up. 
 Set objCategory = Nothing 
 Set objNameSpace = Nothing 
 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/categories-add-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/categories-item-method-outlook%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/categories-remove-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/categories-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/categories-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/categories-count-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/categories-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/categories-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Categories Object Members](http://msdn.microsoft.com/library/categories-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
