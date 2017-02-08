---
title: XMLNamespaces Members (Word)
keywords: vbawd10.chm3799
f1_keywords:
- vbawd10.chm3799
ms.prod: WORD
ms.assetid: f11a6cc1-f33d-e1ab-870c-aa5857d66797
---


# XMLNamespaces Members (Word)

A collection of  **XMLNamespace** objects that represents the entire collection of schemas in the Schema Library.


## Remarks

In Microsoft OfficeWord, you can access the Schema Library from the  **XML Schema** tab in the **Templates and Add-ins** dialog box. The Schema Library represents schemas installed on a user's computer that a user has applied to a Word document or that a user has explicitly added to the Schema Library by using the **Schema Library** dialog box.

Use the  **Item** method of the **XMLNamespaces** collection to return an individual **XMLNameSpace** object. The index value of the **Item** method can be either a **Long** , which indicates the position of the schema in the Schema Library, or a **String** , which represents the name of the schema as returned by using the **URI** property (the TargetNamespace setting defined in the schema).

The following example attaches a schema named SimpleSample to the active document.




```vb
Sub ApplySampleSchema() 
 Dim objSchema As XMLNamespace 
 
 For Each objSchema In Application.XMLNamespaces 
 If objSchema.URI = "SimpleSample" Then 
 objSchema.AttachToDocument ActiveDocument 
 Exit For 
 End If 
 Next 
End Sub
```


 **Note**  The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). For more information, refer to the Smart Document SDK on the Microsoft Developer Network (MSDN) Web site.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)

