---
title: CustomXMLPart Object (Office)
keywords: vbaof11.chm297000
f1_keywords:
- vbaof11.chm297000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CustomXMLPart
ms.assetid: a4f90bac-01d6-bba4-f64b-a64e2b122cfd
---


# CustomXMLPart Object (Office)

Represents a single  **CustomXMLPart** in a **CustomXMLParts** collection.


## Example

The following example adds a part to a  **CustomXMLPart** object.


```
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[CustomXMLPart Object Members](http://msdn.microsoft.com/library/customxmlpart-members-office%28Office.15%29.aspx)
