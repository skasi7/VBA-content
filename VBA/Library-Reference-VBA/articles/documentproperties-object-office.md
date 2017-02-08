---
title: DocumentProperties Object (Office)
keywords: vbaof11.chm250010
f1_keywords:
- vbaof11.chm250010
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.DocumentProperties
ms.assetid: 90d42786-7d9a-b604-dbdf-88db41cbe69b
---


# DocumentProperties Object (Office)

A collection of  **DocumentProperty** objects. Each **DocumentProperty** object represents a built-in or custom property of a container document.


## Remarks

Use the ** Add** method to create a new custom property and add it to the **DocumentProperties** collection. You cannot use the **Add** method to create a built-in document property.

Use  **BuiltinDocumentProperties(index)**, where _index_ is the index number of the built-in document property, to return a single **DocumentProperty** object that represents a specific built-in document property. Use **CustomDocumentProperties(index)**, where _index_ is the number of the custom document property, to return a **DocumentProperty** object that represents a specific custom document property.


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/documentproperties-add-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/documentproperties-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/documentproperties-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/documentproperties-creator-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/documentproperties-item-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/documentproperties-parent-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[DocumentProperties Object Members](http://msdn.microsoft.com/library/documentproperties-members-office%28Office.15%29.aspx)
