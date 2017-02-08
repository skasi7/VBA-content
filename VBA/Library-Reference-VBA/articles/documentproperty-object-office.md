---
title: DocumentProperty Object (Office)
keywords: vbaof11.chm250002
f1_keywords:
- vbaof11.chm250002
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.DocumentProperty
ms.assetid: dd54ca3c-e0e2-4816-539a-17c5b4a928b1
---


# DocumentProperty Object (Office)

Represents a custom or built-in document property of a container document. The  **DocumentProperty** object is a member of the **DocumentProperties** collection.


## Remarks

Use the Microsoft Word  **Document.BuiltinDocumentProperties**( _index_ ) property, where _index_ is the name or index number of the built-in document property, to return a single **DocumentProperty** object that represents a specific built-in document property. Use the Microsoft Word **Document.CustomDocumentProperties**( _index_ ) property, where _index_ is the name or index number of the custom document property, to return a **DocumentProperty** object that represents a specific custom document property. The following list contains the names of all the available built-in document properties:


 **Note**  Properties of type  **msoPropertyTypeString** are limited in length to 255 characters.


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/documentproperty-delete-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/documentproperty-application-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/documentproperty-creator-property-office%28Office.15%29.aspx)|
|[LinkSource](http://msdn.microsoft.com/library/documentproperty-linksource-property-office%28Office.15%29.aspx)|
|[LinkToContent](http://msdn.microsoft.com/library/documentproperty-linktocontent-property-office%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/documentproperty-name-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/documentproperty-parent-property-office%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/documentproperty-type-property-office%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/documentproperty-value-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[DocumentProperty Object Members](http://msdn.microsoft.com/library/documentproperty-members-office%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
