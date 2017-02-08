---
title: MetaProperties Object (Office)
keywords: vbaof11.chm274000
f1_keywords:
- vbaof11.chm274000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.MetaProperties
ms.assetid: 957a6e06-3348-b180-3655-06ffbfb69e12
---


# MetaProperties Object (Office)

Represents a collection of properties describing the metadata stored in a document.


## Example

In the following example, a  **MetaProperties** object is passed to a validation function. The function then validates the value of a single property represented by its index and returns the result.


```
Function ValidateMetaProperty(ByVal metaProps As MetaProperties) As String 
Dim result As String 
 
result = metaProps(1).Validate 
ValidateMetaProperty = result 
End Function
```


## Methods



|**Name**|
|:-----|
|[GetItemByInternalName](http://msdn.microsoft.com/library/metaproperties-getitembyinternalname-method-office%28Office.15%29.aspx)|
|[Validate](http://msdn.microsoft.com/library/metaproperties-validate-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/metaproperties-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/metaproperties-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/metaproperties-creator-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/metaproperties-item-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/metaproperties-parent-property-office%28Office.15%29.aspx)|
|[SchemaXml](http://msdn.microsoft.com/library/metaproperties-schemaxml-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[MetaProperties Object Members](http://msdn.microsoft.com/library/metaproperties-members-office%28Office.15%29.aspx)
