---
title: XPath Members (Excel)
ms.prod: EXCEL
ms.assetid: 2b598d87-ea67-b3fa-fbae-bb8fd1e22274
---


# XPath Members (Excel)
Represents an XPath that has been mapped to a  **[Range](range-object-excel.md)** or **[ListColumn](listcolumn-object-excel.md)** object.

Represents an XPath that has been mapped to a  **[Range](range-object-excel.md)** or **[ListColumn](listcolumn-object-excel.md)** object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Clear](xpath-clear-method-excel.md)|Clears all XPath schema information for the mapped range. |
|[SetValue](xpath-setvalue-method-excel.md)|Maps the specified  **[XPath](xpath-object-excel.md)** object to a **[ListColumn](listcolumn-object-excel.md)** object or **[Range](range-object-excel.md)** collection. If the **XPath** object has previously been mapped to the **ListColumn** object or **Range** collection, the **SetValue** method sets the properties of the **XPath** object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](xpath-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](xpath-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Map](xpath-map-property-excel.md)|Returns an  **[XmlMap](xmlmap-object-excel.md)** object that represents the schema map that contains the specified **[XPath](xpath-object-excel.md)** object. Read-only.|
|[Parent](xpath-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Repeating](xpath-repeating-property-excel.md)| Returns **True** if the specified **[XPath](xpath-object-excel.md)** object is mapped to an XML list; returns **False** if the **XPath** object is mapped to a single cell. Read-only **Boolean** .|
|[Value](xpath-value-property-excel.md)|Returns a  **String** that represents the XPath for the specified object.|

