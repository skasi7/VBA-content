---
title: CubeFields Members (Excel)
ms.prod: EXCEL
ms.assetid: 92d974bf-4956-fd8e-60c7-d0edd3cee734
---


# CubeFields Members (Excel)
A collection of all  **[CubeField](cubefield-object-excel.md)** objects in a PivotTable report that is based on an OLAP cube. Each **CubeField** object represents a hierarchy or measure field from the cube.

A collection of all  **[CubeField](cubefield-object-excel.md)** objects in a PivotTable report that is based on an OLAP cube. Each **CubeField** object represents a hierarchy or measure field from the cube.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddSet](cubefields-addset-method-excel.md)|Adds a new  **[CubeField](cubefield-object-excel.md)** object to the **[CubeFields](cubefields-object-excel.md)** collection. The **CubeField** object corresponds to a set defined on the Online Analytical Processing (OLAP) provider for the cube.|
|[GetMeasure](cubefields-getmeasure-method-excel.md)|Given an attribute hierarchy, returns an implicit measure for the given function that corresponds to this attribute. If an "implicit measure" does not exist, a new implicit measure is created and added to the [CubeFields Object (Excel)](cubefields-object-excel.md) collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](cubefields-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Count](cubefields-count-property-excel.md)|Returns a  **Long** value that represents the number of objects in the collection.|
|[Creator](cubefields-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Item](cubefields-item-property-excel.md)|Returns a single object from a collection.|
|[Parent](cubefields-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|

