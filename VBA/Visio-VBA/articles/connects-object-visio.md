---
title: Connects Object (Visio)
keywords: vis_sdr.chm10070
f1_keywords:
- vis_sdr.chm10070
ms.prod: VISIO
api_name:
- Visio.Connects
ms.assetid: 8ac06fd8-0bbb-e9df-a08c-d697c4ac238e
---


# Connects Object (Visio)

 Includes a **Connect** object for each connection between two shapes in a drawing, such as a line and a box in an organization chart.


## Remarks

The default property of a  **Connects** collection is **Item**.

Use the  **Connects** property of a **Shape** object to retrieve a **Connects** collection with a **Connect** object for every **Shape** object to which the indicated **Shape** object is connected (glued).

Use the  **FromConnects** property of a **Shape** object to retrieve a **Connects** collection with a **Connect** object for every **Shape** object that is connected (glued) to the indicated **Shape** object.

Use the  **Connects** property of a **Page** object to retrieve a **Connects** collection with an entry for every connection on the **Page** object.

Use the  **Connects** property of a **Master** object to retrieve a **Connects** collection with an entry for every connection in the **Master** object.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVConnects.GetEnumerator()** (to enumerate the **Connect** objects.)
    

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/connects-application-property-visio%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/connects-count-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/connects-document-property-visio%28Office.15%29.aspx)|
|[FromSheet](http://msdn.microsoft.com/library/connects-fromsheet-property-visio%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/connects-item-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/connects-objecttype-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/connects-stat-property-visio%28Office.15%29.aspx)|
|[ToSheet](http://msdn.microsoft.com/library/connects-tosheet-property-visio%28Office.15%29.aspx)|

