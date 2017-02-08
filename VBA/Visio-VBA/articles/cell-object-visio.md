---
title: Cell Object (Visio)
keywords: vis_sdr.chm10045
f1_keywords:
- vis_sdr.chm10045
ms.prod: VISIO
api_name:
- Visio.Cell
ms.assetid: 06ac28a6-5749-6c70-94bf-c721e217f375
---


# Cell Object (Visio)

Holds a formula that evaluates to some value.


## Remarks

The default property of a  **Cell** object is **ResultIU**.

You can get or set a cell's formula or value. A cell belongs to a  **Shape**, **Style**, or **Row** object and represents a property of the shape, style, or row. For example, the height of a shape equals the value of the shape's Height cell.

A program can control a shape's appearance and behavior by working with the formulas in the shape's cells. You can visually inspect most of a shape's cells by opening the shape's ShapeSheet window. Use the  **Cells** or **CellsSRC** property of a **Shape** object to retrieve a **Cell** object. To retrieve a cell in a style, use the **Cells** property of a **Style** object.


## Events



|**Name**|
|:-----|
|[CellChanged](http://msdn.microsoft.com/library/cell-cellchanged-event-visio%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/cell-formulachanged-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[GlueTo](http://msdn.microsoft.com/library/cell-glueto-method-visio%28Office.15%29.aspx)|
|[GlueToPos](http://msdn.microsoft.com/library/cell-gluetopos-method-visio%28Office.15%29.aspx)|
|[Trigger](http://msdn.microsoft.com/library/cell-trigger-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/cell-application-property-visio%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/cell-column-property-visio%28Office.15%29.aspx)|
|[ContainingMasterID](http://msdn.microsoft.com/library/cell-containingmasterid-property-visio%28Office.15%29.aspx)|
|[ContainingPageID](http://msdn.microsoft.com/library/cell-containingpageid-property-visio%28Office.15%29.aspx)|
|[ContainingRow](http://msdn.microsoft.com/library/cell-containingrow-property-visio%28Office.15%29.aspx)|
|[Dependents](http://msdn.microsoft.com/library/cell-dependents-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/cell-document-property-visio%28Office.15%29.aspx)|
|[Error](http://msdn.microsoft.com/library/cell-error-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/cell-eventlist-property-visio%28Office.15%29.aspx)|
|[Formula](http://msdn.microsoft.com/library/cell-formula-property-visio%28Office.15%29.aspx)|
|[FormulaForce](http://msdn.microsoft.com/library/cell-formulaforce-property-visio%28Office.15%29.aspx)|
|[FormulaForceU](http://msdn.microsoft.com/library/cell-formulaforceu-property-visio%28Office.15%29.aspx)|
|[FormulaU](http://msdn.microsoft.com/library/cell-formulau-property-visio%28Office.15%29.aspx)|
|[InheritedFormulaSource](http://msdn.microsoft.com/library/cell-inheritedformulasource-property-visio%28Office.15%29.aspx)|
|[InheritedValueSource](http://msdn.microsoft.com/library/cell-inheritedvaluesource-property-visio%28Office.15%29.aspx)|
|[IsConstant](http://msdn.microsoft.com/library/cell-isconstant-property-visio%28Office.15%29.aspx)|
|[IsInherited](http://msdn.microsoft.com/library/cell-isinherited-property-visio%28Office.15%29.aspx)|
|[LocalName](http://msdn.microsoft.com/library/cell-localname-property-visio%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/cell-name-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/cell-objecttype-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/cell-persistsevents-property-visio%28Office.15%29.aspx)|
|[Precedents](http://msdn.microsoft.com/library/cell-precedents-property-visio%28Office.15%29.aspx)|
|[Result](http://msdn.microsoft.com/library/cell-result-property-visio%28Office.15%29.aspx)|
|[ResultForce](http://msdn.microsoft.com/library/cell-resultforce-property-visio%28Office.15%29.aspx)|
|[ResultFromInt](http://msdn.microsoft.com/library/cell-resultfromint-property-visio%28Office.15%29.aspx)|
|[ResultFromIntForce](http://msdn.microsoft.com/library/cell-resultfromintforce-property-visio%28Office.15%29.aspx)|
|[ResultInt](http://msdn.microsoft.com/library/cell-resultint-property-visio%28Office.15%29.aspx)|
|[ResultIU](http://msdn.microsoft.com/library/cell-resultiu-property-visio%28Office.15%29.aspx)|
|[ResultIUForce](http://msdn.microsoft.com/library/cell-resultiuforce-property-visio%28Office.15%29.aspx)|
|[ResultStr](http://msdn.microsoft.com/library/cell-resultstr-property-visio%28Office.15%29.aspx)|
|[ResultStrU](http://msdn.microsoft.com/library/cell-resultstru-property-visio%28Office.15%29.aspx)|
|[Row](http://msdn.microsoft.com/library/cell-row-property-visio%28Office.15%29.aspx)|
|[RowName](http://msdn.microsoft.com/library/cell-rowname-property-visio%28Office.15%29.aspx)|
|[RowNameU](http://msdn.microsoft.com/library/cell-rownameu-property-visio%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/cell-section-property-visio%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/cell-shape-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/cell-stat-property-visio%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/cell-style-property-visio%28Office.15%29.aspx)|
|[Units](http://msdn.microsoft.com/library/cell-units-property-visio%28Office.15%29.aspx)|

