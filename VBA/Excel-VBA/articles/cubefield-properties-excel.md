---
title: CubeField Properties (Excel)
ms.prod: EXCEL
ms.assetid: 1e04e6cd-da01-4bcc-a17d-1a441e1983ed
---


# CubeField Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllItemsVisible](cubefield-allitemsvisible-property-excel.md)| The **AllItemsVisible** property checks whether manual filtering is applied to a PivotField or CubeField. Read-only **Boolean** .|
|[Application](cubefield-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Caption](cubefield-caption-property-excel.md)|Returns a  **String** value that represents the label text for the cube field.|
|[Creator](cubefield-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[CubeFieldSubType](cubefield-cubefieldsubtype-property-excel.md)|Specifies the type of a CubeField. Read-only.|
|[CubeFieldType](cubefield-cubefieldtype-property-excel.md)|Indicates whether the OLAP cube field is a hierarchy field or a measure field. Can be one of the  **[XlCubeFieldType](xlcubefieldtype-enumeration-excel.md)** constants.|
|[CurrentPageName](cubefield-currentpagename-property-excel.md)|Returns or sets the page name for a CubeField. Read/write  **String** .|
|[DragToColumn](cubefield-dragtocolumn-property-excel.md)| **True** if the specified field can be dragged to the column position. The default value is **True** . Read/write **Boolean** .|
|[DragToData](cubefield-dragtodata-property-excel.md)| **True** if the specified field can be dragged to the data position. The default value is **True** . Read/write **Boolean**|
|[DragToHide](cubefield-dragtohide-property-excel.md)| **True** if the field can be hidden by being dragged off the PivotTable report. The default value is **True** . Read/write **Boolean** .|
|[DragToPage](cubefield-dragtopage-property-excel.md)| **True** if the field can be dragged to the page position. The default value is **True** . Read/write **Boolean** .|
|[DragToRow](cubefield-dragtorow-property-excel.md)| **True** if the field can be dragged to the row position. The default value is **True** . Read/write **Boolean** .|
|[EnableMultiplePageItems](cubefield-enablemultiplepageitems-property-excel.md)|Set to  **True** to allow multiple items in the page field area for OLAP PivotTables to be selected. The default value is **False** . Read/write **Boolean** .|
|[FlattenHierarchies](cubefield-flattenhierarchies-property-excel.md)|Returns or sets whether items from all levels of hierarchies in a named set cube field are displayed in the same field of a PivotTable report based on an OLAP cube. Read/write|
|[HasMemberProperties](cubefield-hasmemberproperties-property-excel.md)|Returns  **True** when there are member properties specified to be displayed for the cube field. Read-only **Boolean** .|
|[HierarchizeDistinct](cubefield-hierarchizedistinct-property-excel.md)|Returns or sets whether to order and remove duplicates when displaying the specified named set in a PivotTable report based on an OLAP cube. Read/write|
|[IncludeNewItemsInFilter](cubefield-includenewitemsinfilter-property-excel.md)|The  **IncludeNewItemsInFilter** property is used to track included/excluded items in OLAP PivotTables. Read/write.|
|[IsDate](cubefield-isdate-property-excel.md)|Returns  **True** if the CubeField is a date. Read-only **Boolean** .|
|[LayoutForm](cubefield-layoutform-property-excel.md)|Returns or sets the way the specified PivotTable items appear—in table format or in outline format. Read/write  **[XlLayoutFormType](xllayoutformtype-enumeration-excel.md)** .|
|[LayoutSubtotalLocation](cubefield-layoutsubtotallocation-property-excel.md)|Returns or sets the position of the PivotTable field subtotals in relation to (either above or below) the specified field. Read/write  **[XlSubtototalLocationType](xlsubtototallocationtype-enumeration-excel.md)** .|
|[Name](cubefield-name-property-excel.md)|Returns a  **String** value that represents the name of the object.|
|[Orientation](cubefield-orientation-property-excel.md)|Returns or sets a  **[XlPivotFieldOrientation](xlpivotfieldorientation-enumeration-excel.md)** value that represents the location of the field in the specified PivotTable report.|
|[Parent](cubefield-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PivotFields](cubefield-pivotfields-property-excel.md)|Returns the  **[PivotFields](pivotfields-object-excel.md)** collection. This collection contains all PivotTable fields, including those that aren't currently visible on-screen. Read-only **PivotFields** object.|
|[Position](cubefield-position-property-excel.md)|Returns or sets a  **Long** value that represents the position of the hierarchy field on the PivotTable report when it's dragged from the field well.|
|[ShowInFieldList](cubefield-showinfieldlist-property-excel.md)|When set to  **True** (default), a **CubeField** object will be shown in the field list. Read/write **Boolean** .|
|[TreeviewControl](cubefield-treeviewcontrol-property-excel.md)|Returns the  **[TreeviewControl](treeviewcontrol-object-excel.md)** object of the **[CubeField](cubefield-object-excel.md)** object, representing the cube manipulation control of an OLAP-based PivotTable report. Read-only.|
|[Value](cubefield-value-property-excel.md)|Returns a  **String** value that represents the name of the specified field.|

