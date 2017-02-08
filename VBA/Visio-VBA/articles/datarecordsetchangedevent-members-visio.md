---
title: DataRecordsetChangedEvent Members (Visio)
ms.prod: VISIO
ms.assetid: bcf0589c-717b-26b9-54f0-52b7d6303220
---


# DataRecordsetChangedEvent Members (Visio)
Passed by Microsoft Visio as the pSubjectObj argument to the  **[VisEventProc](iviseventproc-viseventproc-method-visio.md)** method of the **[IVisEventProc](iviseventproc-object-visio.md)** interface when events related to refreshing a data recordset fire.

Passed by Microsoft Visio as the pSubjectObj argument to the  **[VisEventProc](iviseventproc-viseventproc-method-visio.md)** method of the **[IVisEventProc](iviseventproc-object-visio.md)** interface when events related to refreshing a data recordset fire.


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](datarecordsetchangedevent-application-property-visio.md)|Returns the instance of Microsoft Visio associated with a  **DataRecordsetChangedEvent** object. Read-only.|
|[DataColumnsAdded](datarecordsetchangedevent-datacolumnsadded-property-visio.md)|After data in a data recordset are refreshed, returns an array of names of data columns newly added to the data recordset as a result of the refresh operation. Read-only.|
|[DataColumnsChanged](datarecordsetchangedevent-datacolumnschanged-property-visio.md)|Returns an array of names of data columns in a data recordset whose types have changed as a result of data in the data recordset being refreshed. Read-only.|
|[DataColumnsDeleted](datarecordsetchangedevent-datacolumnsdeleted-property-visio.md)|After data in a data recordset are refreshed, returns an array of names of data columns deleted from the data recordset as a result of the refresh operation. Read-only.|
|[DataRecordset](datarecordsetchangedevent-datarecordset-property-visio.md)|Returns the  **DataRecordset** object associated with the **DataRecordsetChanged** event that fires when data in the data recordset are refreshed. Read-only.|
|[DataRowsAdded](datarecordsetchangedevent-datarowsadded-property-visio.md)|Returns an array of IDs of data rows newly added to the data recordset as a result of a data-refresh operation. Read-only.|
|[DataRowsDeleted](datarecordsetchangedevent-datarowsdeleted-property-visio.md)|Returns an array of IDs of data rows deleted from the data recordset as a result of a data-refresh operation. Read-only.|
|[ObjectType](datarecordsetchangedevent-objecttype-property-visio.md)|Returns  **visObjTypeDataRecordsetChangedEvent** , the type of a **DataRecordsetChangedEvent** object. Read-only.|
|[Stat](datarecordsetchangedevent-stat-property-visio.md)|Returns status information for an object. Read-only.|

