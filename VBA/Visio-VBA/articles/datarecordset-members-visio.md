---
title: DataRecordset Members (Visio)
ms.prod: VISIO
ms.assetid: a7600315-638c-3cc9-9c3c-e152b44f7e7e
---


# DataRecordset Members (Visio)
Stores, formats, refreshes, and exposes data queried from a database in Microsoft Visio.

Stores, formats, refreshes, and exposes data queried from a database in Microsoft Visio.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeDataRecordsetDelete](datarecordset-beforedatarecordsetdelete-event-visio.md)|Occurs before a  **DataRecordset** object is deleted from the **DataRecordsets** collection.|
|[DataRecordsetChanged](datarecordset-datarecordsetchanged-event-visio.md)|Occurs when a data recordset changes as a result of being refreshed.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](datarecordset-delete-method-visio.md)|Deletes the  **[DataRecordset](datarecordset-object-visio.md)** object from the **[DataRecordsets](datarecordsets-object-visio.md)** collection of the document. .|
|[GetAllRefreshConflicts](datarecordset-getallrefreshconflicts-method-visio.md)|Returns an array that contains shapes linked to data rows that have non-resolved conflicts after a data recordset is refreshed. .|
|[GetDataRowIDs](datarecordset-getdatarowids-method-visio.md)|Gets an array of the IDs of all the rows in the data recordset.|
|[GetMatchingRowsForRefreshConflict](datarecordset-getmatchingrowsforrefreshconflict-method-visio.md)|Returns an array of the row IDs of data-recordset rows linked to a shape that are in conflict after the data recordset is refreshed.|
|[GetPrimaryKey](datarecordset-getprimarykey-method-visio.md)|Gets the primary key setting and the name of the primary key column or columns for the data recordset.|
|[GetRowData](datarecordset-getrowdata-method-visio.md)|Gets the data in all columns in the specified row.|
|[Refresh](datarecordset-refresh-method-visio.md)|Executes the query string associated with the connected (non-XML-based)  **[DataRecordset](datarecordset-object-visio.md)** and updates linked shapes with new data from the data source returned by the query.|
|[RefreshUsingXML](datarecordset-refreshusingxml-method-visio.md)|Updates linked shapes with data contained in the string that conforms to the ADO classic XML schema passed to the method as a parameter.|
|[RemoveRefreshConflict](datarecordset-removerefreshconflict-method-visio.md)|Clears information about a conflict for a data-linked shape from the current document.|
|[SetPrimaryKey](datarecordset-setprimarykey-method-visio.md)|Sets the primary key setting value and the name of the primary key column or columns for the data recordset.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](datarecordset-application-property-visio.md)|Returns the instance of Microsoft Visio associated with a  **DataRecordset** object. Read-only.|
|[CommandString](datarecordset-commandstring-property-visio.md)|Gets or sets the command string used to query the data source to create a data recordset or refresh an existing one. Read/write.|
|[DataAsXML](datarecordset-dataasxml-property-visio.md)|Returns an XML string that fully describes a data recordset and conforms to the Microsoft ActiveX® Data Objects (ADO) classic XML schema. Read-only.|
|[DataColumns](datarecordset-datacolumns-property-visio.md)|Returns the  **[DataColumns](datacolumns-object-visio.md)** collection associated with the **DataRecordset** object. Read-only.|
|[DataConnection](datarecordset-dataconnection-property-visio.md)|Returns the  **[DataConnection](dataconnection-object-visio.md)** object associated with the **DataRecordset** object. Read-only.|
|[Document](datarecordset-document-property-visio.md)|Gets the  **Document** object that contains the **DataRecordset** object. Read-only.|
|[EventList](datarecordset-eventlist-property-visio.md)|Returns the  **[EventList](eventlist-object-visio.md)** collection of the **DataRecordset** object. Read-only.|
|[ID](datarecordset-id-property-visio.md)|Gets the unique identifier of the  **DataRecordset** object assigned by Visio. Read-only.|
|[LinkReplaceBehavior](datarecordset-linkreplacebehavior-property-visio.md)|Gets or sets how existing links between shapes and data rows are handled when methods that link shapes to data is called. Read/write.|
|[Name](datarecordset-name-property-visio.md)|Gets or sets the display name of the data recordset. Read/write.|
|[ObjectType](datarecordset-objecttype-property-visio.md)|Returns  **visObjTypeDataRecordset** , the type of a **DataRecordset** object. Read-only.|
|[RefreshInterval](datarecordset-refreshinterval-property-visio.md)|Gets or sets how often Microsoft Visio automatically refreshes the data recordset. Read/write.|
|[RefreshSettings](datarecordset-refreshsettings-property-visio.md)|Gets and sets options that determine how the data recordset is refreshed. Read/write.|
|[Stat](datarecordset-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[TimeRefreshed](datarecordset-timerefreshed-property-visio.md)|Returns the date and time the data recordset was last refreshed. Read-only.|

