---
title: DataRecordsets Members (Visio)
ms.prod: VISIO
ms.assetid: eaaf3f60-a4f0-0e0c-5c74-cea72ebb2204
---


# DataRecordsets Members (Visio)
The collection of  **DataRecordset** objects associated with a **Document** object.

The collection of  **DataRecordset** objects associated with a **Document** object.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeDataRecordsetDelete](datarecordsets-beforedatarecordsetdelete-event-visio.md)|Occurs before a  **DataRecordset** object is deleted from the **DataRecordsets** collection.|
|[DataRecordsetAdded](datarecordsets-datarecordsetadded-event-visio.md)|Occurs when a  **DataRecordset** object is added to a **DataRecordsets** collection.|
|[DataRecordsetChanged](datarecordsets-datarecordsetchanged-event-visio.md)|Occurs when a data recordset changes as a result of being refreshed.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](datarecordsets-add-method-visio.md)|Adds a  **[DataRecordset](datarecordset-object-visio.md)** object to the **[DataRecordsets](datarecordsets-object-visio.md)** collection by connecting to and retrieving data from an OLEDB or ODBC data source.|
|[AddFromConnectionFile](datarecordsets-addfromconnectionfile-method-visio.md)|Adds a  **[DataRecordset](datarecordset-object-visio.md)** object to the **[DataRecordsets](datarecordsets-object-visio.md)** collection by using the connection and query information contained in an Office Data Connection (ODC) file to connect to and retrieve data from an OLEDB or ODBC data source.|
|[AddFromXML](datarecordsets-addfromxml-method-visio.md)|Adds a  **[DataRecordset](datarecordset-object-visio.md)** object to the **[DataRecordsets](datarecordsets-object-visio.md)** collection, and fills the resulting data recordset with data supplied in the form of an XML string.|
|[GetLastDataError](datarecordsets-getlastdataerror-method-visio.md)|Gets the Active X Data Objects (ADO) error code, ADO description, and data recordset ID associated with an error that results from adding a new data recordset or refreshing the data in an existing one.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](datarecordsets-application-property-visio.md)|Returns the instance of Microsoft Visio associated with the  **DataRecordsets** collection. Read-only.|
|[Count](datarecordsets-count-property-visio.md)|Returns the number of  **DataRecordset** objects in the **DataRecordsets** collection. Read-only.|
|[Document](datarecordsets-document-property-visio.md)|Gets the  **Document** object that contains the **DataRecordsets** collection. Read-only.|
|[EventList](datarecordsets-eventlist-property-visio.md)|Returns the  **[EventList](eventlist-object-visio.md)** collection of the **DataRecordsets** collection. Read-only.|
|[Item](datarecordsets-item-property-visio.md)|Returns the  **DataRecordset** object at the specified index position in the **DataRecordsets** collection. Read-only.|
|[ItemFromID](datarecordsets-itemfromid-property-visio.md)|Returns a  **DataRecordset** object from the **DataRecordsets** collection by using the unique ID of the object. Read-only.|
|[ObjectType](datarecordsets-objecttype-property-visio.md)|Returns  **visObjTypeDataRecordsets** , the type of a **DataRecordsets** object. Read-only.|
|[Stat](datarecordsets-stat-property-visio.md)|Returns status information for an object. Read-only.|

