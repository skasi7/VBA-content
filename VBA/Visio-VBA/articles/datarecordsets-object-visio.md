---
title: DataRecordsets Object (Visio)
keywords: vis_sdr.chm61000
f1_keywords:
- vis_sdr.chm61000
ms.prod: VISIO
api_name:
- Visio.DataRecordsets
ms.assetid: edf6d0dc-2f16-eee0-fd4c-ec4c9409179e
---


# DataRecordsets Object (Visio)

The collection of  **DataRecordset** objects associated with a **Document** object.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the  **DataRecordsets** collection is **[Item](http://msdn.microsoft.com/library/datarecordsets-item-property-visio%28Office.15%29.aspx)**.

Every Visio  **Document** object has a **DataRecordsets** collection, which is empty until you import data into Visio. To connect a Visio document to a data source, you add a **DataRecordset** object to the **DataRecordsets** collection of the document.

To add a  **DataRecordset** object to the **DataRecordsets** collection, you can use one of the following three methods, depending on the type of data source you want to connect to (OLEDB/ODBC or XML) and how you want to pass connection string and query command strings to Visio. By using the




-  **[DataRecordsets.Add](http://msdn.microsoft.com/library/datarecordsets-add-method-visio%28Office.15%29.aspx)** method, you can connect to an OLEDB or ODBC data source and pass connection and query command string information to Visio directly as method parameters.
    
-  **[DataRecordsets.AddFromConnectionFile](http://msdn.microsoft.com/library/datarecordsets-addfromconnectionfile-method-visio%28Office.15%29.aspx)** method, you can connect to an OLEBD or ODBC data source by passing the method an Office Data Connection (ODC) file that contains the connection and query command string information you want to supply to Visio.
    
-  **[DataRecordsets.AddFromXML](http://msdn.microsoft.com/library/datarecordsets-addfromxml-method-visio%28Office.15%29.aspx)** method, you pass the method an ADO classic XML string that contains all the data that you want to include in the data recordset.
    


Once you have created a data recordset, the connection string and query command string associated with the data recordset are represented by the  **[DataConnection.ConnectionString](http://msdn.microsoft.com/library/dataconnection-connectionstring-property-visio%28Office.15%29.aspx)** and **[DataRecordset.CommandString](http://msdn.microsoft.com/library/datarecordset-commandstring-property-visio%28Office.15%29.aspx)** properties respectively.


## Events



|**Name**|
|:-----|
|[BeforeDataRecordsetDelete](http://msdn.microsoft.com/library/datarecordset-beforedatarecordsetdelete-event-visio%28Office.15%29.aspx)|
|[DataRecordsetChanged](http://msdn.microsoft.com/library/datarecordset-datarecordsetchanged-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/datarecordset-delete-method-visio%28Office.15%29.aspx)|
|[GetAllRefreshConflicts](http://msdn.microsoft.com/library/datarecordset-getallrefreshconflicts-method-visio%28Office.15%29.aspx)|
|[GetDataRowIDs](http://msdn.microsoft.com/library/datarecordset-getdatarowids-method-visio%28Office.15%29.aspx)|
|[GetMatchingRowsForRefreshConflict](http://msdn.microsoft.com/library/datarecordset-getmatchingrowsforrefreshconflict-method-visio%28Office.15%29.aspx)|
|[GetPrimaryKey](http://msdn.microsoft.com/library/datarecordset-getprimarykey-method-visio%28Office.15%29.aspx)|
|[GetRowData](http://msdn.microsoft.com/library/datarecordset-getrowdata-method-visio%28Office.15%29.aspx)|
|[Refresh](http://msdn.microsoft.com/library/datarecordset-refresh-method-visio%28Office.15%29.aspx)|
|[RefreshUsingXML](http://msdn.microsoft.com/library/datarecordset-refreshusingxml-method-visio%28Office.15%29.aspx)|
|[RemoveRefreshConflict](http://msdn.microsoft.com/library/datarecordset-removerefreshconflict-method-visio%28Office.15%29.aspx)|
|[SetPrimaryKey](http://msdn.microsoft.com/library/datarecordset-setprimarykey-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/datarecordset-application-property-visio%28Office.15%29.aspx)|
|[CommandString](http://msdn.microsoft.com/library/datarecordset-commandstring-property-visio%28Office.15%29.aspx)|
|[DataAsXML](http://msdn.microsoft.com/library/datarecordset-dataasxml-property-visio%28Office.15%29.aspx)|
|[DataColumns](http://msdn.microsoft.com/library/datarecordset-datacolumns-property-visio%28Office.15%29.aspx)|
|[DataConnection](http://msdn.microsoft.com/library/datarecordset-dataconnection-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/datarecordset-document-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/datarecordset-eventlist-property-visio%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/datarecordset-id-property-visio%28Office.15%29.aspx)|
|[LinkReplaceBehavior](http://msdn.microsoft.com/library/datarecordset-linkreplacebehavior-property-visio%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/datarecordset-name-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/datarecordset-objecttype-property-visio%28Office.15%29.aspx)|
|[RefreshInterval](http://msdn.microsoft.com/library/datarecordset-refreshinterval-property-visio%28Office.15%29.aspx)|
|[RefreshSettings](http://msdn.microsoft.com/library/datarecordset-refreshsettings-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/datarecordset-stat-property-visio%28Office.15%29.aspx)|
|[TimeRefreshed](http://msdn.microsoft.com/library/datarecordset-timerefreshed-property-visio%28Office.15%29.aspx)|

