---
title: MailMergeDataSource Members (Word)
ms.prod: WORD
ms.assetid: a52f088c-2507-8f39-17b9-9b97c8a8ed7e
---


# MailMergeDataSource Members (Word)
Represents the mail merge data source in a mail merge operation.

Represents the mail merge data source in a mail merge operation.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Close](mailmergedatasource-close-method-word.md)|Closes the specified Mail Merge data source.|
|[FindRecord](mailmergedatasource-findrecord-method-word.md)|Searches the contents of the specified mail merge data source for text in a particular field. Returns  **True** if the search text is found. **Boolean** .|
|[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-word.md)|Marks all records in a mail merge data source as containing invalid data in an address field.|
|[SetAllIncludedFlags](mailmergedatasource-setallincludedflags-method-word.md)|Includes or excludes flagged records in a data source from a mail merge.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActiveRecord](mailmergedatasource-activerecord-property-word.md)|Returns or sets the active mail merge record. Can be either a valid record number in the query result or one of the  **WdMailMergeActiveRecord** constants.|
|[Application](mailmergedatasource-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[ConnectString](mailmergedatasource-connectstring-property-word.md)|Returns the connection string for the specified mail merge data source. Read-only  **String** .|
|[Creator](mailmergedatasource-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DataFields](mailmergedatasource-datafields-property-word.md)|Returns a  **[MailMergeDataFields](mailmergedatafields-object-word.md)** collection that represents the fields in the specified mail merge data source. Read-only.|
|[FieldNames](mailmergedatasource-fieldnames-property-word.md)|Returns a  **[MailMergeFieldNames](mailmergefieldnames-object-word.md)** collection that represents the names of all the fields in the specified mail merge data source. Read-only.|
|[FirstRecord](mailmergedatasource-firstrecord-property-word.md)|Returns or sets the number of the first record to be merged in a mail merge operation. Read/write  **Long** .|
|[HeaderSourceName](mailmergedatasource-headersourcename-property-word.md)|Returns the path and file name of the header source attached to the specified mail merge main document. Read-only  **String** .|
|[HeaderSourceType](mailmergedatasource-headersourcetype-property-word.md)|Returns a value that indicates the way the header source is being supplied for the mail merge operation. Read-only  **WdMailMergeDataSource** .|
|[Included](mailmergedatasource-included-property-word.md)| **True** if a record is included in a mail merge. Read/write **Boolean** .|
|[InvalidAddress](mailmergedatasource-invalidaddress-property-word.md)| **True** for Microsoft Word to mark a record in a mail merge data source if it contains invalid data in an address field. Read/write **Boolean** .|
|[InvalidComments](mailmergedatasource-invalidcomments-property-word.md)|If the  **[InvalidAddress](mailmergedatasource-invalidaddress-property-word.md)** property is **True** , returns or sets a **String** that describes an invalid address error. Read/write.|
|[LastRecord](mailmergedatasource-lastrecord-property-word.md)|Returns or sets the number of the last record to be merged in a mail merge operation. Read/write  **Long** .|
|[MappedDataFields](mailmergedatasource-mappeddatafields-property-word.md)|Returns a  **[MappedDataFields](mappeddatafields-object-word.md)** collection that represents the mapped data fields available in Microsoft Word.|
|[Name](mailmergedatasource-name-property-word.md)|Returns name of the specified object. Read-only  **String** .|
|[Parent](mailmergedatasource-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **MailMergeDataSource** object.|
|[QueryString](mailmergedatasource-querystring-property-word.md)|Returns or sets the query string (SQL statement) used to retrieve a subset of the data in a mail merge data source. Read/write  **String** .|
|[RecordCount](mailmergedatasource-recordcount-property-word.md)|Returns a  **Long** that represents the number of records in the data source. Read-only.|
|[TableName](mailmergedatasource-tablename-property-word.md)|Returns a  **String** with the SQL query used to retrieve the records from the data source file attached to a mail merge document. Read-only.|
|[Type](mailmergedatasource-type-property-word.md)|Returns the type of mail merge data source. Read-only  **[WdMailMergeDataSource](wdmailmergedatasource-enumeration-word.md)** .|

