---
title: MappedDataField Members (Word)
ms.prod: WORD
ms.assetid: dd2aadd0-7211-73ff-88a1-f48a44948adf
---


# MappedDataField Members (Word)
A mapped data field is a field contained within Microsoft Word that represents commonly used name or address information, such as "First Name." If a data source contains a "First Name" field or a variation (such as "First_Name," "FirstName," "First," or "FName"), the field in the data source will automatically map to the corresponding mapped data field in Word. If a document or template is to be merged with more than one data source, mapped data fields make it unnecessary to reenter the fields into the document to agree with the field names in the database.

A mapped data field is a field contained within Microsoft Word that represents commonly used name or address information, such as "First Name." If a data source contains a "First Name" field or a variation (such as "First_Name," "FirstName," "First," or "FName"), the field in the data source will automatically map to the corresponding mapped data field in Word. If a document or template is to be merged with more than one data source, mapped data fields make it unnecessary to reenter the fields into the document to agree with the field names in the database.


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](mappeddatafield-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](mappeddatafield-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DataFieldIndex](mappeddatafield-datafieldindex-property-word.md)|Returns or sets a  **Long** that represents the corresponding field number in the mail merge data source to which a mapped data field maps. Read/write.|
|[DataFieldName](mappeddatafield-datafieldname-property-word.md)|Sets or returns a  **String** that represents the name of the field in the mail merge data source to which a mapped data field maps. Read/write.|
|[Index](mappeddatafield-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[Name](mappeddatafield-name-property-word.md)|Returns name of the specified object. Read-only  **String** .|
|[Parent](mappeddatafield-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **MappedDataField** object.|
|[Value](mappeddatafield-value-property-word.md)|Returns the contents of the mail merge data field or mapped data field for the current record. Read-only  **String** .|

