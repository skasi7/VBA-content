---
title: LinkFormat Members (Word)
ms.prod: WORD
ms.assetid: 028d048f-df8c-0dec-17f2-56f0d0a332c7
---


# LinkFormat Members (Word)
Represents the linking characteristics for an OLE object or picture.

Represents the linking characteristics for an OLE object or picture.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[BreakLink](linkformat-breaklink-method-word.md)|Breaks the link between the source file and the specified OLE object, picture, or linked field.|
|[Update](linkformat-update-method-word.md)|Updates the specified link format.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](linkformat-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoUpdate](linkformat-autoupdate-property-word.md)| **True** if the specified link is updated automatically when the container file is opened or when the source file is changed. Read/write **Boolean** .|
|[Creator](linkformat-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Locked](linkformat-locked-property-word.md)| **True** if a **Field** , **InlineShape** , or **Shape** object is locked to prevent automatic updating. Read/write **Boolean** .|
|[Parent](linkformat-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **LinkFormat** object.|
|[SavePictureWithDocument](linkformat-savepicturewithdocument-property-word.md)| **True** if the specified picture is saved with the document. Read/write **Boolean** .|
|[SourceFullName](linkformat-sourcefullname-property-word.md)|Returns or sets the path and name of the source file for the specified linked OLE object, picture, or field. Read/write  **String** .|
|[SourceName](linkformat-sourcename-property-word.md)|Returns the name of the source file for the specified linked OLE object, picture, or field. Read-only  **String** .|
|[SourcePath](linkformat-sourcepath-property-word.md)|Returns the path of the source file for the specified linked OLE object, picture, or field. Read-only  **String** .|
|[Type](linkformat-type-property-word.md)|Returns the link type. Read-only  **[WdLinkType](wdlinktype-enumeration-word.md)** .|

