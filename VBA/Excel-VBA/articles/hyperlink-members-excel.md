---
title: Hyperlink Members (Excel)
ms.prod: EXCEL
ms.assetid: b0566d1c-404f-b79e-7770-e7189a1c817a
---


# Hyperlink Members (Excel)
Represents a hyperlink.

Represents a hyperlink.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddToFavorites](hyperlink-addtofavorites-method-excel.md)|Adds a shortcut to the workbook or hyperlink to the Favorites folder.|
|[CreateNewDocument](hyperlink-createnewdocument-method-excel.md)|Creates a new document linked to the specified hyperlink.|
|[Delete](hyperlink-delete-method-excel.md)|Deletes the object.|
|[Follow](hyperlink-follow-method-excel.md)|Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Address](hyperlink-address-property-excel.md)|Returns or sets a  **String** value that represents the address of the target document.|
|[Application](hyperlink-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](hyperlink-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[EmailSubject](hyperlink-emailsubject-property-excel.md)|Returns or sets the text string of the specified hyperlink's e-mail subject line. The subject line is appended to the hyperlink's address. Read/write  **String** .|
|[Name](hyperlink-name-property-excel.md)|Returns a  **String** value that represents the name of the object.|
|[Parent](hyperlink-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Range](hyperlink-range-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range the specified hyperlink is attached to.|
|[ScreenTip](hyperlink-screentip-property-excel.md)|Returns or sets the ScreenTip text for the specified hyperlink. Read/write  **String** .|
|[Shape](hyperlink-shape-property-excel.md)|Returns a  **[Shape](shape-object-excel.md)** object that represents the shape attached to the specified hyperlink.|
|[SubAddress](hyperlink-subaddress-property-excel.md)|Returns or sets the location within the document associated with the hyperlink. Read/write  **String** .|
|[TextToDisplay](hyperlink-texttodisplay-property-excel.md)|Returns or sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink. Read/write  **String** .|
|[Type](hyperlink-type-property-excel.md)|Returns a  **Long** value, containing a **[MsoHyperlinkType](msohyperlinktype-enumeration-office.md)** constant, that represents the location of the HTML frame.|

