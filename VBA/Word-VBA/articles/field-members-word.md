---
title: Field Members (Word)
ms.prod: WORD
ms.assetid: 6920f70a-3164-ce35-3b6d-01edb32fc02b
---


# Field Members (Word)
Represents a field. The  **Field** object is a member of the **Fields** collection. The **[Fields](fields-object-word.md)** collection represents the fields in a selection, range, or document.

Represents a field. The  **Field** object is a member of the **Fields** collection. The **[Fields](fields-object-word.md)** collection represents the fields in a selection, range, or document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Copy](field-copy-method-word.md)|Copies the specified field to the Clipboard.|
|[Cut](field-cut-method-word.md)|Removes the specified field from the document and places it on the Clipboard.|
|[Delete](field-delete-method-word.md)|Deletes the specified field.|
|[DoClick](field-doclick-method-word.md)|Clicks the specified field.|
|[Select](field-select-method-word.md)|Selects the specified field.|
|[Unlink](field-unlink-method-word.md)|Replaces the specified field with its most recent result.|
|[Update](field-update-method-word.md)|Updates the result of the field. Returns  **True** if the field is updated successfully.|
|[UpdateSource](field-updatesource-method-word.md)|Saves the changes made to the results of an INCLUDETEXT field back to the source document.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](field-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Code](field-code-property-word.md)|Returns a  **[Range](range-object-word.md)** object that represents a field's code. Read/write.|
|[Creator](field-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Data](field-data-property-word.md)|Returns or sets data in an ADDIN field. Read/write  **String** .|
|[Index](field-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[InlineShape](field-inlineshape-property-word.md)|Returns an  **[InlineShape](inlineshape-object-word.md)** object that represents the picture, OLE object, or ActiveX control that is the result of an INCLUDEPICTURE or EMBED field.|
|[Kind](field-kind-property-word.md)|Returns the type of link for a  **Field** object. Read-only **[WdFieldKind](wdfieldkind-enumeration-word.md)** .|
|[LinkFormat](field-linkformat-property-word.md)|Returns a  **LinkFormat** object that represents the link options of the specified field. Read/only.|
|[Locked](field-locked-property-word.md)| **True** if the specified field is locked. Read/write **Boolean** .|
|[Next](field-next-property-word.md)|Returns the next object in the collection. Read-only.|
|[OLEFormat](field-oleformat-property-word.md)|Returns an  **OLEFormat** object that represents the OLE characteristics (other than linking) for the specified field. Read-only.|
|[Parent](field-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Field** object.|
|[Previous](field-previous-property-word.md)|Returns the previous object in the collection. Read-only.|
|[Result](field-result-property-word.md)|Returns a  **Range** object that represents a field's result. Read/write.|
|[ShowCodes](field-showcodes-property-word.md)| **True** if field codes are displayed for the specified field instead of field results. Read/write **Boolean** .|
|[Type](field-type-property-word.md)|Returns the field type. Read-only  **[WdFieldType](wdfieldtype-enumeration-word.md)** .|

