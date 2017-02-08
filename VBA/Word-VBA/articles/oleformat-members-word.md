---
title: OLEFormat Members (Word)
ms.prod: WORD
ms.assetid: 62aae4c1-c2c6-fbf7-193d-c078ea88a527
---


# OLEFormat Members (Word)
Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field.

Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](oleformat-activate-method-word.md)|Activates the specified  **OLEFormat** object.|
|[ActivateAs](oleformat-activateas-method-word.md)|Sets the Windows registry value that determines the default application used to activate the specified OLE object.|
|[ConvertTo](oleformat-convertto-method-word.md)|Converts the specified OLE object from one class to another, making it possible for you to edit the object in a different server application, or changing how the object is displayed in the document.|
|[DoVerb](oleformat-doverb-method-word.md)|Requests that an OLE object perform one of its available verbs ? the actions an OLE object takes to activate its contents.|
|[Edit](oleformat-edit-method-word.md)|Opens the specified OLE object for editing in the application it was created in.|
|[Open](oleformat-open-method-word.md)|Opens the specified  **OLEFormat** object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](oleformat-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[ClassType](oleformat-classtype-property-word.md)|Returns or sets the class type for the specified OLE object, picture, or field. Read/write  **String** .|
|[Creator](oleformat-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DisplayAsIcon](oleformat-displayasicon-property-word.md)| **True** if the specified object is displayed as an icon. Read/write **Boolean** .|
|[IconIndex](oleformat-iconindex-property-word.md)|Returns or sets the icon that is used when the  **[DisplayAsIcon](oleformat-displayasicon-property-word.md)** property is **True** . Read/write **Long** .|
|[IconLabel](oleformat-iconlabel-property-word.md)|Returns or sets the text displayed below the icon for an OLE object. Read/write  **String** .|
|[IconName](oleformat-iconname-property-word.md)|Returns or sets the program file in which the icon for an OLE object is stored. Read/write  **String** .|
|[IconPath](oleformat-iconpath-property-word.md)|Returns the path of the file in which the icon for an OLE object is stored. Read-only  **String** .|
|[Label](oleformat-label-property-word.md)|Returns a string that's used to identify the portion of the source file that's being linked. Read-only  **String** .|
|[Object](oleformat-object-property-word.md)|Returns an  **Object** that represents the specified OLE object's top-level interface. .|
|[Parent](oleformat-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **OLEFormat** object.|
|[PreserveFormattingOnUpdate](oleformat-preserveformattingonupdate-property-word.md)| **True** preserves formatting done in Microsoft Word to a linked OLE object, such as a table linked to a Microsoft Excel spreadsheet. Read/write **Boolean** .|
|[ProgID](oleformat-progid-property-word.md)|Returns the programmatic identifier (ProgID) for the specified OLE object. Read-only  **String** .|

