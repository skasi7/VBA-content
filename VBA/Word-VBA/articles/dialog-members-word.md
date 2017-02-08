---
title: Dialog Members (Word)
ms.prod: WORD
ms.assetid: f5c755d5-9fdf-bfb4-2c4b-8999ae176635
---


# Dialog Members (Word)
Represents a built-in dialog box. The  **Dialog** object is a member of the **[Dialogs](dialogs-object-word.md)** collection. The **Dialogs** collection contains all the built-in dialog boxes in Word. You cannot create a new built-in dialog box or add one to the **Dialogs** collection.

Represents a built-in dialog box. The  **Dialog** object is a member of the **[Dialogs](dialogs-object-word.md)** collection. The **Dialogs** collection contains all the built-in dialog boxes in Word. You cannot create a new built-in dialog box or add one to the **Dialogs** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Display](dialog-display-method-word.md)|Displays the specified built-in Word dialog box until either the user closes it or the specified amount of time has passed. Returns a  **Long** that indicates which button was clicked to close the dialog box.|
|[Execute](dialog-execute-method-word.md)|Applies the current settings of a Microsoft Word dialog box.|
|[Show](dialog-show-method-word.md)|Displays and carries out actions initiated in the specified built-in Word dialog box. Returns a  **Long** that indicates which button was clicked to close the dialog box.|
|[Update](dialog-update-method-word.md)|Updates the values shown in a built-in Microsoft Word dialog box.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](dialog-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[CommandBarId](dialog-commandbarid-property-word.md)|Returns a  **Long** that represents the toolbar control id for a built-in Microsoft Word dialog box. Read-only.|
|[CommandName](dialog-commandname-property-word.md)|Returns the name of the procedure that displays the specified built-in dialog box. Read-only  **String** .|
|[Creator](dialog-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DefaultTab](dialog-defaulttab-property-word.md)|Returns or sets the active tab when the specified dialog box is displayed. Read/write  **WdWordDialogTab** .|
|[Parent](dialog-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Dialog** object.|
|[Type](dialog-type-property-word.md)|Returns the type of built-in Microsoft Word dialog box. Read-only  **[WdWordDialog](wdworddialog-enumeration-word.md)** .|

