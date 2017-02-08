---
title: Inspector Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: acd3e13f-4727-7966-d2a5-a95e4528425c
---


# Inspector Members (Outlook)
Represents the window in which an Outlook item is displayed.

Represents the window in which an Outlook item is displayed.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Activate](inspector-activate-event-outlook.md)|Occurs when an inspector becomes the active window, either as a result of user action or through program code. |
|[AttachmentSelectionChange](inspector-attachmentselectionchange-event-outlook.md)|Occurs when the user selects a different or additional attachment of an item in the active inspector programmatically or by interacting with the user interface.|
|[BeforeMaximize](inspector-beforemaximize-event-outlook.md)|Occurs when an inspector is maximized by the user.|
|[BeforeMinimize](inspector-beforeminimize-event-outlook.md)|Occurs when the active inspector is minimized by the user.|
|[BeforeMove](inspector-beforemove-event-outlook.md)|Occurs when the  **[Inspector](inspector-object-outlook.md)** is moved by the user.|
|[BeforeSize](inspector-beforesize-event-outlook.md)|Occurs when the user sizes the current  **[Inspector](inspector-object-outlook.md)** .|
|[Close](inspector-close-event-outlook.md)|Occurs when the inspector associated with a Microsoft Outlook item is being closed.|
|[Deactivate](inspector-deactivate-event-outlook.md)|Occurs when an inspector stops being the active window, either as a result of user action or through program code.|
|[PageChange](inspector-pagechange-event-outlook.md)|Occurs when the active form page changes, either programmatically or by user action, on an [Inspector](inspector-object-outlook.md) object.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](inspector-activate-method-outlook.md)|Activates an inspector window by bringing it to the foreground and setting keyboard focus.|
|[Close](inspector-close-method-outlook.md)|Closes the  **[Inspector](inspector-object-outlook.md)** and optionally saves changes to the displayed Outlook item.|
|[Display](inspector-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.|
|[HideFormPage](inspector-hideformpage-method-outlook.md)|Hides a form page or a form region in the inspector.|
|[IsWordMail](inspector-iswordmail-method-outlook.md)|Determines whether the mail message associated with an inspector is displayed in an Outlook  **[Inspector](inspector-object-outlook.md)** or in Microsoft Word.|
|[NewFormRegion](inspector-newformregion-method-outlook.md)|Opens a new page in design mode in the inspector for a new form region.|
|[OpenFormRegion](inspector-openformregion-method-outlook.md)|Opens a page in design mode in the inspector for the specified form region.|
|[SaveFormRegion](inspector-saveformregion-method-outlook.md)|Saves the specified page in design mode in the inspector to the specified file.|
|[SetControlItemProperty](inspector-setcontrolitemproperty-method-outlook.md)|Binds a built-in property or custom property to a control in an inspector. |
|[SetCurrentFormPage](inspector-setcurrentformpage-method-outlook.md)|Displays the specified form page or form region in the inspector.|
|[SetSchedulingStartTime](inspector-setschedulingstarttime-method-outlook.md)|Sets the start time for a meeting item in the free/busy grid on the  **Scheduling Assistant** tab of the inspector.|
|[ShowFormPage](inspector-showformpage-method-outlook.md)|Displays a button in the  **Show** group of the Microsoft Office Fluent ribbon for the inspector, clicking which shows the page or form region specified by _PageName_.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](inspector-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[AttachmentSelection](inspector-attachmentselection-property-outlook.md)|Returns an  **[AttachmentSelection](attachmentselection-object-outlook.md)** object consisting of one or more attachments that are selected in the inspector. Read-only.|
|[Caption](inspector-caption-property-outlook.md)|Returns a  **String** representing the title. Read-only.|
|[Class](inspector-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[CurrentItem](inspector-currentitem-property-outlook.md)|Returns an  **Object** representing the current item being displayed in the inspector. Read-only.|
|[EditorType](inspector-editortype-property-outlook.md)|Returns an  **[OlEditorType](oleditortype-enumeration-outlook.md)** constant indicating the type of editor. Read-only.|
|[Height](inspector-height-property-outlook.md)|Returns or sets a  **Long** specifying the height (in pixels) of the inspector window. Read/write.|
|[Left](inspector-left-property-outlook.md)|Returns or sets a  **Long** specifying the position (in pixels) of the left vertical edge of an inspector window from the edge of the screen. Read/write.|
|[ModifiedFormPages](inspector-modifiedformpages-property-outlook.md)|Returns the  **[Pages](pages-object-outlook.md)** collection that represents all the pages for the item in the inspector. Read-only.|
|[Parent](inspector-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Session](inspector-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Top](inspector-top-property-outlook.md)|Returns or sets a  **Long** indicating the position (in pixels) of the top horizontal edge of an inspector window from the edge of the screen. Read/write.|
|[Width](inspector-width-property-outlook.md)|Returns or sets a  **Long** indicating the width (in pixels) of the specified object. Read/write.|
|[WindowState](inspector-windowstate-property-outlook.md)|Returns or sets the property with a constant in the  **[OlWindowState](olwindowstate-enumeration-outlook.md)** enumeration specifying the window state of an explorer or inspector window. Read/write.|
|[WordEditor](inspector-wordeditor-property-outlook.md)|Returns the Microsoft Word Document Object Model of the message being displayed. Read-only.|

