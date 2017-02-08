---
title: Explorer Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 4412c507-4dcd-6005-b9c8-11824624250d
---


# Explorer Members (Outlook)
Represents the window in which the contents of a folder are displayed.

Represents the window in which the contents of a folder are displayed.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Activate](explorer-activate-event-outlook.md)|Occurs when an explorer becomes the active window, either as a result of user action or through program code.|
|[AttachmentSelectionChange](explorer-attachmentselectionchange-event-outlook.md)|Occurs when the user selects a different or additional attachment in the active explorer programmatically or by interacting with the user interface.|
|[BeforeFolderSwitch](explorer-beforefolderswitch-event-outlook.md)|Occurs before the explorer goes to a new folder, either as a result of user action or through program code.|
|[BeforeItemCopy](explorer-beforeitemcopy-event-outlook.md)|Occurs when an Outlook item is copied.|
|[BeforeItemCut](explorer-beforeitemcut-event-outlook.md)|Occurs when an Outlook item is cut from a folder.|
|[BeforeItemPaste](explorer-beforeitempaste-event-outlook.md)|Occurs when an Outlook item is pasted.|
|[BeforeMaximize](explorer-beforemaximize-event-outlook.md)|Occurs when an explorer is maximized by the user.|
|[BeforeMinimize](explorer-beforeminimize-event-outlook.md)|Occurs when the active explorer is minimized by the user.|
|[BeforeMove](explorer-beforemove-event-outlook.md)|Occurs when the  **[Explorer](explorer-object-outlook.md)** is moved by the user.|
|[BeforeSize](explorer-beforesize-event-outlook.md)|Occurs when the user sizes the current  **[Explorer](explorer-object-outlook.md)** .|
|[BeforeViewSwitch](explorer-beforeviewswitch-event-outlook.md)|Occurs before the explorer changes to a new view, either as a result of user action or through program code. |
|[Close](explorer-close-event-outlook.md)|Occurs when an explorer is being closed.|
|[Deactivate](explorer-deactivate-event-outlook.md)|Occurs when an explorer stops being the active window, either as a result of user action or through program code.|
|[FolderSwitch](explorer-folderswitch-event-outlook.md)|Occurs when the explorer goes to a new folder, either as a result of user action or through program code. |
|[InlineResponse](explorer-inlineresponse-event-outlook.md)|Occurs when the user performs an action that causes an inline response to appear in the Reading Pane.|
|[InlineResponseClose](explorer-inlineresponseclose-event-outlook.md)|Occurs when the user performs an action that causes the active inline response to close in the Reading Pane.|
|[SelectionChange](explorer-selectionchange-event-outlook.md)|Occurs when the user selects a different or additional Microsoft Outlook item programmatically or by interacting with the user interface.|
|[ViewSwitch](explorer-viewswitch-event-outlook.md)|Occurs when the view in the explorer changes, either as a result of user action or through program code. |
|[DisplayModeChange](explorer-displaymodechange-event-outlook.md)|Occurs when the user performs an action that changes the display mode. Possible modes include Normal, Portrait View, and Portrait Reading Pane.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](explorer-activate-method-outlook.md)|Activates an explorer window by bringing it to the foreground and setting keyboard focus.|
|[AddToSelection](explorer-addtoselection-method-outlook.md)|Adds the specified Microsoft Outlook item to the selection in the active explorer.|
|[ClearSearch](explorer-clearsearch-method-outlook.md)|Clears results from a Microsoft Instant Search in an  **[Explorer](explorer-object-outlook.md)** if results are displayed in the **Explorer** .|
|[ClearSelection](explorer-clearselection-method-outlook.md)|Cancels any selection in the active explorer.|
|[Close](explorer-close-method-outlook.md)|Closes the  **[Explorer](explorer-object-outlook.md)** object.|
|[Display](explorer-display-method-outlook.md)|Displays a new  **[Explorer](explorer-object-outlook.md)** object for the folder.|
|[IsItemSelectableInView](explorer-isitemselectableinview-method-outlook.md)|Returns a value that indicates whether the specified Microsoft Outlook item can be selected in the current view of the active explorer.|
|[IsPaneVisible](explorer-ispanevisible-method-outlook.md)|Returns a  **Boolean** indicating whether a specific explorer pane is visible.|
|[RemoveFromSelection](explorer-removefromselection-method-outlook.md)|Cancels the selection of the specified Microsoft Outlook item in the active explorer.|
|[Search](explorer-search-method-outlook.md)|Performs a Microsoft Instant Search on the current folder displayed in the Explorer using the given  _Query_.|
|[SelectAllItems](explorer-selectallitems-method-outlook.md)|Selects all items that are contained in the current view of the active explorer. |
|[ShowPane](explorer-showpane-method-outlook.md)|Displays or hides a specific pane in the explorer.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AccountSelector](explorer-accountselector-property-outlook.md)|Returns an  **[AccountSelector](accountselector-object-outlook.md)** object that represents the Microsoft Office Backstage view for the **[Explorer](explorer-object-outlook.md)** object. Read-only.|
|[ActiveInlineResponse](explorer-activeinlineresponse-property-outlook.md)|Returns an item object representing the active inline response item in the explorer reading pane. Read-only.|
|[ActiveInlineResponseWordEditor](explorer-activeinlineresponsewordeditor-property-outlook.md)|Returns the Word [Document](document-object-word.md) object of the active inline response that is displayed in the explorer Reading Pane. Read-only.|
|[Application](explorer-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[AttachmentSelection](explorer-attachmentselection-property-outlook.md)|Returns an  **[AttachmentSelection](attachmentselection-object-outlook.md)** object consisting of one or more attachments that are selected in the current view of the explorer. Read-only.|
|[Caption](explorer-caption-property-outlook.md)|Returns a  **String** representing the title. Read-only.|
|[Class](explorer-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[CurrentFolder](explorer-currentfolder-property-outlook.md)|Returns or sets a  **[Folder](folder-object-outlook.md)** object that represents the current folder displayed in the explorer. Read/write.|
|[CurrentView](explorer-currentview-property-outlook.md)|Returns or sets a  **Variant** representing the current view. Read/write.|
|[Height](explorer-height-property-outlook.md)|Returns or sets a  **Long** specifying the height (in pixels) of the explorer window. Read/write.|
|[HTMLDocument](explorer-htmldocument-property-outlook.md)|Returns an  **HTMLDocument** object that specifies the HTML object model associated with the HTML document in the current view (assuming one exists). Read-only.|
|[Left](explorer-left-property-outlook.md)|Returns or sets a  **Long** specifying the position (in pixels) of the left vertical edge of an explorer window from the edge of the screen. Read/write.|
|[NavigationPane](explorer-navigationpane-property-outlook.md)|Returns a  **[NavigationPane](navigationpane-object-outlook.md)** object that represents the Navigation Pane for an **[Explorer](explorer-object-outlook.md)** object. Read-only.|
|[Panes](explorer-panes-property-outlook.md)|Returns a  **[Panes](panes-object-outlook.md)** collection object representing the panes displayed by the specified explorer.|
|[Parent](explorer-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Selection](explorer-selection-property-outlook.md)|Returns a  **[Selection](selection-object-outlook.md)** object that contains the item or items that are selected in the explorer window. Read-only.|
|[Session](explorer-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Top](explorer-top-property-outlook.md)|Returns or sets a  **Long** indicating the position (in pixels) of the top horizontal edge of an explorer window from the edge of the screen. Read/write.|
|[Width](explorer-width-property-outlook.md)|Returns or sets a  **Long** indicating the width (in pixels) of the specified object. Read/write.|
|[WindowState](explorer-windowstate-property-outlook.md)|Returns or sets the property with a constant in the  **[OlWindowState](olwindowstate-enumeration-outlook.md)** enumeration specifying the window state of an explorer or inspector window. Read/write.|
|[DisplayMode](explorer-displaymode-property-outlook.md)|Indicates the display mode: Normal, Portrait View, or Portrait Reading Pane.|
|[PreviewPane](explorer-previewpane-property-outlook.md)|The [PreviewPane](previewpane-object-outlook.md) object displays content in a "single pane mode" by showing only the Preview Pane view.|

