---
title: Explorer Object (Outlook)
keywords: vbaol11.chm2985
f1_keywords:
- vbaol11.chm2985
ms.prod: OUTLOOK
api_name:
- Outlook.Explorer
ms.assetid: 026591e5-049f-503a-4166-34e6dbc225fb
---


# Explorer Object (Outlook)

Represents the window in which the contents of a folder are displayed.


## Remarks




- Use the  **[Item](http://msdn.microsoft.com/library/explorers-item-method-outlook%28Office.15%29.aspx)** method of the **[Explorers](http://msdn.microsoft.com/library/explorers-object-outlook%28Office.15%29.aspx)** object to return the object representing a specific explorer.
    
- Use the  **[ActiveExplorer](http://msdn.microsoft.com/library/application-activeexplorer-method-outlook%28Office.15%29.aspx)** method to return the object representing the currently active explorer (if there is one).
    
- Use the  **[GetExplorer](http://msdn.microsoft.com/library/folder-getexplorer-method-outlook%28Office.15%29.aspx)** method to return the **Explorer** object associated with a folder.
    
- Use the  **[Display](http://msdn.microsoft.com/library/folder-display-method-outlook%28Office.15%29.aspx)** method of a **[Folder](folder-object-outlook.md)** object to display a folder in its associated explorer.
    

## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/explorer-activate-event-outlook%28Office.15%29.aspx)|
|[AttachmentSelectionChange](http://msdn.microsoft.com/library/explorer-attachmentselectionchange-event-outlook%28Office.15%29.aspx)|
|[BeforeFolderSwitch](http://msdn.microsoft.com/library/explorer-beforefolderswitch-event-outlook%28Office.15%29.aspx)|
|[BeforeItemCopy](http://msdn.microsoft.com/library/explorer-beforeitemcopy-event-outlook%28Office.15%29.aspx)|
|[BeforeItemCut](http://msdn.microsoft.com/library/explorer-beforeitemcut-event-outlook%28Office.15%29.aspx)|
|[BeforeItemPaste](http://msdn.microsoft.com/library/explorer-beforeitempaste-event-outlook%28Office.15%29.aspx)|
|[BeforeMaximize](http://msdn.microsoft.com/library/explorer-beforemaximize-event-outlook%28Office.15%29.aspx)|
|[BeforeMinimize](http://msdn.microsoft.com/library/explorer-beforeminimize-event-outlook%28Office.15%29.aspx)|
|[BeforeMove](http://msdn.microsoft.com/library/explorer-beforemove-event-outlook%28Office.15%29.aspx)|
|[BeforeSize](http://msdn.microsoft.com/library/explorer-beforesize-event-outlook%28Office.15%29.aspx)|
|[BeforeViewSwitch](http://msdn.microsoft.com/library/explorer-beforeviewswitch-event-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/explorer-close-event-outlook%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/explorer-deactivate-event-outlook%28Office.15%29.aspx)|
|[FolderSwitch](http://msdn.microsoft.com/library/explorer-folderswitch-event-outlook%28Office.15%29.aspx)|
|[InlineResponse](http://msdn.microsoft.com/library/explorer-inlineresponse-event-outlook%28Office.15%29.aspx)|
|[InlineResponseClose](http://msdn.microsoft.com/library/explorer-inlineresponseclose-event-outlook%28Office.15%29.aspx)|
|[SelectionChange](http://msdn.microsoft.com/library/explorer-selectionchange-event-outlook%28Office.15%29.aspx)|
|[ViewSwitch](http://msdn.microsoft.com/library/explorer-viewswitch-event-outlook%28Office.15%29.aspx)|
|[DisplayModeChange](http://msdn.microsoft.com/library/explorer-displaymodechange-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/explorer-activate-method-outlook%28Office.15%29.aspx)|
|[AddToSelection](http://msdn.microsoft.com/library/explorer-addtoselection-method-outlook%28Office.15%29.aspx)|
|[ClearSearch](http://msdn.microsoft.com/library/explorer-clearsearch-method-outlook%28Office.15%29.aspx)|
|[ClearSelection](http://msdn.microsoft.com/library/explorer-clearselection-method-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/explorer-close-method-outlook%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/explorer-display-method-outlook%28Office.15%29.aspx)|
|[IsItemSelectableInView](http://msdn.microsoft.com/library/explorer-isitemselectableinview-method-outlook%28Office.15%29.aspx)|
|[IsPaneVisible](http://msdn.microsoft.com/library/explorer-ispanevisible-method-outlook%28Office.15%29.aspx)|
|[RemoveFromSelection](http://msdn.microsoft.com/library/explorer-removefromselection-method-outlook%28Office.15%29.aspx)|
|[Search](http://msdn.microsoft.com/library/explorer-search-method-outlook%28Office.15%29.aspx)|
|[SelectAllItems](http://msdn.microsoft.com/library/explorer-selectallitems-method-outlook%28Office.15%29.aspx)|
|[ShowPane](http://msdn.microsoft.com/library/explorer-showpane-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AccountSelector](http://msdn.microsoft.com/library/explorer-accountselector-property-outlook%28Office.15%29.aspx)|
|[ActiveInlineResponse](http://msdn.microsoft.com/library/explorer-activeinlineresponse-property-outlook%28Office.15%29.aspx)|
|[ActiveInlineResponseWordEditor](http://msdn.microsoft.com/library/explorer-activeinlineresponsewordeditor-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/explorer-application-property-outlook%28Office.15%29.aspx)|
|[AttachmentSelection](http://msdn.microsoft.com/library/explorer-attachmentselection-property-outlook%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/explorer-caption-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/explorer-class-property-outlook%28Office.15%29.aspx)|
|[CurrentFolder](http://msdn.microsoft.com/library/explorer-currentfolder-property-outlook%28Office.15%29.aspx)|
|[CurrentView](http://msdn.microsoft.com/library/explorer-currentview-property-outlook%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/explorer-height-property-outlook%28Office.15%29.aspx)|
|[HTMLDocument](http://msdn.microsoft.com/library/explorer-htmldocument-property-outlook%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/explorer-left-property-outlook%28Office.15%29.aspx)|
|[NavigationPane](http://msdn.microsoft.com/library/explorer-navigationpane-property-outlook%28Office.15%29.aspx)|
|[Panes](http://msdn.microsoft.com/library/explorer-panes-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/explorer-parent-property-outlook%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/explorer-selection-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/explorer-session-property-outlook%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/explorer-top-property-outlook%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/explorer-width-property-outlook%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/explorer-windowstate-property-outlook%28Office.15%29.aspx)|
|[DisplayMode](http://msdn.microsoft.com/library/explorer-displaymode-property-outlook%28Office.15%29.aspx)|
|[PreviewPane](http://msdn.microsoft.com/library/explorer-previewpane-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Explorer Object Members](http://msdn.microsoft.com/library/explorer-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
