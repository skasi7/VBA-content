---
title: TableView Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 2cc17ec6-12cf-d335-9370-d3922b45510e
---


# TableView Members (Outlook)
Represents a view that displays Outlook items in a table, with each item in a row and the details of the item in the columns.

Represents a view that displays Outlook items in a table, with each item in a row and the details of the item in the columns.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Apply](tableview-apply-method-outlook.md)|Applies the  **[TableView](tableview-object-outlook.md)** object to the current view.|
|[Copy](tableview-copy-method-outlook.md)|Creates a new  **[View](view-object-outlook.md)** object based on the existing **[TableView](tableview-object-outlook.md)** object.|
|[Delete](tableview-delete-method-outlook.md)|Deletes an object from a collection.|
|[GetTable](tableview-gettable-method-outlook.md)|Returns a  **[Table](table-object-outlook.md)** object that represents all of the Microsoft Outlook items that are contained in a **[TableView](tableview-object-outlook.md)** object.|
|[GoToDate](tableview-gotodate-method-outlook.md)|Changes the date used by the current view to display information.|
|[Reset](tableview-reset-method-outlook.md)|Resets a built-in Microsoft Outlook view to its original settings.|
|[Save](tableview-save-method-outlook.md)|Saves the view, or saves the changes to a view.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowInCellEditing](tableview-allowincellediting-property-outlook.md)|Returns or sets a  **Boolean** value that determines whether in-cell editing is allowed in the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[AlwaysExpandConversation](tableview-alwaysexpandconversation-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether conversations are always fully expanded in the table view. Read/write.|
|[Application](tableview-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[AutoFormatRules](tableview-autoformatrules-property-outlook.md)|Returns an  **[AutoFormatRules](autoformatrules-object-outlook.md)** object that represents the set of formatting rules applicable to the **[TableView](tableview-object-outlook.md)** object. Read-only.|
|[AutomaticColumnSizing](tableview-automaticcolumnsizing-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether the columns in the **[TableView](tableview-object-outlook.md)** object are automatically sized by Outlook. Read/write.|
|[AutomaticGrouping](tableview-automaticgrouping-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether the automatic grouping is active in the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[AutoPreview](tableview-autopreview-property-outlook.md)|Returns or sets an  **[OlAutoPreview](olautopreview-enumeration-outlook.md)** constant that determines how items are automatically previewed by the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[AutoPreviewFont](tableview-autopreviewfont-property-outlook.md)|Returns a  **[ViewFont](viewfont-object-outlook.md)** object that represents the font used when automatically previewing Outlook items in the **[TableView](tableview-object-outlook.md)** object. Read-only.|
|[Class](tableview-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[ColumnFont](tableview-columnfont-property-outlook.md)|Returns a  **[ViewFont](viewfont-object-outlook.md)** object that represents the font used when displaying column headers in the **[TableView](tableview-object-outlook.md)** object. Read-only.|
|[DefaultExpandCollapseSetting](tableview-defaultexpandcollapsesetting-property-outlook.md)|Returns or sets an  **[OlDefaultExpandCollapseSetting](oldefaultexpandcollapsesetting-enumeration-outlook.md)** constant that determines the default expansion setting for groups in the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[Filter](tableview-filter-property-outlook.md)|Returns or sets a  **String** value that represents the filter for a view. Read/write.|
|[GridLineStyle](tableview-gridlinestyle-property-outlook.md)|Returns or sets an  **[OlGridLineStyle](olgridlinestyle-enumeration-outlook.md)** constant that represents the line style used for grid lines in the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[GroupByFields](tableview-groupbyfields-property-outlook.md)|Returns an  **[OrderFields](orderfields-object-outlook.md)** object that represents the set of fields by which the items displayed in the **[TableView](tableview-object-outlook.md)** object are grouped. Read-only.|
|[HideReadingPaneHeaderInfo](tableview-hidereadingpaneheaderinfo-property-outlook.md)|Returns or sets a  **Boolean** value that determines whether the header for an Outlook item is displayed in the Reading Pane for the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[Language](tableview-language-property-outlook.md)|Returns or sets a  **String** value that represents the language setting for the view. Read/write.|
|[LockUserChanges](tableview-lockuserchanges-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether a user can modify the settings of the view. Read/write.|
|[MaxLinesInMultiLineView](tableview-maxlinesinmultilineview-property-outlook.md)|Returns or sets a  **Long** value that determines the maximum number of lines displayed in multiline mode for the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[MultiLine](tableview-multiline-property-outlook.md)|Returns or sets an  **[OlMultiLine](olmultiline-enumeration-outlook.md)** constant that determines how multiple lines are displayed in the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[MultiLineWidth](tableview-multilinewidth-property-outlook.md)|Returns or sets a  **Long** value that represents the text width (in characters) needed to trigger multiline mode in the **[TableView](tableview-object-outlook.md)** object . Read/write|
|[Name](tableview-name-property-outlook.md)|Returns or sets a  **String** value that represents the display name for the object. Read/write.|
|[Parent](tableview-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[RowFont](tableview-rowfont-property-outlook.md)|Returns a  **[ViewFont](viewfont-object-outlook.md)** object that represents the font used when displaying rows in the **[TableView](tableview-object-outlook.md)** object. Read-only.|
|[SaveOption](tableview-saveoption-property-outlook.md)|Returns an  **[OlViewSaveOption](olviewsaveoption-enumeration-outlook.md)** constant that specifies the folders in which the specified view is available and the read permissions attached to the view. Read-only.|
|[Session](tableview-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[ShowConversationByDate](tableview-showconversationbydate-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether items in a conversation are organized vertically left-aligned and ordered by the received date and time, with the most recent item on top. Read/write.|
|[ShowConversationSendersAboveSubject](tableview-showconversationsendersabovesubject-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether the table view displays the participating senders above the subject line in the conversation header, or below it. Read/write.|
|[ShowFullConversations](tableview-showfullconversations-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether to display conversation items from other folders, such as the Sent Items folder, as part of the conversation in the table view. Read/write.|
|[ShowItemsInGroups](tableview-showitemsingroups-property-outlook.md)|Returns or sets a  **Boolean** value that determines whether Outlook items are shown in groups within the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[ShowNewItemRow](tableview-shownewitemrow-property-outlook.md)|Returns or sets a  **Boolean** value that determines if the new item row is displayed in the **[TableView](tableview-object-outlook.md)** object. Read/write|
|[ShowReadingPane](tableview-showreadingpane-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether the Reading Pane is displayed in the **[TableView](tableview-object-outlook.md)** object. Read/write.|
|[SortFields](tableview-sortfields-property-outlook.md)|Returns an  **[OrderFields](orderfields-object-outlook.md)** object that represents the set of fields by which the items displayed in the **[TableView](tableview-object-outlook.md)** object are ordered. Read-only.|
|[Standard](tableview-standard-property-outlook.md)|Returns a  **Boolean** value that indicates whether the **[TableView](tableview-object-outlook.md)** object is a built-in Outlook view. Read-only.|
|[ViewFields](tableview-viewfields-property-outlook.md)|Returns a  **[ViewFields](viewfields-object-outlook.md)** object that represents the set of fields with which Outlook items are displayed in the **[TableView](tableview-object-outlook.md)** object. Read-only.|
|[ViewType](tableview-viewtype-property-outlook.md)|Returns an  **[OlViewType](olviewtype-enumeration-outlook.md)** constant that indicates the view type of the view. Read-only.|
|[XML](tableview-xml-property-outlook.md)|Returns or sets a  **String** value that specifies the XML definition of the view. Read/write.|

