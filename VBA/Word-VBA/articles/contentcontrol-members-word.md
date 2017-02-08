---
title: ContentControl Members (Word)
ms.prod: WORD
ms.assetid: d5aa195c-8d7a-0bad-09fa-6f1bfc9828cc
---


# ContentControl Members (Word)
An individual content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. The  **ContentControl** object is a member of the **[ContentControls](contentcontrols-object-word.md)** collection.

An individual content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. The  **ContentControl** object is a member of the **[ContentControls](contentcontrols-object-word.md)** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Copy](contentcontrol-copy-method-word.md)|Copies the content control from the active document to the Clipboard.|
|[Cut](contentcontrol-cut-method-word.md)|Removes the content control from the active document and moves the content control to the Clipboard.|
|[Delete](contentcontrol-delete-method-word.md)|Deletes the specified content control and the contents of the content control.|
|[SetCheckedSymbol](contentcontrol-setcheckedsymbol-method-word.md)|Sets the symbol used to represent the checked state of a check box content control.|
|[SetPlaceholderText](contentcontrol-setplaceholdertext-method-word.md)|Sets the placeholder text that displays in the content control until a user enters their own text.|
|[SetUncheckedSymbol](contentcontrol-setuncheckedsymbol-method-word.md)|Sets the symbol used to represent the unchecked state of a check box content control.|
|[Ungroup](contentcontrol-ungroup-method-word.md)|Removes a group content control from a document so that its child content controls are no longer nested and can be freely edited.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowInsertDeleteSection](contentcontrol-allowinsertdeletesection-property-word.md)|Gets or sets whether users can add or remove sections from the specified repeating section content control by using the user interface.|
|[Appearance](contentcontrol-appearance-property-word.md)|Returns or sets the appearance of the content control. Read/write [WdContentControlAppearance](wdcontentcontrolappearance-enumeration-word.md).|
|[Application](contentcontrol-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[BuildingBlockCategory](contentcontrol-buildingblockcategory-property-word.md)|Returns or sets a  **String** that represents the category for a building block content control. Read/write.|
|[BuildingBlockType](contentcontrol-buildingblocktype-property-word.md)|Returns or sets a  **WdBuildingBlockTypes** constant that represents they type of building block for a building block content control. Read/write.|
|[Checked](contentcontrol-checked-property-word.md)|Returns or sets a  **Boolean** that represents the current state (checked/unchecked) for a check box. Read/Write.|
|[Color](contentcontrol-color-property-word.md)|Returns or sets the color of the content control. Read/write [WdColor](contentcontrol-color-property-word.md).|
|[Creator](contentcontrol-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .|
|[DateCalendarType](contentcontrol-datecalendartype-property-word.md)|Returns or sets a  **[WdCalendarType](wdcalendartype-enumeration-word.md)** constant that represents the calendar type for a calendar content control. Read/write.|
|[DateDisplayFormat](contentcontrol-datedisplayformat-property-word.md)|Returns or sets a  **String** that represents the format in which dates are displayed. Read/write.|
|[DateDisplayLocale](contentcontrol-datedisplaylocale-property-word.md)|Returns a  **WdLanguageID** that represents the language format for the date displayed in a date content control. Read/write.|
|[DateStorageFormat](contentcontrol-datestorageformat-property-word.md)|Returns or sets a  **[WdContentControlDateStorageFormat](wdcontentcontroldatestorageformat-enumeration-word.md)** that represents the format for storage and retrieval of dates when a date content control is bound to the XML data store of the active document. Read/write.|
|[DefaultTextStyle](contentcontrol-defaulttextstyle-property-word.md)|Returns or sets a  **Variant** that represents the name of the character style to use to format text in a text content control. Read/write.|
|[DropdownListEntries](contentcontrol-dropdownlistentries-property-word.md)|Returns a  **[ContentControlListEntries](contentcontrollistentries-object-word.md)** collection that represents the items in a drop-down list content control or in a combo box content control. Read-only.|
|[ID](contentcontrol-id-property-word.md)|Returns a  **String** that represents the identification for a content control. Read-only.|
|[Level](contentcontrol-level-property-word.md)|Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline. Read-only [WdContentControlLevel](wdcontentcontrollevel-enumeration-word.md).|
|[LockContentControl](contentcontrol-lockcontentcontrol-property-word.md)|Returns or sets a  **Boolean** that represents whether the user can delete a content control from the active document. Read/write.|
|[LockContents](contentcontrol-lockcontents-property-word.md)|Returns or sets a  **Boolean** that represents whether the user can edit the contents of a content control. Read/write.|
|[MultiLine](contentcontrol-multiline-property-word.md)|Returns a  **Boolean** that represents whether a text content control allows multiple lines of text. Read/write.|
|[Parent](contentcontrol-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **ContentControl** object.|
|[ParentContentControl](contentcontrol-parentcontentcontrol-property-word.md)|Returns a  **ContentControl** that represents the parent content control for a content control that is nested inside a rich-text control or group control. Read-only.|
|[PlaceholderText](contentcontrol-placeholdertext-property-word.md)|Returns a  **BuildingBlock** object that represents the placeholder text for a content control. Read-only.|
|[Range](contentcontrol-range-property-word.md)|Returns a  **[Range](range-object-word.md)** that represents the contents of the content control in the active document. Read-only.|
|[RepeatingSectionItems](contentcontrol-repeatingsectionitems-property-word.md)|Returns the collection of repeating section items in the specified repeating section content control. Read-only.|
|[RepeatingSectionItemTitle](contentcontrol-repeatingsectionitemtitle-property-word.md)|Returns or sets the name of the repeating section items used in the context menu associated with the specified repeating section content control. Read/write.|
|[ShowingPlaceholderText](contentcontrol-showingplaceholdertext-property-word.md)|Returns a  **Boolean** that indicates whether the placeholder text for the content control is displayed. Read-only.|
|[Tag](contentcontrol-tag-property-word.md)|Returns or sets a  **String** that represents a value to identify a content control. Read/write.|
|[Temporary](contentcontrol-temporary-property-word.md)|Returns or sets a  **Boolean** that represents whether to remove a content control from the active document when the user edits the contents of the control. Read/write.|
|[Title](contentcontrol-title-property-word.md)|Returns or sets a  **String** that represents the title for a content control. Read/write.|
|[Type](contentcontrol-type-property-word.md)|Returns or sets a  **[WdContentControlType](wdcontentcontroltype-enumeration-word.md)** that represents the type for a content control. Read/write.|
|[XMLMapping](contentcontrol-xmlmapping-property-word.md)|Returns an  **[XMLMapping](xmlmapping-object-word.md)** object that represents the mapping of a content control to XML data in the data store of a document. Read-only.|

