---
title: Hyperlink Members (Word)
ms.prod: WORD
ms.assetid: 49699791-6b9c-2061-aff7-c9269747ecea
---


# Hyperlink Members (Word)
Represents a hyperlink. The  **Hyperlink** object is a member of the **Hyperlinks** collection.

Represents a hyperlink. The  **Hyperlink** object is a member of the **Hyperlinks** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddToFavorites](hyperlink-addtofavorites-method-word.md)|Creates a shortcut to the document or hyperlink and adds it to the Favorites folder.|
|[CreateNewDocument](hyperlink-createnewdocument-method-word.md)|Creates a new document linked to the specified hyperlink.|
|[Delete](hyperlink-delete-method-word.md)|Deletes the specified hyperlink.|
|[Follow](hyperlink-follow-method-word.md)|Displays a cached document associated with the specified  **Hyperlink** object, if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Address](hyperlink-address-property-word.md)|Returns or sets the address (for example, a file name or URL) of the specified hyperlink. Read/write  **String** .|
|[Application](hyperlink-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](hyperlink-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[EmailSubject](hyperlink-emailsubject-property-word.md)|Returns or sets the text string for the specified hyperlink's subject line. Read/write  **String** .|
|[ExtraInfoRequired](hyperlink-extrainforequired-property-word.md)| **True** if extra information is required to resolve the specified hyperlink. Read-only **Boolean** .|
|[Name](hyperlink-name-property-word.md)|Returns the name of the specified object. Read-only  **String** .|
|[Parent](hyperlink-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Hyperlink** object.|
|[Range](hyperlink-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that is contained within a hyperlink.|
|[ScreenTip](hyperlink-screentip-property-word.md)|Returns or sets the text that appears as a ScreenTip when the mouse pointer is positioned over the specified hyperlink. Read/write  **String** .|
|[Shape](hyperlink-shape-property-word.md)|Returns a  **Shape** object for the specified hyperlink or diagram node.|
|[SubAddress](hyperlink-subaddress-property-word.md)|Returns or sets a named location in the destination of the specified hyperlink. Read/write  **String** .|
|[Target](hyperlink-target-property-word.md)|Returns or sets the name of the frame or window in which to load the hyperlink. Read/write  **String** .|
|[TextToDisplay](hyperlink-texttodisplay-property-word.md)|Returns or sets the specified hyperlink's visible text in a document. Read/write  **String** .|
|[Type](hyperlink-type-property-word.md)|Returns the hyperlink type. Read-only  **MsoHyperlinkType** .|

