---
title: TextRange Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: cb8dc5ff-34de-3d04-1d56-ed387daaf6b9
---


# TextRange Members (PowerPoint)
Contains the text that's attached to a shape, and properties and methods for manipulating the text.

Contains the text that's attached to a shape, and properties and methods for manipulating the text.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddPeriods](textrange-addperiods-method-powerpoint.md)|Adds a period at the end of each paragraph in the specified text.|
|[ChangeCase](textrange-changecase-method-powerpoint.md)|Changes the case of the specified text.|
|[Characters](textrange-characters-method-powerpoint.md)|Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents the specified subset of text characters. For information about counting or looping through the characters in a text range, see the **[TextRange](textrange-object-powerpoint.md)** object.|
|[Copy](textrange-copy-method-powerpoint.md)|Copies the specified object to the Clipboard.|
|[Cut](textrange-cut-method-powerpoint.md)|Deletes the specified object and places it on the Clipboard.|
|[Delete](textrange-delete-method-powerpoint.md)|Deletes the specified  **TextRange** object.|
|[Find](textrange-find-method-powerpoint.md)|Finds the specified text in a text range, and returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents the first text range where the text is found. Returns **Nothing** if no match is found.|
|[InsertAfter](textrange-insertafter-method-powerpoint.md)|Appends a string to the end of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.|
|[InsertBefore](textrange-insertbefore-method-powerpoint.md)|Appends a string to the beginning of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.|
|[InsertDateTime](textrange-insertdatetime-method-powerpoint.md)|Inserts the date and time in the specified text range. Returns a  **TextRange** object that represents the inserted text.|
|[InsertSlideNumber](textrange-insertslidenumber-method-powerpoint.md)|Inserts the slide number of the current slide into the specified text range. Returns a  **TextRange** object that represents the slide number.|
|[InsertSymbol](textrange-insertsymbol-method-powerpoint.md)|Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents a symbol inserted into the specified text range.|
|[Lines](textrange-lines-method-powerpoint.md)|Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents the specified subset of text lines. For information about counting or looping through the lines in a text range, see the **[TextRange](textrange-object-powerpoint.md)** object.|
|[LtrRun](textrange-ltrrun-method-powerpoint.md)|Sets the direction of text in a text range to read from left to right.|
|[Paragraphs](textrange-paragraphs-method-powerpoint.md)|Returns a  **TextRange** object that represents the specified subset of text paragraphs.|
|[Paste](textrange-paste-method-powerpoint.md)|Pastes the text on the Clipboard into the specified text range, and returns a  **TextRange** object that represents the pasted text.|
|[PasteSpecial](textrange-pastespecial-method-powerpoint.md)|Replaces the text range with the contents of the Clipboard in the format specified. |
|[RemovePeriods](textrange-removeperiods-method-powerpoint.md)|Removes the period at the end of each paragraph in the specified text.|
|[Replace](textrange-replace-method-powerpoint.md)|Finds specific text in a text range, replaces the found text with a specified string, and returns a  **TextRange** object that represents the first occurrence of the found text. Returns **Nothing** if no match is found.|
|[RotatedBounds](textrange-rotatedbounds-method-powerpoint.md)|Returns the coordinates of the vertices of the text bounding box for the specified text range.|
|[RtlRun](textrange-rtlrun-method-powerpoint.md)|Sets the direction of text in a text range to read from right to left.|
|[Runs](textrange-runs-method-powerpoint.md)|Returns a  **TextRange** object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes.|
|[Select](textrange-select-method-powerpoint.md)|Selects the specified object.|
|[Sentences](textrange-sentences-method-powerpoint.md)|Returns a  **TextRange** object that represents the specified subset of text sentences.|
|[TrimText](textrange-trimtext-method-powerpoint.md)|Returns a  **TextRange** object that represents the specified text minus any trailing spaces.|
|[Words](textrange-words-method-powerpoint.md)|Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents the specified subset of text words.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActionSettings](textrange-actionsettings-property-powerpoint.md)|Returns an  **[ActionSettings](actionsettings-object-powerpoint.md)** object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.|
|[Application](textrange-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[BoundHeight](textrange-boundheight-property-powerpoint.md)|Returns the height (in points) of the text bounding box for the specified text frame. Read-only.|
|[BoundLeft](textrange-boundleft-property-powerpoint.md)|Returns the distance (in points) from the left edge of the text bounding box for the specified text frame to the left edge of the slide. Read-only.|
|[BoundTop](textrange-boundtop-property-powerpoint.md)|Returns the distance (in points) from the top of the of the text bounding box for the specified text frame to the top of the slide. Read-only.|
|[BoundWidth](textrange-boundwidth-property-powerpoint.md)|Returns the width (in points) of the text bounding box for the specified text frame. Read-only.|
|[Count](textrange-count-property-powerpoint.md)|Returns the number of objects in the specified collection. Read-only.|
|[Font](textrange-font-property-powerpoint.md)|Returns a  **[Font](font-object-powerpoint.md)** object that represents character formatting. Read-only.|
|[IndentLevel](textrange-indentlevel-property-powerpoint.md)|Returns or sets the the indent level for the specified text as an integer from 1 to 5, where 1 indicates a first-level paragraph with no indentation. Read/write.|
|[LanguageID](textrange-languageid-property-powerpoint.md)|Returns or sets the language for the specified text range. Read/write.|
|[Length](textrange-length-property-powerpoint.md)|Returns the length of the specified text range, in characters. Read-only.|
|[ParagraphFormat](textrange-paragraphformat-property-powerpoint.md)|Returns a  **[ParagraphFormat](paragraphformat-object-powerpoint.md)** object that represents paragraph formatting for the specified text. Read-only.|
|[Parent](textrange-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[Start](textrange-start-property-powerpoint.md)|Returns the position of the first character in the specified text range relative to the first character in the shape that contains the text. Read-only.|
|[Text](textrange-text-property-powerpoint.md)|Returns or sets a  **String** that represents the text contained in the specified object. Read/write.|

