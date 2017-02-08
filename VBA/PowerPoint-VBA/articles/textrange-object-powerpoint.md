---
title: TextRange Object (PowerPoint)
keywords: vbapp10.chm569000
f1_keywords:
- vbapp10.chm569000
ms.prod: POWERPOINT
ms.assetid: 7c234107-c423-7ec9-e8bd-a82cc3b345de
---


# TextRange Object (PowerPoint)

Contains the text that's attached to a shape, and properties and methods for manipulating the text.


## Remarks

The following examples describe how to:


- Return the text range in any shape you specify.
    
- Return a text range from the selection.
    
- Return particular characters, words, lines, sentences, or paragraphs from a text range.
    
- Find and replace text in a text range.
    
- Insert text, the date and time, or the slide number into a text range.
    
- Position the cursor wherever you want in a text range.
    

## Example

Use the [TextRange](http://msdn.microsoft.com/library/textframe-textrange-property-powerpoint%28Office.15%29.aspx)property of the  **[TextFrame](textframe-object-powerpoint.md)** object to return a **TextRange** object for any shape you specify. Use the[Text](http://msdn.microsoft.com/library/textrange-text-property-powerpoint%28Office.15%29.aspx)property to return the string of text in the  **TextRange** object. The following example adds a rectangle to `myDocument` and sets the text it contains.


```
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140) _

    .TextFrame.TextRange.Text = "Here is some test text"
```

Because the  **Text** property is the default property of the **TextRange** object, the following two statements are equivalent.




```
ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange.Text = "Here is some test text"

ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange = "Here is some test text"
```

Use the [HasTextFrame](http://msdn.microsoft.com/library/shape-hastextframe-property-powerpoint%28Office.15%29.aspx)property to determine whether a shape has a text frame, and use the [HasText](http://msdn.microsoft.com/library/textframe-hastext-property-powerpoint%28Office.15%29.aspx)property to determine whether the text frame contains text.

Use the  **TextRange** property of the **Selection** object to return the currently selected text. The following example copies the selection to the Clipboard.




```
ActiveWindow.Selection.TextRange.Copy
```

Use one of the following methods to return a portion of the text of a  **TextRange** object: **[Characters](http://msdn.microsoft.com/library/textrange-characters-method-powerpoint%28Office.15%29.aspx)**, **[Lines](http://msdn.microsoft.com/library/textrange-lines-method-powerpoint%28Office.15%29.aspx)**, **[Paragraphs](http://msdn.microsoft.com/library/textrange-paragraphs-method-powerpoint%28Office.15%29.aspx)**, **[Runs](http://msdn.microsoft.com/library/textrange-runs-method-powerpoint%28Office.15%29.aspx)**, **[Sentences](http://msdn.microsoft.com/library/textrange-sentences-method-powerpoint%28Office.15%29.aspx)**, or **[Words](http://msdn.microsoft.com/library/textrange-words-method-powerpoint%28Office.15%29.aspx)**.

Use the [Find](http://msdn.microsoft.com/library/textrange-find-method-powerpoint%28Office.15%29.aspx)and [Replace](http://msdn.microsoft.com/library/textrange-replace-method-powerpoint%28Office.15%29.aspx)methods to find and replace text in a text range.

Use one of the following methods to insert characters into a  **TextRange** object:[InsertAfter](http://msdn.microsoft.com/library/textrange-insertafter-method-powerpoint%28Office.15%29.aspx), [InsertBefore](http://msdn.microsoft.com/library/textrange-insertbefore-method-powerpoint%28Office.15%29.aspx), [InsertDateTime](http://msdn.microsoft.com/library/textrange-insertdatetime-method-powerpoint%28Office.15%29.aspx), [InsertSlideNumber](http://msdn.microsoft.com/library/textrange-insertslidenumber-method-powerpoint%28Office.15%29.aspx), or [InsertSymbol](http://msdn.microsoft.com/library/textrange-insertsymbol-method-powerpoint%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[AddPeriods](http://msdn.microsoft.com/library/textrange-addperiods-method-powerpoint%28Office.15%29.aspx)|
|[ChangeCase](http://msdn.microsoft.com/library/textrange-changecase-method-powerpoint%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/textrange-characters-method-powerpoint%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/textrange-copy-method-powerpoint%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/textrange-cut-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/textrange-delete-method-powerpoint%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/textrange-find-method-powerpoint%28Office.15%29.aspx)|
|[InsertAfter](http://msdn.microsoft.com/library/textrange-insertafter-method-powerpoint%28Office.15%29.aspx)|
|[InsertBefore](http://msdn.microsoft.com/library/textrange-insertbefore-method-powerpoint%28Office.15%29.aspx)|
|[InsertDateTime](http://msdn.microsoft.com/library/textrange-insertdatetime-method-powerpoint%28Office.15%29.aspx)|
|[InsertSlideNumber](http://msdn.microsoft.com/library/textrange-insertslidenumber-method-powerpoint%28Office.15%29.aspx)|
|[InsertSymbol](http://msdn.microsoft.com/library/textrange-insertsymbol-method-powerpoint%28Office.15%29.aspx)|
|[Lines](http://msdn.microsoft.com/library/textrange-lines-method-powerpoint%28Office.15%29.aspx)|
|[LtrRun](http://msdn.microsoft.com/library/textrange-ltrrun-method-powerpoint%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/textrange-paragraphs-method-powerpoint%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/textrange-paste-method-powerpoint%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/textrange-pastespecial-method-powerpoint%28Office.15%29.aspx)|
|[RemovePeriods](http://msdn.microsoft.com/library/textrange-removeperiods-method-powerpoint%28Office.15%29.aspx)|
|[Replace](http://msdn.microsoft.com/library/textrange-replace-method-powerpoint%28Office.15%29.aspx)|
|[RotatedBounds](http://msdn.microsoft.com/library/textrange-rotatedbounds-method-powerpoint%28Office.15%29.aspx)|
|[RtlRun](http://msdn.microsoft.com/library/textrange-rtlrun-method-powerpoint%28Office.15%29.aspx)|
|[Runs](http://msdn.microsoft.com/library/textrange-runs-method-powerpoint%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/textrange-select-method-powerpoint%28Office.15%29.aspx)|
|[Sentences](http://msdn.microsoft.com/library/textrange-sentences-method-powerpoint%28Office.15%29.aspx)|
|[TrimText](http://msdn.microsoft.com/library/textrange-trimtext-method-powerpoint%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/textrange-words-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActionSettings](http://msdn.microsoft.com/library/textrange-actionsettings-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/textrange-application-property-powerpoint%28Office.15%29.aspx)|
|[BoundHeight](http://msdn.microsoft.com/library/textrange-boundheight-property-powerpoint%28Office.15%29.aspx)|
|[BoundLeft](http://msdn.microsoft.com/library/textrange-boundleft-property-powerpoint%28Office.15%29.aspx)|
|[BoundTop](http://msdn.microsoft.com/library/textrange-boundtop-property-powerpoint%28Office.15%29.aspx)|
|[BoundWidth](http://msdn.microsoft.com/library/textrange-boundwidth-property-powerpoint%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/textrange-count-property-powerpoint%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/textrange-font-property-powerpoint%28Office.15%29.aspx)|
|[IndentLevel](http://msdn.microsoft.com/library/textrange-indentlevel-property-powerpoint%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/textrange-languageid-property-powerpoint%28Office.15%29.aspx)|
|[Length](http://msdn.microsoft.com/library/textrange-length-property-powerpoint%28Office.15%29.aspx)|
|[ParagraphFormat](http://msdn.microsoft.com/library/textrange-paragraphformat-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/textrange-parent-property-powerpoint%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/textrange-start-property-powerpoint%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/textrange-text-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
