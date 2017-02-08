---
title: Font Object (PowerPoint)
keywords: vbapp10.chm575000
f1_keywords:
- vbapp10.chm575000
ms.prod: POWERPOINT
ms.assetid: ad62daaa-01a5-36cc-5451-e0da0134ac95
---


# Font Object (PowerPoint)

Represents character formatting for text or a bullet. The  **Font** object is a member of the **[Fonts](http://msdn.microsoft.com/library/fonts-object-powerpoint%28Office.15%29.aspx)** collection. The **Fonts** collection contains all the fonts used in a presentation.


## Example

The following examples describes how to do the following:


- Return the  **Font** object that represents the font attributes of a specified bullet, a specified range of text, or all text at a specified outline level
    
- Return a  **Font** object from the collection of all the fonts used in the presentation
    
Use the [Font](http://msdn.microsoft.com/library/textrange-font-property-powerpoint%28Office.15%29.aspx)property to return the  **Font** object that represents the font attributes for a specific bullet, text range, or outline level. The following example sets the title text on slide one and sets the font properties.




```
With ActivePresentation.Slides(1).Shapes.Title _

        .TextFrame.TextRange

    .Text = "Volcano Coffee"

    With .Font

        .Italic = True

        .Name = "Palatino"

        .Color.RGB = RGB(0, 0, 255)

    End With

End With
```

Use  **Fonts** (index), where index is the font's name or index number, to return a single **Font** object. The following example checks to see whether font one in the active presentation is embedded in the presentation.




```
If ActivePresentation.Fonts(1).Embedded = _

    True Then MsgBox "Font 1 is embedded"
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/font-application-property-powerpoint%28Office.15%29.aspx)|
|[AutoRotateNumbers](http://msdn.microsoft.com/library/font-autorotatenumbers-property-powerpoint%28Office.15%29.aspx)|
|[BaselineOffset](http://msdn.microsoft.com/library/font-baselineoffset-property-powerpoint%28Office.15%29.aspx)|
|[Bold](http://msdn.microsoft.com/library/font-bold-property-powerpoint%28Office.15%29.aspx)|
|[Color](http://msdn.microsoft.com/library/font-color-property-powerpoint%28Office.15%29.aspx)|
|[Embeddable](http://msdn.microsoft.com/library/font-embeddable-property-powerpoint%28Office.15%29.aspx)|
|[Embedded](http://msdn.microsoft.com/library/font-embedded-property-powerpoint%28Office.15%29.aspx)|
|[Emboss](http://msdn.microsoft.com/library/font-emboss-property-powerpoint%28Office.15%29.aspx)|
|[Italic](http://msdn.microsoft.com/library/font-italic-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/font-name-property-powerpoint%28Office.15%29.aspx)|
|[NameAscii](http://msdn.microsoft.com/library/font-nameascii-property-powerpoint%28Office.15%29.aspx)|
|[NameComplexScript](http://msdn.microsoft.com/library/font-namecomplexscript-property-powerpoint%28Office.15%29.aspx)|
|[NameFarEast](http://msdn.microsoft.com/library/font-namefareast-property-powerpoint%28Office.15%29.aspx)|
|[NameOther](http://msdn.microsoft.com/library/font-nameother-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/font-parent-property-powerpoint%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/font-shadow-property-powerpoint%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/font-size-property-powerpoint%28Office.15%29.aspx)|
|[Subscript](http://msdn.microsoft.com/library/font-subscript-property-powerpoint%28Office.15%29.aspx)|
|[Superscript](http://msdn.microsoft.com/library/font-superscript-property-powerpoint%28Office.15%29.aspx)|
|[Underline](http://msdn.microsoft.com/library/font-underline-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
