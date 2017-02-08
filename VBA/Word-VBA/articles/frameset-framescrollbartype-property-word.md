---
title: Frameset.FrameScrollbarType Property (Word)
keywords: vbawd10.chm165806110
f1_keywords:
- vbawd10.chm165806110
ms.prod: WORD
api_name:
- Word.Frameset.FrameScrollbarType
ms.assetid: dacd6394-872e-beac-85dc-575234f9ce29
---


# Frameset.FrameScrollbarType Property (Word)

Returns or sets when scroll bars are available for the specified frame when viewing its frames page in a Web browser. Read/write  **WdScrollbarType** .


## Syntax

 _expression_ . **FrameScrollbarType**

 _expression_ Required. A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](http://msdn.microsoft.com/library/creating-frames-pages%28Office.15%29.aspx).


## Example

This example makes scroll bars always available for the specified frame, regardless of whether the contents of the frame require scrolling.


```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\Documents\Order.htm" 
 .FrameScrollBarType = wdScrollBarTypeYes 
End With
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

