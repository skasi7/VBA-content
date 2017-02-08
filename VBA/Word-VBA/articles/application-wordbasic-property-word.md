---
title: Application.WordBasic Property (Word)
keywords: vbawd10.chm158334977
f1_keywords:
- vbawd10.chm158334977
ms.prod: WORD
api_name:
- Word.Application.WordBasic
ms.assetid: 8c405ea6-0073-f994-42b2-cacb986f1f1f
---


# Application.WordBasic Property (Word)

Returns an automation object (Word.Basic) that includes methods for all the WordBasic statements and functions available in Word version 6.0 and Word for Windows 95. Read-only.


## Syntax

 _expression_ . **WordBasic**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

In Word 2000 and later, when you open a Word version 6.0 or Word for Windows 95 template that contains WordBasic macros, the macros are automatically converted to Visual Basic modules. Each WordBasic statement and function in the macro is converted to the corresponding Word.Basic method.

For information about WordBasic statements and functions, see the WordBasic Help in Word version 6.0 or Word for Windows 95. For information about converting WordBasic to Visual Basic, see [Converting WordBasic Macros to Visual Basic](http://msdn.microsoft.com/library/converting-wordbasic-macros-to-visual-basic%28Office.15%29.aspx). For general information, see [Conceptual Differences Between WordBasic and Visual Basic](http://msdn.microsoft.com/library/conceptual-differences-between-wordbasic-and-visual-basic%28Office.15%29.aspx).


## Example

This example uses the Word.Basic object to create a new document in Word version 6.0 or Word for Windows 95 and insert the available font names. Each font name is formatted in its corresponding font.


```vb
With WordBasic 
 .FileNewDefault 
 For aCount = 1 To .CountFonts() 
 .Font .[Font$](aCount) 
 .Insert .[Font$](aCount) 
 .InsertPara 
 Next 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

