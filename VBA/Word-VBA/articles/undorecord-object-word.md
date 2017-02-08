---
title: UndoRecord Object (Word)
keywords: vbawd10.chm856
f1_keywords:
- vbawd10.chm856
ms.prod: WORD
api_name:
- Word.UndoRecord
ms.assetid: 77bf9801-e940-e661-6bbe-20a8714d5dbd
---


# UndoRecord Object (Word)

Provides an entry point into the undo stack.


## Remarks

Use the  **UndoRecord** object to create and modify custom undo records in the Word undo stack.


## Example

The following code example instantiates an  **UndoRecord** object.


```vb
Dim objUndo As UndoRecord 
Set objUndo = Application.UndoRecord
```


## See also


#### Other resources


[Working With the UndoRecord Object](http://msdn.microsoft.com/library/working-with-the-undorecord-object%28Office.15%29.aspx)
[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)


