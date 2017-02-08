---
title: ProtectedViewWindows Object (Word)
ms.prod: WORD
api_name:
- Word.ProtectedViewWindows
ms.assetid: 62c2f4d5-1080-548e-730b-388308144dfe
---


# ProtectedViewWindows Object (Word)

A collection of all the [ProtectedViewWindow](protectedviewwindow-object-word.md) objects that are currently open in Word.


## Remarks

Use the  **ProtectedViewWindows** property to return the **ProtectedViewWindows** collection.


## Example

The following code example displays the number of protected view windows that are open.


```vb
MsgBox "There are " &; ProtectedViewWindows.Count &; _ 
 " protected view windows currently open."
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)


