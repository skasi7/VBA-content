---
title: CoAuthors Object (Word)
ms.prod: WORD
api_name:
- Word.CoAuthors
ms.assetid: 47fc864d-5f1b-b113-85b5-6e8b1b75c225
---


# CoAuthors Object (Word)

A collection of all the  **[CoAuthor](coauthor-object-word.md)** objects in the document.


## Remarks

The  **CoAuthors** collection contains all the co-authors in the document (authors that are actively editing the document).


## Example

The following code example gets the number of co-authors in the active document.


```vb
Dim i As Integer 
 
i = ActiveDocument.CoAuthoring.Authors.Count 
 
MsgBox "The number of co-authors is " &; i
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)


