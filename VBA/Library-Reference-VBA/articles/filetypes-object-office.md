---
title: FileTypes Object (Office)
keywords: vbaof11.chm257000
f1_keywords:
- vbaof11.chm257000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.FileTypes
ms.assetid: 5e8b5240-5ebd-704d-72e6-1f4ad951dfdc
---


# FileTypes Object (Office)

A collection of values of the type  **msoFileType** that determine which types of files are returned during a search.


## Remarks

There is only one  **FileTypes** collection for all searches so it's important to clear the **FileTypes** collection before executing a search unless you wish to search for file types from previous searches. The easiest way to clear the collection is to set the **FileType** property to the first file type for which you want to search. You can also remove individual types using the **Remove** method. To determine the file type of each item in the collection, use the **Item** method to return the **msoFileType** value.


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/filetypes-add-method-office%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/filetypes-remove-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/filetypes-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/filetypes-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/filetypes-creator-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/filetypes-item-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[FileTypes Object Members](http://msdn.microsoft.com/library/filetypes-members-office%28Office.15%29.aspx)
