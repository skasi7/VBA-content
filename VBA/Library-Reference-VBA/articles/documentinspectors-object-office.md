---
title: DocumentInspectors Object (Office)
keywords: vbaof11.chm278000
f1_keywords:
- vbaof11.chm278000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.DocumentInspectors
ms.assetid: 8366d7cd-e016-bb99-d27f-749ca10352f1
---


# DocumentInspectors Object (Office)

Represents a collection of  **DocumentInspector** objects.


## Remarks

The  **DocumentInspectors** collection is part of the **Document** object in Microsoft Word, the **Workbook** object in Microsoft Excel, and the **Presentation** object in MicrosoftPowerPoint. A **DocumentInspectors** collection contains multiple **DocumentInspector** objects, one for some built-in options and each installed custom Document Inspector module. For more information, see the **DocumentInspector** help topic.


## Example

The following example calls the  **Fix** method of a Document Inspector module and displays the status of the action and the specific items that are removed.


```
Public Sub FixDocument() 
Dim docStatus As MsoDocInspectorStatus 
Dim results As String 
 ActiveDocument.DocumentInspectors(3).Fix docStatus, results 
 
 MsgBox docStatus 
 MsgBox("The following items were removed " &amp; results) 
 
End Sub 

```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/documentinspectors-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/documentinspectors-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/documentinspectors-creator-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/documentinspectors-item-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/documentinspectors-parent-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[DocumentInspectors Object Members](http://msdn.microsoft.com/library/documentinspectors-members-office%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
