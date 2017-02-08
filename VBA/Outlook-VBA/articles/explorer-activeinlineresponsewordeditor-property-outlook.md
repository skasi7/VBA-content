---
title: Explorer.ActiveInlineResponseWordEditor Property (Outlook)
keywords: vbaol11.chm3597
f1_keywords:
- vbaol11.chm3597
ms.assetid: b9058694-ab8f-4962-ab7d-afac1704dd29
---


# Explorer.ActiveInlineResponseWordEditor Property (Outlook)
Returns the Word [Document](http://msdn.microsoft.com/library/document-object-word%28Office.15%29.aspx) object of the active inline response that is displayed in the explorer Reading Pane. Read-only.

## Syntax

 _expression_ . **ActiveInlineResponseWordEditor**

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


## Remarks

This property returns  **Null** ( **Nothing** in Visual Basic) if no inline response is visible in the Reading Pane. The returned Word **Document** object provides access to most of the Word object model except for the following members:


- [InlineShapes.AddChart2](http://msdn.microsoft.com/library/inlineshapes-addchart2-method-word%28Office.15%29.aspx)
    
- [Range.ConvertToTable](http://msdn.microsoft.com/library/range-converttotable-method-word%28Office.15%29.aspx)
    
- [Range.ImportFragment](http://msdn.microsoft.com/library/range-importfragment-method-word%28Office.15%29.aspx)
    
- [Range.InsertXML](http://msdn.microsoft.com/library/range-insertxml-method-word%28Office.15%29.aspx)
    
- [Shapes.AddChart2](http://msdn.microsoft.com/library/shapes-addchart2-method-word%28Office.15%29.aspx)
    
- [Selection.InsertXML](http://msdn.microsoft.com/library/selection-insertxml-method-word%28Office.15%29.aspx)
    
- [Tables.Add](http://msdn.microsoft.com/library/tables-add-method-word%28Office.15%29.aspx)
    

## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

