---
title: Table Object (Word)
keywords: vbawd10.chm2385
f1_keywords:
- vbawd10.chm2385
ms.prod: WORD
api_name:
- Word.Table
ms.assetid: 996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6
---


# Table Object (Word)

Represents a single table. The  **Table** object is a member of the **[Tables](http://msdn.microsoft.com/library/tables-object-word%28Office.15%29.aspx)** collection. The **Tables** collection includes all the tables in the specified selection, range, or document.


## Remarks

Use  **Tables** (Index), where Index is the index number, to return a single **Table** object. The index number represents the position of the table in the selection, range, or document. The following example converts the first table in the active document to text.


```
ActiveDocument.Tables(1).ConvertToText Separator:=wdSeparateByTabs
```

Use the  **Add** method to add a table at the specified range. The following example adds a 3x4 table at the beginning of the active document.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Tables.Add Range:=myRange, NumRows:=3, NumColumns:=4
```


## Methods



|**Name**|
|:-----|
|[ApplyStyleDirectFormatting](http://msdn.microsoft.com/library/table-applystyledirectformatting-method-word%28Office.15%29.aspx)|
|[AutoFitBehavior](http://msdn.microsoft.com/library/table-autofitbehavior-method-word%28Office.15%29.aspx)|
|[AutoFormat](http://msdn.microsoft.com/library/table-autoformat-method-word%28Office.15%29.aspx)|
|[Cell](http://msdn.microsoft.com/library/table-cell-method-word%28Office.15%29.aspx)|
|[ConvertToText](http://msdn.microsoft.com/library/table-converttotext-method-word%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/table-delete-method-word%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/table-select-method-word%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/table-sort-method-word%28Office.15%29.aspx)|
|[SortAscending](http://msdn.microsoft.com/library/table-sortascending-method-word%28Office.15%29.aspx)|
|[SortDescending](http://msdn.microsoft.com/library/table-sortdescending-method-word%28Office.15%29.aspx)|
|[Split](http://msdn.microsoft.com/library/table-split-method-word%28Office.15%29.aspx)|
|[UpdateAutoFormat](http://msdn.microsoft.com/library/table-updateautoformat-method-word%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AllowAutoFit](http://msdn.microsoft.com/library/table-allowautofit-property-word%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/table-application-property-word%28Office.15%29.aspx)|
|[ApplyStyleColumnBands](http://msdn.microsoft.com/library/table-applystylecolumnbands-property-word%28Office.15%29.aspx)|
|[ApplyStyleFirstColumn](http://msdn.microsoft.com/library/table-applystylefirstcolumn-property-word%28Office.15%29.aspx)|
|[ApplyStyleHeadingRows](http://msdn.microsoft.com/library/table-applystyleheadingrows-property-word%28Office.15%29.aspx)|
|[ApplyStyleLastColumn](http://msdn.microsoft.com/library/table-applystylelastcolumn-property-word%28Office.15%29.aspx)|
|[ApplyStyleLastRow](http://msdn.microsoft.com/library/table-applystylelastrow-property-word%28Office.15%29.aspx)|
|[ApplyStyleRowBands](http://msdn.microsoft.com/library/table-applystylerowbands-property-word%28Office.15%29.aspx)|
|[AutoFormatType](http://msdn.microsoft.com/library/table-autoformattype-property-word%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/table-borders-property-word%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/table-bottompadding-property-word%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/table-columns-property-word%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/table-creator-property-word%28Office.15%29.aspx)|
|[Descr](http://msdn.microsoft.com/library/table-descr-property-word%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/table-id-property-word%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/table-leftpadding-property-word%28Office.15%29.aspx)|
|[NestingLevel](http://msdn.microsoft.com/library/table-nestinglevel-property-word%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/table-parent-property-word%28Office.15%29.aspx)|
|[PreferredWidth](http://msdn.microsoft.com/library/table-preferredwidth-property-word%28Office.15%29.aspx)|
|[PreferredWidthType](http://msdn.microsoft.com/library/table-preferredwidthtype-property-word%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/table-range-property-word%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/table-rightpadding-property-word%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/table-rows-property-word%28Office.15%29.aspx)|
|[Shading](http://msdn.microsoft.com/library/table-shading-property-word%28Office.15%29.aspx)|
|[Spacing](http://msdn.microsoft.com/library/table-spacing-property-word%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/table-style-property-word%28Office.15%29.aspx)|
|[TableDirection](http://msdn.microsoft.com/library/table-tabledirection-property-word%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/table-tables-property-word%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/table-title-property-word%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/table-toppadding-property-word%28Office.15%29.aspx)|
|[Uniform](http://msdn.microsoft.com/library/table-uniform-property-word%28Office.15%29.aspx)|

## See also


#### Other resources


<<<<<<< HEAD
=======

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5
[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)

