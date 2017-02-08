---
title: BuildingBlockEntries Object (Word)
keywords: vbawd10.chm553
f1_keywords:
- vbawd10.chm553
ms.prod: WORD
api_name:
- Word.BuildingBlockEntries
ms.assetid: 9c5946e9-947d-7284-ab16-b570bf7f0ff3
---


# BuildingBlockEntries Object (Word)

Represents a collection of all  **[BuildingBlock](buildingblock-object-word.md)** objects in a template.


## Remarks

Use the  **[Add](buildingblockentries-add-method-word.md)** method to create a new building block and add it to a template. The following example adds the selected text to the watermarks building block gallery of the first template in the **[Templates](templates-object-word.md)** collection.


```
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
 
Set objBB = objTemplate.BuildingBlockEntries _ 
 .Add(Name:="New Building Block Entry", _ 
 Type:=wdTypeWatermarks, _ 
 Category:="General", _ 
 Range:=Selection.Range)
```

Unlike the  **Add** method for the **BuildingBlocks** collection, you need to specify the type and category when you add a building block using the **Add** method of the **BuildingBlockEntries** collection. This is because building blocks are organized by using types and categories. When you use the **BuildingBlockEntries** collection, you are accessing the entire collection of building blocks in a template; however, when you use the **BuildingBlocks** collection, you are accessing the collection of building blocks for a specific type and category in a template.


 **Note**  Using the  **Category** and **Type** properties for the **BuildingBlock** object enables you to determine the category and type for a building block.

For more information about building blocks, see [Working with Building Blocks](http://msdn.microsoft.com/library/working-with-building-blocks%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[Add](buildingblockentries-add-method-word.md)|
|[Item](buildingblockentries-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](buildingblockentries-application-property-word.md)|
|[Count](buildingblockentries-count-property-word.md)|
|[Creator](buildingblockentries-creator-property-word.md)|
|[Parent](buildingblockentries-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)
