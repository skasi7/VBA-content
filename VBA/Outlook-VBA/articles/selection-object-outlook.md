---
title: Selection Object (Outlook)
keywords: vbaol11.chm80
f1_keywords:
- vbaol11.chm80
ms.prod: OUTLOOK
api_name:
- Outlook.Selection
ms.assetid: 0b06a3ce-0445-db8f-e6e8-bb7bd469c50f
---


# Selection Object (Outlook)

Contains the set of Outlook items currently selected in an explorer.


## Remarks

Use the  **[Selection](http://msdn.microsoft.com/library/explorer-selection-property-outlook%28Office.15%29.aspx)** property to return the **Selection** collection from the **[Explorer](explorer-object-outlook.md)** object.


## Example

The following example returns a  **Selection** object from an **Explorer** object.


```
Set mySelectedItems = myExplorer.Selection
```


## Methods



|**Name**|
|:-----|
|[GetSelection](http://msdn.microsoft.com/library/selection-getselection-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/selection-item-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/selection-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/selection-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/selection-count-property-outlook%28Office.15%29.aspx)|
|[Location](http://msdn.microsoft.com/library/selection-location-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/selection-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/selection-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[Selection Object Members](http://msdn.microsoft.com/library/selection-members-outlook%28Office.15%29.aspx)
