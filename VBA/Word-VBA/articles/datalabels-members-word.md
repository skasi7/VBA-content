---
title: DataLabels Members (Word)
ms.prod: WORD
ms.assetid: 4b219908-2cdc-1c13-d243-b3a7c47c9987
---


# DataLabels Members (Word)
A collection of all the  **[DataLabel](datalabel-object-word.md)** objects for the specified series.

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](datalabels-delete-method-word.md)|Deletes the object.|
|[Item](datalabels-item-method-word.md)|Returns a single object from a collection.|
|[Propagate](datalabels-propagate-method-word.md)|Propagates the contents and formatting of the specified data label to all the other data labels in the series.|
|[Select](datalabels-select-method-word.md)|Selects the object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](datalabels-application-property-word.md)|When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
|[AutoText](datalabels-autotext-property-word.md)| **True** if all objects in the collection automatically generate appropriate text based on context. Read/write **Boolean** .|
|[Count](datalabels-count-property-word.md)|Returns the number of objects in the collection. Read-only  **Long** .|
|[Creator](datalabels-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Format](datalabels-format-property-word.md)|Returns the line, fill, and effect formatting for the object. Read-only  **[ChartFormat](chartformat-object-word.md)** .|
|[HorizontalAlignment](datalabels-horizontalalignment-property-word.md)|Returns or sets the horizontal alignment for the specified object. Read/write  **Variant** .|
|[Name](datalabels-name-property-word.md)|Returns the name of the object. Read-only  **String** .|
|[NumberFormat](datalabels-numberformat-property-word.md)|Returns or sets the format code for the object. Read/write  **String** .|
|[NumberFormatLinked](datalabels-numberformatlinked-property-word.md)| **True** if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write **Boolean** .|
|[NumberFormatLocal](datalabels-numberformatlocal-property-word.md)|Returns or sets the format code for the object as a string in the language of the user. Read/write  **Variant** .|
|[Orientation](datalabels-orientation-property-word.md)|Returns or sets the text orientation. Read/write  **Long** .|
|[Parent](datalabels-parent-property-word.md)|Returns the parent for the specified object. Read-only  **Object** .|
|[Position](datalabels-position-property-word.md)|Returns or sets the position of the data labels. Read/write  **[XlDataLabelPosition](xldatalabelposition-enumeration-word.md)** .|
|[ReadingOrder](datalabels-readingorder-property-word.md)|Returns or sets an  **[XlReadingOrder](xlreadingorder-enumeration-word.md)** constant that represents the reading order for the specified object. Read/write **Long** .|
|[Separator](datalabels-separator-property-word.md)|Sets or returns the separator for the data labels on a chart. Read/write  **Variant** .|
|[Shadow](datalabels-shadow-property-word.md)|Returns or sets a value that indicates whether the object has a shadow. Read/write  **Boolean** .|
|[ShowBubbleSize](datalabels-showbubblesize-property-word.md)| **True** to show the bubble size for the data labels on a chart. **False** to hide the bubble size. Read/write **Boolean** .|
|[ShowCategoryName](datalabels-showcategoryname-property-word.md)| **True** to display the category name for the data labels on a chart. **False** to hide the name. Read/write **Boolean** .|
|[ShowLegendKey](datalabels-showlegendkey-property-word.md)| **True** if the data label legend key is visible. Read/write **Boolean** .|
|[ShowPercentage](datalabels-showpercentage-property-word.md)| **True** to display the percentage value for the data labels on a chart. **False** to hide the value. Read/write **Boolean** .|
|[ShowRange](datalabels-showrange-property-word.md)|Set to  **True** to display the **Value From Cells** range field in all the chart data labels for a specified chart. Set to **False** to hide that field. Read/write **Boolean**.|
|[ShowSeriesName](datalabels-showseriesname-property-word.md)| **True** to show the series name for the data labels on a chart. **False** to hide the name. Read/write **Boolean** .|
|[ShowValue](datalabels-showvalue-property-word.md)| **True** to display the data label values for a specified chart. **False** to hide the values. Read/write **Boolean** .|
|[VerticalAlignment](datalabels-verticalalignment-property-word.md)|Returns or sets the vertical alignment of the specified object. Read/write  **Variant** .|

