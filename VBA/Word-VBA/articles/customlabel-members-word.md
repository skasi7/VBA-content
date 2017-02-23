---
title: CustomLabel Members (Word)
ms.prod: WORD
ms.assetid: 92ab60f7-48c8-151c-df5a-31aa885ec269
---


# CustomLabel Members (Word)
Represents a custom mailing label. The  **CustomLabel** object is a member of the **[CustomLabels](customlabels-object-word.md)** collection. The **CustomLabels** collection contains all the custom mailing labels listed in the **Label Options** dialog box.

Represents a custom mailing label. The  **CustomLabel** object is a member of the **[CustomLabels](customlabels-object-word.md)** collection. The **CustomLabels** collection contains all the custom mailing labels listed in the **Label Options** dialog box.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](customlabel-delete-method-word.md)|Deletes the specified custom label.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](customlabel-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](customlabel-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DotMatrix](customlabel-dotmatrix-property-word.md)| **True** if the printer type for the specified custom label is dot matrix. **False** if the printer type is either laser or ink jet. Read-only **Boolean** .|
|[Height](customlabel-height-property-word.md)|Returns or sets the height of a specified custom mailing label, in points. Read/write  **Single** .|
|[HorizontalPitch](customlabel-horizontalpitch-property-word.md)|Returns or sets the horizontal distance (in points) between the left edge of one custom mailing label and the left edge of the next mailing label. Read/write  **Single** .|
|[Index](customlabel-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[Name](customlabel-name-property-word.md)|Returns or sets the name of the specified object. Read/write  **CustomLabel** .|
|[NumberAcross](customlabel-numberacross-property-word.md)|Returns or sets the number of custom mailing labels across a page. Read/write  **Long** .|
|[NumberDown](customlabel-numberdown-property-word.md)|Returns or sets the number of custom mailing labels down the length of a page. Read/write  **Long** .|
|[PageSize](customlabel-pagesize-property-word.md)|Returns or sets the page size for the specified custom mailing label. Read/write  **WdCustomLabelPageSize** .|
|[Parent](customlabel-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **CustomLabel** object.|
|[SideMargin](customlabel-sidemargin-property-word.md)|Returns or sets the side margin widths (in points) for the specified custom mailing label. Read/write  **Single** .|
|[TopMargin](customlabel-topmargin-property-word.md)|Returns or sets the distance (in points) between the top edge of the page and the top boundary of the body text. Read/write  **Single** .|
|[Valid](customlabel-valid-property-word.md)| **True** if the various properties (for example, **Height** , **Width** , and **NumberDown** ) for the specified custom label work together to produce a valid mailing label. Read-only **Boolean** .|
|[VerticalPitch](customlabel-verticalpitch-property-word.md)|Returns or sets the vertical distance between the top of one mailing label and the top of the next mailing label. Read/write  **Single** .|
|[Width](customlabel-width-property-word.md)|Returns or sets the width of a custom mailing label, in points. Read/write  **Long** .|

