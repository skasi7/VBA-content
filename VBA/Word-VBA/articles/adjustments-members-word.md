---
title: Adjustments Members (Word)
ms.prod: WORD
ms.assetid: 68793a8c-b1c0-d0ea-0b06-f40edcf6ee71
---


# Adjustments Members (Word)
Contains a collection of adjustment values for the specified AutoShape or WordArt object. Each adjustment value represents one way an adjustment handle can be adjusted. Because some adjustment handles can be adjusted in two ways ? for instance, some handles can be adjusted both horizontally and vertically ? a shape can have more adjustment values than it has adjustment handles. A shape can have up to eight adjustments.

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](adjustments-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Count](adjustments-count-property-word.md)|Returns the number of items in the  **Adjustments** collection. Read-only **Long** .|
|[Creator](adjustments-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Item](adjustments-item-property-word.md)|Returns or sets the adjustment value specified by the  _Index_ argument. Read/write **Single** .|
|[Parent](adjustments-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified object. This is usually a **Shape** or **ShapeRange** object.|

