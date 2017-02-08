---
title: TextEffectFormat Members (Excel)
ms.prod: EXCEL
ms.assetid: 10d920d6-b96f-7afa-8e27-c22ba0926146
---


# TextEffectFormat Members (Excel)
Contains properties and methods that apply to WordArt objects.

Contains properties and methods that apply to WordArt objects.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ToggleVerticalText](texteffectformat-toggleverticaltext-method-excel.md)|Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Alignment](texteffectformat-alignment-property-excel.md)|Returns or sets an  **[MsoTextEffectAlignment](msotexteffectalignment-enumeration-office.md)** value that represents the alignment for WordArt.|
|[Application](texteffectformat-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](texteffectformat-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[FontBold](texteffectformat-fontbold-property-excel.md)| **True** if the font in the specified WordArt is bold. Read/write **[MsoTriState](msotristate-enumeration-office.md)** .|
|[FontItalic](texteffectformat-fontitalic-property-excel.md)|Returns  **msoTrue** if the font in the specified WordArt is italic. Read/write **[MsoTriState](msotristate-enumeration-office.md)** .|
|[FontName](texteffectformat-fontname-property-excel.md)|Returns or sets the name of the font in the specified WordArt. Read/write  **String** .|
|[FontSize](texteffectformat-fontsize-property-excel.md)|Returns or sets the font size for the specified WordArt, in points. Read/write  **Single** .|
|[KernedPairs](texteffectformat-kernedpairs-property-excel.md)| **True** if character pairs in the specified WordArt are kerned. Read/write **MsoTriState** .|
|[NormalizedHeight](texteffectformat-normalizedheight-property-excel.md)| **True** if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write **MsoTriState** .|
|[Parent](texteffectformat-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PresetShape](texteffectformat-presetshape-property-excel.md)|Returns or sets the shape of the specified WordArt. Read/write  **MsoPresetTextEffectShape** .|
|[PresetTextEffect](texteffectformat-presettexteffect-property-excel.md)|Returns or sets the style of the specified WordArt. Read/write  **MsoPresetTextEffect** .|
|[RotatedChars](texteffectformat-rotatedchars-property-excel.md)| **True** if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. **False** if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write **MsoTriState** .|
|[Text](texteffectformat-text-property-excel.md)|Returns or sets the text for the specified object. Read/write  **String** .|
|[Tracking](texteffectformat-tracking-property-excel.md)|Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt to the width of the character. Can be a value from 0 (zero) through 5. (Large values for this property specify ample space between characters; values less than 1 can produce character overlap.) Read/write  **Single** .|

