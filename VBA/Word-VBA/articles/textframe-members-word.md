---
title: TextFrame Members (Word)
ms.prod: WORD
ms.assetid: bb2efcc6-474f-3de5-6d20-940be7549112
---


# TextFrame Members (Word)
Represents the text frame in a  **Shape** object. The **TextFrame** object contains the text in the text frame and the properties that control the margins and orientation of the text frame.

Represents the text frame in a  **Shape** object. The **TextFrame** object contains the text in the text frame and the properties that control the margins and orientation of the text frame.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[BreakForwardLink](textframe-breakforwardlink-method-word.md)|Breaks the forward link for the specified text frame, if such a link exists.|
|[DeleteText](textframe-deletetext-method-word.md)|Deletes the text from a text frame and all the associated properties of the text, including font attributes.|
|[ValidLinkTarget](textframe-validlinktarget-method-word.md)|Determines whether the text frame of one shape can be linked to the text frame of another shape. .|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](textframe-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoSize](textframe-autosize-property-word.md)|Returns or sets a  **Long** that represents whether a text frame is sized automatically. Read/write.|
|[Column](textframe-column-property-word.md)|This object, member, or enumeration is deprecated and is not intended to be used in your code. |
|[ContainingRange](textframe-containingrange-property-word.md)|Returns a  **[Range](range-object-word.md)** object that represents the entire story in a series of shapes with linked text frames that the specified text frame belongs to. Read-only.|
|[Creator](textframe-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[HasText](textframe-hastext-property-word.md)| **True** if the specified shape has text associated with it. Read-only **Boolean** .|
|[HorizontalAnchor](textframe-horizontalanchor-property-word.md)|Returns or sets the horizontal alignment of text in a text frame. Read/write  **[MsoHorizontalAnchor](msohorizontalanchor-enumeration-office.md)** .|
|[MarginBottom](textframe-marginbottom-property-word.md)|Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .|
|[MarginLeft](textframe-marginleft-property-word.md)|Returns or sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .|
|[MarginRight](textframe-marginright-property-word.md)|Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .|
|[MarginTop](textframe-margintop-property-word.md)|Returns or sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .|
|[Next](textframe-next-property-word.md)|Returns a  **TextFrame** object that represents the next text frame in a collection of shapes. Read-only.|
|[NoTextRotation](textframe-notextrotation-property-word.md)|True if text in the text frame should not rotate when the shape is rotated. Read/write [MsoTriState](msotristate-enumeration-office.md).|
|[Orientation](textframe-orientation-property-word.md)|Returns or sets the orientation of the text inside the frame. Read/write  **MsoTextOrientation** .|
|[Overflowing](textframe-overflowing-property-word.md)| **True** if the text inside the specified text frame doesn't all fit within the frame. Read-only **Boolean** .|
|[Parent](textframe-parent-property-word.md)|Returns a  **Shape** object that represents the parent shape of the text frame.|
|[PathFormat](textframe-pathformat-property-word.md)|Returns or sets the path type for the specified text frame. Read/write  **MsoPathType** .|
|[Previous](textframe-previous-property-word.md)|Returns a  **TextFrame** object that represents the previous text frame in a collection of shapes. Read-only.|
|[TextRange](textframe-textrange-property-word.md)|Returns a  **[Range](range-object-word.md)** object that represents the text in the specified text frame.|
|[ThreeD](textframe-threed-property-word.md)|Returns a [ThreeDFormat](threedformat-object-word.md) object that contains 3-D effect formatting properties for the specified text frame. Read-only.|
|[VerticalAnchor](textframe-verticalanchor-property-word.md)|Returns or sets an  **MsoVerticalAnchor** constant that represents the vertical alignment of the text within a shape. Read/write.|
|[WarpFormat](textframe-warpformat-property-word.md)|Returns or sets the warp format (how the text is warped) for the specified text frame. Read/write [MsoWarpFormat](msowarpformat-enumeration-office.md).|
|[WordWrap](textframe-wordwrap-property-word.md)| **True** if Microsoft Word wraps Latin text in the middle of a word in the specified text frames. Read/write **Long** . .|

