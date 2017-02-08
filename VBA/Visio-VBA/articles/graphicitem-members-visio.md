---
title: GraphicItem Members (Visio)
ms.prod: VISIO
ms.assetid: a6e2de62-e0c6-ef87-df2e-ae3dc15702d4
---


# GraphicItem Members (Visio)
Represents a single component part of a data graphic master (a  **[Master](master-object-visio.md)** object of type **visTypeDataGraphic** ) that is responsible for a specific graphical adornment of the master.

Represents a single component part of a data graphic master (a  **[Master](master-object-visio.md)** object of type **visTypeDataGraphic** ) that is responsible for a specific graphical adornment of the master.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](graphicitem-delete-method-visio.md)|Deletes a  **GraphicItem** object from the **GraphicItems** collection of a **Master** object of type **visTypeDataGraphic** .|
|[GetExpression](graphicitem-getexpression-method-visio.md)|Gets the label of the shape data item (custom property) that the  **GraphicItem** represents, or the value of the expression string that is part of a **GraphicItem** object?s rule, against which shape data is evaluated.|
|[SetExpression](graphicitem-setexpression-method-visio.md)|Sets the value of the expression string that is part of a  **GraphicItem** object?s rule, against which shape data (custom properties) are evaluated.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](graphicitem-application-property-visio.md)|Returns the instance of Microsoft Visio associated with a  **GraphicItem** object. Read-only.|
|[DataGraphic](graphicitem-datagraphic-property-visio.md)|Returns the  **Master** object of type **visTypeDataGraphic** that contains the graphic item. Read-only.|
|[Description](graphicitem-description-property-visio.md)|Returns a string that describes the type of the graphic item. Read-only.|
|[Document](graphicitem-document-property-visio.md)|Gets the  **Document** object that contains the **Master** object of type **visTypeDataGraphic** that contains the **GraphicItem** object. Read-only.|
|[HorizontalPosition](graphicitem-horizontalposition-property-visio.md)|Gets or sets the horizontal position of the  **GraphicItem** object relative to the shape to which it is applied. Read/write.|
|[ID](graphicitem-id-property-visio.md)|Gets the unique identifier of the  **GraphicItem** object. Read-only.|
|[Index](graphicitem-index-property-visio.md)|Gets or sets the ordinal position of a  **GraphicItem** object in the **GraphicItems** collection of a data graphic master—a **Master** object of type **visTypeDataGraphic** . Read/write.|
|[ObjectType](graphicitem-objecttype-property-visio.md)|Returns  **visObjTypeGraphicItem** , the type of a **GraphicItem** object. Read-only.|
|[Stat](graphicitem-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[Tag](graphicitem-tag-property-visio.md)|Gets or sets a user-defined string expression that can store extra data related to your program. Read/write.|
|[Type](graphicitem-type-property-visio.md)|Returns the type of the graphic item. Read-only.|
|[UseDataGraphicPosition](graphicitem-usedatagraphicposition-property-visio.md)|Gets or sets whether to use the current default callout position for graphic items of the data graphic master to whose  **GraphicItems** collection the graphic item belongs. Read/write.|
|[VerticalPosition](graphicitem-verticalposition-property-visio.md)|Gets or sets the vertical position of the  **GraphicItem** object relative to the shape to which it is applied. Read/write.|

