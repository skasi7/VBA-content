---
title: Hyperlink Object (PowerPoint)
keywords: vbapp10.chm526000
f1_keywords:
- vbapp10.chm526000
ms.prod: POWERPOINT
ms.assetid: c8d53079-b280-c93c-a3c9-b865d09abe1a
---


# Hyperlink Object (PowerPoint)

Represents a hyperlink associated with a non-placeholder shape or text. 


## Remarks

You can use a hyperlink to jump to an Internet or intranet site, to another file, or to a slide within the active presentation. The  **Hyperlink** object is a member of the **[Hyperlinks](http://msdn.microsoft.com/library/hyperlinks-object-powerpoint%28Office.15%29.aspx)** collection. The **Hyperlinks** collection contains all the hyperlinks on a slide or a master.


## Example

Use the [Hyperlink](http://msdn.microsoft.com/library/actionsetting-hyperlink-property-powerpoint%28Office.15%29.aspx)property to return a hyperlink for a shape. A shape can have two different hyperlinks assigned to it: one that is followed when the user clicks the shape during a slide show, and another that is followed when the user passes the mouse pointer over the shape during a slide show. For the hyperlink to be active during a slide show, the  **Action** property must be set to **ppActionHyperlink**. The following example sets the mouse-click action for shape three on slide one in the active presentation to an Internet link.


```
With ActivePresentation.Slides(1).Shapes(3) _

        .ActionSettings(ppMouseClick)

    .Action = ppActionHyperlink

    .Hyperlink.Address = "http://www.microsoft.com"

End With
```

A slide can contain more than one hyperlink. Each non-placeholder shape can have a hyperlink; the text within a shape can have its own hyperlink; and each individual character can have its own hyperlink. Use  **Hyperlinks** (index), where index is the hyperlink number, to return a single **Hyperlink** object. The following example adds the shape three mouse-click hyperlink to the Favorites folder.




```
ActivePresentation.Slides(1).Shapes(3) _

    .ActionSettings(ppMouseClick).Hyperlink.AddToFavorites
```


 **Note**  When you use this method to add a hyperlink to the Internet Explorer Favorites folder, an icon is added to the  **Favorites** menu without a corresponding name. You must add the name from within Internet Explorer.


## Methods



|**Name**|
|:-----|
|[AddToFavorites](http://msdn.microsoft.com/library/hyperlink-addtofavorites-method-powerpoint%28Office.15%29.aspx)|
|[CreateNewDocument](http://msdn.microsoft.com/library/hyperlink-createnewdocument-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/hyperlink-delete-method-powerpoint%28Office.15%29.aspx)|
|[Follow](http://msdn.microsoft.com/library/hyperlink-follow-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/hyperlink-address-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/hyperlink-application-property-powerpoint%28Office.15%29.aspx)|
|[EmailSubject](http://msdn.microsoft.com/library/hyperlink-emailsubject-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/hyperlink-parent-property-powerpoint%28Office.15%29.aspx)|
|[ScreenTip](http://msdn.microsoft.com/library/hyperlink-screentip-property-powerpoint%28Office.15%29.aspx)|
|[ShowAndReturn](http://msdn.microsoft.com/library/hyperlink-showandreturn-property-powerpoint%28Office.15%29.aspx)|
|[SubAddress](http://msdn.microsoft.com/library/hyperlink-subaddress-property-powerpoint%28Office.15%29.aspx)|
|[TextToDisplay](http://msdn.microsoft.com/library/hyperlink-texttodisplay-property-powerpoint%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/hyperlink-type-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
