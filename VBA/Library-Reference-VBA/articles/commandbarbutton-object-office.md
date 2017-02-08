---
title: CommandBarButton Object (Office)
keywords: vbaof11.chm244000
f1_keywords:
- vbaof11.chm244000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarButton
ms.assetid: e6d8209d-2c87-f1b5-bc3f-d4e5e5d3ab73
---


# CommandBarButton Object (Office)

Represents a button control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use  **Controls(index)**, where _index_ is the index number of the control, to return a **CommandBarButton** object. Note that the **Type** property of the control must be **msoControlButton**. Assuming that the second control on the command bar named "Custom" is a button, the following example changes the style of that button.


```
Set c = CommandBars("Custom").Controls(2) 
With c 
If .Type = msoControlButton Then 
    If .Style = msoButtonIcon Then 
        .Style = msoButtonIconAndCaption 
    Else 
        .Style = msoButtonIcon 
    End If 
End If 
End With
```


 **Note**  


 **Note**  You can also use the  **FindControl** method to return a **CommandBarButton** object.


## Events



|**Name**|
|:-----|
|[Click](http://msdn.microsoft.com/library/commandbarbutton-click-event-office%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/commandbarbutton-copy-method-office%28Office.15%29.aspx)|
|[CopyFace](http://msdn.microsoft.com/library/commandbarbutton-copyface-method-office%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/commandbarbutton-delete-method-office%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/commandbarbutton-execute-method-office%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/commandbarbutton-move-method-office%28Office.15%29.aspx)|
|[PasteFace](http://msdn.microsoft.com/library/commandbarbutton-pasteface-method-office%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/commandbarbutton-reset-method-office%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/commandbarbutton-setfocus-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/commandbarbutton-application-property-office%28Office.15%29.aspx)|
|[BeginGroup](http://msdn.microsoft.com/library/commandbarbutton-begingroup-property-office%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/commandbarbutton-builtin-property-office%28Office.15%29.aspx)|
|[BuiltInFace](http://msdn.microsoft.com/library/commandbarbutton-builtinface-property-office%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/commandbarbutton-caption-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/commandbarbutton-creator-property-office%28Office.15%29.aspx)|
|[DescriptionText](http://msdn.microsoft.com/library/commandbarbutton-descriptiontext-property-office%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/commandbarbutton-enabled-property-office%28Office.15%29.aspx)|
|[FaceId](http://msdn.microsoft.com/library/commandbarbutton-faceid-property-office%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/commandbarbutton-height-property-office%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/commandbarbutton-helpcontextid-property-office%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/commandbarbutton-helpfile-property-office%28Office.15%29.aspx)|
|[HyperlinkType](http://msdn.microsoft.com/library/commandbarbutton-hyperlinktype-property-office%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/commandbarbutton-id-property-office%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/commandbarbutton-index-property-office%28Office.15%29.aspx)|
|[IsPriorityDropped](http://msdn.microsoft.com/library/commandbarbutton-isprioritydropped-property-office%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/commandbarbutton-left-property-office%28Office.15%29.aspx)|
|[Mask](http://msdn.microsoft.com/library/commandbarbutton-mask-property-office%28Office.15%29.aspx)|
|[OLEUsage](http://msdn.microsoft.com/library/commandbarbutton-oleusage-property-office%28Office.15%29.aspx)|
|[OnAction](http://msdn.microsoft.com/library/commandbarbutton-onaction-property-office%28Office.15%29.aspx)|
|[Parameter](http://msdn.microsoft.com/library/commandbarbutton-parameter-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/commandbarbutton-parent-property-office%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/commandbarbutton-picture-property-office%28Office.15%29.aspx)|
|[Priority](http://msdn.microsoft.com/library/commandbarbutton-priority-property-office%28Office.15%29.aspx)|
|[ShortcutText](http://msdn.microsoft.com/library/commandbarbutton-shortcuttext-property-office%28Office.15%29.aspx)|
|[State](http://msdn.microsoft.com/library/commandbarbutton-state-property-office%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/commandbarbutton-style-property-office%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/commandbarbutton-tag-property-office%28Office.15%29.aspx)|
|[TooltipText](http://msdn.microsoft.com/library/commandbarbutton-tooltiptext-property-office%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/commandbarbutton-top-property-office%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/commandbarbutton-type-property-office%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/commandbarbutton-visible-property-office%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/commandbarbutton-width-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[CommandBarButton Object Members](http://msdn.microsoft.com/library/commandbarbutton-members-office%28Office.15%29.aspx)
