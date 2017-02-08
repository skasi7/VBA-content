---
title: CommandBarPopup Object (Office)
keywords: vbaof11.chm7000
f1_keywords:
- vbaof11.chm7000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarPopup
ms.assetid: a8ae06a3-1d7b-a531-91df-756fafee5314
---


# CommandBarPopup Object (Office)

Represents a pop-up control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

Every pop-up control contains a  **CommandBar** object. To return the command bar from a pop-up control, apply the **CommandBar** property to the **CommandBarPopup** object.

 Use Controls(index), where _index_ is the number of the control, to return a **CommandBarPopup** object. Note that the **Type** property of the control must be **msoControlPopup**, **msoControlGraphicPopup**, **msoControlButtonPopup**, **msoControlSplitButtonPopup**, or **msoControlSplitButtonMRUPopup**.


## Example

You can also use the  **FindControl** method to return a **CommandBarPopup** object. The following example searches all command bars for a **CommandBarPopup** object whose tag is "Graphics."


```
Set myControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics")
```


## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/commandbarpopup-copy-method-office%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/commandbarpopup-delete-method-office%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/commandbarpopup-execute-method-office%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/commandbarpopup-move-method-office%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/commandbarpopup-reset-method-office%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/commandbarpopup-setfocus-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/commandbarpopup-application-property-office%28Office.15%29.aspx)|
|[BeginGroup](http://msdn.microsoft.com/library/commandbarpopup-begingroup-property-office%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/commandbarpopup-builtin-property-office%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/commandbarpopup-caption-property-office%28Office.15%29.aspx)|
|[CommandBar](http://msdn.microsoft.com/library/commandbarpopup-commandbar-property-office%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/commandbarpopup-controls-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/commandbarpopup-creator-property-office%28Office.15%29.aspx)|
|[DescriptionText](http://msdn.microsoft.com/library/commandbarpopup-descriptiontext-property-office%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/commandbarpopup-enabled-property-office%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/commandbarpopup-height-property-office%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/commandbarpopup-helpcontextid-property-office%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/commandbarpopup-helpfile-property-office%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/commandbarpopup-id-property-office%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/commandbarpopup-index-property-office%28Office.15%29.aspx)|
|[IsPriorityDropped](http://msdn.microsoft.com/library/commandbarpopup-isprioritydropped-property-office%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/commandbarpopup-left-property-office%28Office.15%29.aspx)|
|[OLEMenuGroup](http://msdn.microsoft.com/library/commandbarpopup-olemenugroup-property-office%28Office.15%29.aspx)|
|[OLEUsage](http://msdn.microsoft.com/library/commandbarpopup-oleusage-property-office%28Office.15%29.aspx)|
|[OnAction](http://msdn.microsoft.com/library/commandbarpopup-onaction-property-office%28Office.15%29.aspx)|
|[Parameter](http://msdn.microsoft.com/library/commandbarpopup-parameter-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/commandbarpopup-parent-property-office%28Office.15%29.aspx)|
|[Priority](http://msdn.microsoft.com/library/commandbarpopup-priority-property-office%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/commandbarpopup-tag-property-office%28Office.15%29.aspx)|
|[TooltipText](http://msdn.microsoft.com/library/commandbarpopup-tooltiptext-property-office%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/commandbarpopup-top-property-office%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/commandbarpopup-type-property-office%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/commandbarpopup-visible-property-office%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/commandbarpopup-width-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[CommandBarPopup Object Members](http://msdn.microsoft.com/library/commandbarpopup-members-office%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
