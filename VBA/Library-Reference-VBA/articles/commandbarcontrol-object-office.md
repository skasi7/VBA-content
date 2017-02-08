---
title: CommandBarControl Object (Office)
keywords: vbaof11.chm5000
f1_keywords:
- vbaof11.chm5000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarControl
ms.assetid: b104ec00-beeb-a927-4b7b-108f4e3164f5
---


# CommandBarControl Object (Office)

Represents a command bar control. The  **CommandBarControl** object is a member of the **CommandBarControls** collection. The properties and methods of the **CommandBarControl** object are all shared by the **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

When writing Visual Basic code to work with custom command bar controls, you use the  **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects. When writing code to work with built-in controls in the container application that cannot be represented by one of those three objects, you use the **CommandBarControl** object. Use **Controls** ( _index_ ), where _index_ is the index number of a control, to return a **CommandBarControl** object. (The **Type** property of the control must be **msoControlLabel**, **msoControlExpandingGrid**, **msoControlSplitExpandingGrid**, **msoControlGrid**, or **msoControlGauge** ). Variables declared as **CommandBarControl** can be assigned **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** values.


## Example

You can also use the  **FindControl** method to return a **CommandBarControl** object. The following example searches for a control of type **msoControlGauge**; if it finds one, it displays the index number of the control and the name of the command bar that contains it. In this example, the variable _lbl_ represents a **CommandBarControl** object.


```
Set lbl = CommandBars.FindControl(Type:= msoControlGauge) 
If lbl Is Nothing Then 
    MsgBox "A control of type msoControlGauge was not found." 
Else 
    MsgBox "Control " &amp; lbl.Index &amp; " on command bar " _ 
        &amp; lbl.Parent.Name &amp; " is type msoControlGauge" 
End If
```


## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/commandbarcontrol-copy-method-office%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/commandbarcontrol-delete-method-office%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/commandbarcontrol-execute-method-office%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/commandbarcontrol-move-method-office%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/commandbarcontrol-reset-method-office%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/commandbarcontrol-setfocus-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/commandbarcontrol-application-property-office%28Office.15%29.aspx)|
|[BeginGroup](http://msdn.microsoft.com/library/commandbarcontrol-begingroup-property-office%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/commandbarcontrol-builtin-property-office%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/commandbarcontrol-caption-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/commandbarcontrol-creator-property-office%28Office.15%29.aspx)|
|[DescriptionText](http://msdn.microsoft.com/library/commandbarcontrol-descriptiontext-property-office%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/commandbarcontrol-enabled-property-office%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/commandbarcontrol-height-property-office%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/commandbarcontrol-helpcontextid-property-office%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/commandbarcontrol-helpfile-property-office%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/commandbarcontrol-id-property-office%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/commandbarcontrol-index-property-office%28Office.15%29.aspx)|
|[IsPriorityDropped](http://msdn.microsoft.com/library/commandbarcontrol-isprioritydropped-property-office%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/commandbarcontrol-left-property-office%28Office.15%29.aspx)|
|[OLEUsage](http://msdn.microsoft.com/library/commandbarcontrol-oleusage-property-office%28Office.15%29.aspx)|
|[OnAction](http://msdn.microsoft.com/library/commandbarcontrol-onaction-property-office%28Office.15%29.aspx)|
|[Parameter](http://msdn.microsoft.com/library/commandbarcontrol-parameter-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/commandbarcontrol-parent-property-office%28Office.15%29.aspx)|
|[Priority](http://msdn.microsoft.com/library/commandbarcontrol-priority-property-office%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/commandbarcontrol-tag-property-office%28Office.15%29.aspx)|
|[TooltipText](http://msdn.microsoft.com/library/commandbarcontrol-tooltiptext-property-office%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/commandbarcontrol-top-property-office%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/commandbarcontrol-type-property-office%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/commandbarcontrol-visible-property-office%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/commandbarcontrol-width-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[CommandBarControl Object Members](http://msdn.microsoft.com/library/commandbarcontrol-members-office%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
