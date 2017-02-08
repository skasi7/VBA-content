---
title: OlkDateControl Object (Outlook)
keywords: vbaol11.chm1000376
f1_keywords:
- vbaol11.chm1000376
ms.prod: OUTLOOK
api_name:
- Outlook.OlkDateControl
ms.assetid: bd0c6bbe-c348-c748-41fe-0cf7ecebcc1e
---


# OlkDateControl Object (Outlook)

A control that supports the drop-down date picker used in inspectors for task and appointment items to select a date. 


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Date Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor.

The following is an example of the date control at runtime. This control supports Microsoft Windows themes.


![Date](images/olDate_ZA10120280.gif)



This control can bind to any built-in or custom  **DateTime** field. However, the control does not support any date format setting for the field, nor does it support the select range behavior that is available in the appointment inspector.

If the  **[Click](http://msdn.microsoft.com/library/olkdatecontrol-click-event-outlook%28Office.15%29.aspx)** event is implemented but the **[DropButtonClick](http://msdn.microsoft.com/library/olkdatecontrol-dropbuttonclick-event-outlook%28Office.15%29.aspx)** event is not implemented, then clicking the drop button will fire only the **Click** event.

For more information about Outlook controls, see [Controls in a Custom Form](http://msdn.microsoft.com/library/controls-in-a-custom-form%28Office.15%29.aspx). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/olkdatecontrol-afterupdate-event-outlook%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/olkdatecontrol-beforeupdate-event-outlook%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/olkdatecontrol-change-event-outlook%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/olkdatecontrol-click-event-outlook%28Office.15%29.aspx)|
|[DoubleClick](http://msdn.microsoft.com/library/olkdatecontrol-doubleclick-event-outlook%28Office.15%29.aspx)|
|[DropButtonClick](http://msdn.microsoft.com/library/olkdatecontrol-dropbuttonclick-event-outlook%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/olkdatecontrol-enter-event-outlook%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/olkdatecontrol-exit-event-outlook%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/olkdatecontrol-keydown-event-outlook%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/olkdatecontrol-keypress-event-outlook%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/olkdatecontrol-keyup-event-outlook%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/olkdatecontrol-mousedown-event-outlook%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/olkdatecontrol-mousemove-event-outlook%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/olkdatecontrol-mouseup-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[DropDown](http://msdn.microsoft.com/library/olkdatecontrol-dropdown-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AutoSize](http://msdn.microsoft.com/library/olkdatecontrol-autosize-property-outlook%28Office.15%29.aspx)|
|[AutoWordSelect](http://msdn.microsoft.com/library/olkdatecontrol-autowordselect-property-outlook%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/olkdatecontrol-backcolor-property-outlook%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/olkdatecontrol-backstyle-property-outlook%28Office.15%29.aspx)|
|[Date](http://msdn.microsoft.com/library/olkdatecontrol-date-property-outlook%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/olkdatecontrol-enabled-property-outlook%28Office.15%29.aspx)|
|[EnterFieldBehavior](http://msdn.microsoft.com/library/olkdatecontrol-enterfieldbehavior-property-outlook%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/olkdatecontrol-font-property-outlook%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/olkdatecontrol-forecolor-property-outlook%28Office.15%29.aspx)|
|[HideSelection](http://msdn.microsoft.com/library/olkdatecontrol-hideselection-property-outlook%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/olkdatecontrol-locked-property-outlook%28Office.15%29.aspx)|
|[MouseIcon](http://msdn.microsoft.com/library/olkdatecontrol-mouseicon-property-outlook%28Office.15%29.aspx)|
|[MousePointer](http://msdn.microsoft.com/library/olkdatecontrol-mousepointer-property-outlook%28Office.15%29.aspx)|
|[ShowNoneButton](http://msdn.microsoft.com/library/olkdatecontrol-shownonebutton-property-outlook%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/olkdatecontrol-text-property-outlook%28Office.15%29.aspx)|
|[TextAlign](http://msdn.microsoft.com/library/olkdatecontrol-textalign-property-outlook%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/olkdatecontrol-value-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[OlkDateControl Object Members](http://msdn.microsoft.com/library/olkdatecontrol-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
