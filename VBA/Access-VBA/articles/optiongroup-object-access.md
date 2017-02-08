---
title: OptionGroup Object (Access)
keywords: vbaac10.chm10894
f1_keywords:
- vbaac10.chm10894
ms.prod: ACCESS
api_name:
- Access.OptionGroup
ms.assetid: aa9e5607-7892-9ab2-dabc-822372b23811
---


# OptionGroup Object (Access)

An option group on a form or report displays a limited set of alternatives. An option group makes selecting a value easy since you can just click the value you want. Only one option in an option group can be selected at a time.


## Remarks

An option group consists of a group frame and a set of check boxes, toggle buttons, or option buttons.

If an option group is bound to a field, only the group frame itself is bound to the field, not the check boxes, toggle buttons, or option buttons inside the frame. Instead of etting the  **ControlSource** property for each control in the option group, you set the **OptionValue** property of each check box, toggle button, or option button to a number that's meaningful for the field to which the group frame is bound. When you select an option in an option group, Microsoft Access sets the value of the field to which the option group is bound to the value of the selected option's **OptionValue** property.




 **Note**  The  **OptionValue** property is set to a number because the value of an option group can only be a number, not text. Microsoft Access stores this number in the underlying table. In the preceding example, if you want to display the name of the shipper instead of a number in the Orders table, you can create a separate table called Shippers that stores shipper names, and then make the ShipVia field in the Orders table a **Lookup** field that looks up data in the Shippers table.

An option group can also be set to an expression, or it can be unbound. You can use an unbound option group in a custom dialog box to accept user input and then carry out an action based on that input.


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/optiongroup-afterupdate-event-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/optiongroup-beforeupdate-event-access%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/optiongroup-click-event-access%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/optiongroup-dblclick-event-access%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/optiongroup-enter-event-access%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/optiongroup-exit-event-access%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/optiongroup-mousedown-event-access%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/optiongroup-mousemove-event-access%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/optiongroup-mouseup-event-access%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/optiongroup-move-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/optiongroup-requery-method-access%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/optiongroup-setfocus-method-access%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/optiongroup-sizetofit-method-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/optiongroup-undo-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/optiongroup-addcolon-property-access%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/optiongroup-afterupdate-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/optiongroup-application-property-access%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/optiongroup-autolabel-property-access%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/optiongroup-backcolor-property-access%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/optiongroup-backshade-property-access%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/optiongroup-backstyle-property-access%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/optiongroup-backthemecolorindex-property-access%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/optiongroup-backtint-property-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/optiongroup-beforeupdate-property-access%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/optiongroup-bordercolor-property-access%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/optiongroup-bordershade-property-access%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/optiongroup-borderstyle-property-access%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/optiongroup-borderthemecolorindex-property-access%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/optiongroup-bordertint-property-access%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/optiongroup-borderwidth-property-access%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/optiongroup-columnhidden-property-access%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/optiongroup-columnorder-property-access%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/optiongroup-columnwidth-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/optiongroup-controls-property-access%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/optiongroup-controlsource-property-access%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/optiongroup-controltiptext-property-access%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/optiongroup-controltype-property-access%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/optiongroup-defaultvalue-property-access%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/optiongroup-displaywhen-property-access%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/optiongroup-enabled-property-access%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/optiongroup-eventprocprefix-property-access%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/optiongroup-height-property-access%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/optiongroup-helpcontextid-property-access%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/optiongroup-hideduplicates-property-access%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/optiongroup-horizontalanchor-property-access%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/optiongroup-inselection-property-access%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/optiongroup-isvisible-property-access%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/optiongroup-labelalign-property-access%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/optiongroup-labelx-property-access%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/optiongroup-labely-property-access%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/optiongroup-left-property-access%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/optiongroup-locked-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/optiongroup-name-property-access%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/optiongroup-oldborderstyle-property-access%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/optiongroup-oldvalue-property-access%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/optiongroup-onclick-property-access%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/optiongroup-ondblclick-property-access%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/optiongroup-onenter-property-access%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/optiongroup-onexit-property-access%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/optiongroup-onmousedown-property-access%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/optiongroup-onmousemove-property-access%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/optiongroup-onmouseup-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/optiongroup-parent-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/optiongroup-properties-property-access%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/optiongroup-section-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/optiongroup-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/optiongroup-specialeffect-property-access%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/optiongroup-statusbartext-property-access%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/optiongroup-tabindex-property-access%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/optiongroup-tabstop-property-access%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/optiongroup-tag-property-access%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/optiongroup-top-property-access%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/optiongroup-validationrule-property-access%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/optiongroup-validationtext-property-access%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/optiongroup-value-property-access%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/optiongroup-verticalanchor-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/optiongroup-visible-property-access%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/optiongroup-width-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[OptionGroup Object Members](http://msdn.microsoft.com/library/optiongroup-members-access%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
