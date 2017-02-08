---
title: OptionGroup Members (Access)
ms.prod: ACCESS
ms.assetid: 90e68eb2-20f2-510c-4332-241eeac27f14
---


# OptionGroup Members (Access)


An option group on a form or report displays a limited set of alternatives. An option group makes selecting a value easy since you can just click the value you want. Only one option in an option group can be selected at a time.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterUpdate](optiongroup-afterupdate-event-access.md)|The  **AfterUpdate** event occurs after changed data in a control or record is updated.|
|[BeforeUpdate](optiongroup-beforeupdate-event-access.md)|The  **BeforeUpdate** event occurs before changed data in a control or record is updated.|
|[Click](optiongroup-click-event-access.md)|The  **Click** event occurs when the user presses and then releases a mouse button over an object.|
|[DblClick](optiongroup-dblclick-event-access.md)|The  **DblClick** event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.|
|[Enter](optiongroup-enter-event-access.md)|The  **Enter** event occurs before a control actually receives the focus from a control on the same form or report.|
|[Exit](optiongroup-exit-event-access.md)|The  **Exit** event occurs just before a control loses the focus to another control on the same form or report.|
|[MouseDown](optiongroup-mousedown-event-access.md)|The  **MouseDown** event occurs when the user presses a mouse button.|
|[MouseMove](optiongroup-mousemove-event-access.md)|The  **MouseMove** event occurs when the user moves the mouse.|
|[MouseUp](optiongroup-mouseup-event-access.md)|The  **MouseUp** event occurs when the user releases a mouse button.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Move](optiongroup-move-method-access.md)|Moves the specified object to the coordinates specified by the argument values.|
|[Requery](optiongroup-requery-method-access.md)|The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.|
|[SetFocus](optiongroup-setfocus-method-access.md)|The  **SetFocus** method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.|
|[SizeToFit](optiongroup-sizetofit-method-access.md)|You can use the  **SizeToFit** method to size a control so it fits the text or image that it contains.|
|[Undo](optiongroup-undo-method-access.md)|You can use the  **Undo** method to reset a control or form when its value has been changed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AddColon](optiongroup-addcolon-property-access.md)|Specifies whether a colon follows the text in labels for new controls. Read/write  **Boolean**.|
|[AfterUpdate](optiongroup-afterupdate-property-access.md)|Returns or sets which macro, event procedure, or user-defined function runs when the  **AfterUpdate** event occurs. Read/write **String**.|
|[Application](optiongroup-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.|
|[AutoLabel](optiongroup-autolabel-property-access.md)|Specifies whether labels are automatically created and attached to new controls. Read/write  **Boolean**.|
|[BackColor](optiongroup-backcolor-property-access.md)|Gets or sets the interior color of the specified object. Read/write  **Long**.|
|[BackShade](optiongroup-backshade-property-access.md)|Gets or sets the shade applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
|[BackStyle](optiongroup-backstyle-property-access.md)|You can use the  **BackStyle** property to specify whether a control will be transparent. Read/write **Byte**.|
|[BackThemeColorIndex](optiongroup-backthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BackColor** property of the specified object. Read/write **Long**.|
|[BackTint](optiongroup-backtint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
|[BeforeUpdate](optiongroup-beforeupdate-property-access.md)|Returns or sets which macro, event procedure, or user-defined function runs when the  **BeforeUpdate** event occurs. Read/write **String**.|
|[BorderColor](optiongroup-bordercolor-property-access.md)|You can use the  **BorderColor** property to specify the color of a control's border. Read/write **Long**.|
|[BorderShade](optiongroup-bordershade-property-access.md)|Gets or sets the shade that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderStyle](optiongroup-borderstyle-property-access.md)|Specifies how a control's border appears.Read/write  **Byte**.|
|[BorderThemeColorIndex](optiongroup-borderthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BorderColor** property of the specified object. Read/write **Long**.|
|[BorderTint](optiongroup-bordertint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderWidth](optiongroup-borderwidth-property-access.md)|You can use the  **BorderWidth** property to specify the width of a control's border. Read/write **Byte**.|
|[ColumnHidden](optiongroup-columnhidden-property-access.md)|You can use the  **ColumnHidden** property to show or hide a specified column in Datasheet view. Read/write **Boolean**.|
|[ColumnOrder](optiongroup-columnorder-property-access.md)|You can use the  **ColumnOrder** property to specify the order of the columns in Datasheet view. Read/write **Integer**.|
|[ColumnWidth](optiongroup-columnwidth-property-access.md)|You can use the  **ColumnWidth** property to specify the width of a column in Datasheet view. Read/write **Integer**.|
|[Controls](optiongroup-controls-property-access.md)|Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.|
|[ControlSource](optiongroup-controlsource-property-access.md)|You can use the  **ControlSource** property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write **String**.|
|[ControlTipText](optiongroup-controltiptext-property-access.md)|You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.|
|[ControlType](optiongroup-controltype-property-access.md)|You can use the  **ControlType** property in Visual Basic to determine the type of a control on a form or report. Read/write **Byte**.|
|[DefaultValue](optiongroup-defaultvalue-property-access.md)|Specifies a value that is automatically entered in a field when a new record is created. For example, in an Addresses table you can set the default value for the City field to New York. When users add a record to the table, they can either accept this value or enter the name of a different city. Read/write  **String**.|
|[DisplayWhen](optiongroup-displaywhen-property-access.md)|You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.|
|[Enabled](optiongroup-enabled-property-access.md)|You can use the  **Enabled** property to set or return the status of the conditional format in the **[FormatCondition](formatcondition-object-access.md)** object. Read/write **Boolean**.|
|[EventProcPrefix](optiongroup-eventprocprefix-property-access.md)|Gets or sets the prefix portion of an event procedure name. Read/write  **String**.|
|[Height](optiongroup-height-property-access.md)|Gets or sets the height of the specified object in twips. Read/write  **Integer**.|
|[HelpContextId](optiongroup-helpcontextid-property-access.md)|The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.|
|[HideDuplicates](optiongroup-hideduplicates-property-access.md)|You can use the  **HideDuplicates** property to hide a control on a report when its value is the same as in the preceding record. Read/write **Boolean**.|
|[HorizontalAnchor](optiongroup-horizontalanchor-property-access.md)|Gets or sets an  **[AcHorizontalAnchor](achorizontalanchor-enumeration-access.md)** constant that indicates how the option group is anchored horizontally within its layout. Read/write.|
|[InSelection](optiongroup-inselection-property-access.md)|You can use the  **InSelection** property to determine or specify whether a control on a form in Design view is selected. Read/write **Boolean**.|
|[IsVisible](optiongroup-isvisible-property-access.md)|You can use the  **IsVisible** property in to determine whether a control on a report is visible. Read/write **Boolean**.|
|[LabelAlign](optiongroup-labelalign-property-access.md)|The property specifies the text alignment within attached labels on new controls. Read/write  **Byte**.|
|[LabelX](optiongroup-labelx-property-access.md)|The  **LabelX** property (along with the **LabelY** property) specifies the placement of the label for a new control. Read/write **Integer**.|
|[LabelY](optiongroup-labely-property-access.md)|The  **LabelY** property (along with the **LabelX** property) specifies the placement of the label for a new control. Read/write **Integer**.|
|[Left](optiongroup-left-property-access.md)|You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.|
|[Locked](optiongroup-locked-property-access.md)|The  **Locked** property specifies whether you can edit data in a control in Form view. Read/write **Boolean**.|
|[Name](optiongroup-name-property-access.md)|You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.|
|[OldBorderStyle](optiongroup-oldborderstyle-property-access.md)|You can use this property to set or returns the unedited value of the  **BorderStyle** property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write **Byte**.|
|[OldValue](optiongroup-oldvalue-property-access.md)|You can use the  **OldValue** property to determine the unedited value of a bound control. Read-only **Variant**.|
|[OnClick](optiongroup-onclick-property-access.md)|Sets or returns the value of the  **On Click** box in the **Properties** window. Read/write **String**.|
|[OnDblClick](optiongroup-ondblclick-property-access.md)|Sets or returns the value of the  **On Dbl Click** box in the **Properties** window. Read/write **String**.|
|[OnEnter](optiongroup-onenter-property-access.md)|Sets or returns the value of the  **On Enter** box in the **Properties** window of specified object. Read/write **String**. .|
|[OnExit](optiongroup-onexit-property-access.md)|Sets or returns the value of the  **On Exit** box in the **Properties** window of specified object. Read/write **String**. .|
|[OnMouseDown](optiongroup-onmousedown-property-access.md)|Sets or returns the value of the  **On Mouse Down** box in the **Properties** window. Read/write **String**.|
|[OnMouseMove](optiongroup-onmousemove-property-access.md)|Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.|
|[OnMouseUp](optiongroup-onmouseup-property-access.md)|Sets or returns the value of the  **On Mouse Up** box in the **Properties** window. Read/write **String**.|
|[Parent](optiongroup-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Properties](optiongroup-properties-property-access.md)|Returns a reference to a control's **[Properties](properties-object-access.md)** collection object. Read-only.|
|[Section](optiongroup-section-property-access.md)|You can identify these controls by the section of a form or report where the control appears. Read/write  **Integer**.|
|[ShortcutMenuBar](optiongroup-shortcutmenubar-property-access.md)|You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.|
|[SpecialEffect](optiongroup-specialeffect-property-access.md)|You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.|
|[StatusBarText](optiongroup-statusbartext-property-access.md)|You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.|
|[TabIndex](optiongroup-tabindex-property-access.md)|You can use the  **TabIndex** property to specify a control's place in the tab order on a form or report. Read/write **Integer**.|
|[TabStop](optiongroup-tabstop-property-access.md)|You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control. Read/write **Boolean**.|
|[Tag](optiongroup-tag-property-access.md)|Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write  **String**.|
|[Top](optiongroup-top-property-access.md)|You can use the  **Top** property to specify an object's location on a form or report. Read/write **Integer**. .|
|[ValidationRule](optiongroup-validationrule-property-access.md)|You can use the  **ValidationRule** property to specify requirements for data entered into a record, field, or control. When data is entered that violates the **ValidationRule** setting, you can use the **ValidationText** property to specify the message to be displayed to the user. Read/write **String**.|
|[ValidationText](optiongroup-validationtext-property-access.md)|Use the  **ValidationText** property to specify a message to be displayed to the user when data is entered that violates a **ValidationRule** setting for a record, field, or control. Read/write **String**.|
|[Value](optiongroup-value-property-access.md)|Gets or sets whether or not the specified option button is selected. Read/write  **Variant**.|
|[VerticalAnchor](optiongroup-verticalanchor-property-access.md)|Gets or sets an [AcVerticalAnchor](acverticalanchor-enumeration-access.md) constant that indicates how the specified option group is anchored vertically within its layout. Read/write.|
|[Visible](optiongroup-visible-property-access.md)|Returns or sets whether the object is visible. Read/write  **Boolean**.|
|[Width](optiongroup-width-property-access.md)|Gets or sets the width of the specified object in twips. Read/write  **Integer**.|

