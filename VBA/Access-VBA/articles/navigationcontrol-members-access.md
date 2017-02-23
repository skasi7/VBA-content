---
title: NavigationControl Members (Access)
ms.prod: ACCESS
ms.assetid: c972327e-9b46-f9fb-d69d-104d1d130ee4
---


# NavigationControl Members (Access)


This object represents a navigation control on a form.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterUpdate](navigationcontrol-afterupdate-event-access.md)|The  **AfterUpdate** event occurs after changed data in a control or record is updated.|
|[BeforeUpdate](navigationcontrol-beforeupdate-event-access.md)|The  **BeforeUpdate** event occurs before changed data in a control or record is updated.|
|[Change](navigationcontrol-change-event-access.md)|The  **Change** event occurs when the contents of the specified control changes.|
|[Click](navigationcontrol-click-event-access.md)|The  **Click** event occurs when the user presses and then releases a mouse button over an object.|
|[DblClick](navigationcontrol-dblclick-event-access.md)|The  **DblClick** event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.|
|[Dirty](navigationcontrol-dirty-event-access.md)|The Dirty event occurs when the contents of the specified control changes.|
|[Enter](navigationcontrol-enter-event-access.md)|The  **Enter** event occurs before a control actually receives the focus from a control on the same form or report.|
|[Exit](navigationcontrol-exit-event-access.md)|The  **Exit** event occurs just before a control loses the focus to another control on the same form or report.|
|[GotFocus](navigationcontrol-gotfocus-event-access.md)|The  **GotFocus** event occurs when the specified object receives the focus.|
|[KeyDown](navigationcontrol-keydown-event-access.md)|The  **KeyDown** event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyPress](navigationcontrol-keypress-event-access.md)|The  **KeyPress** event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyUp](navigationcontrol-keyup-event-access.md)|The  **KeyUp** event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[LostFocus](navigationcontrol-lostfocus-event-access.md)|The  **LostFocus** event occurs when the specified object loses the focus.|
|[MouseDown](navigationcontrol-mousedown-event-access.md)|The  **MouseDown** event occurs when the user presses a mouse button.|
|[MouseMove](navigationcontrol-mousemove-event-access.md)|The  **MouseMove** event occurs when the user moves the mouse.|
|[MouseUp](navigationcontrol-mouseup-event-access.md)|The  **MouseUp** event occurs when the user releases a mouse button.|
|[Undo](navigationcontrol-undo-event-access.md)|Occurs when the user undoes a change.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Move](navigationcontrol-move-method-access.md)|Moves the specified object to the coordinates specified by the argument values.|
|[Requery](navigationcontrol-requery-method-access.md)|The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.|
|[SetFocus](navigationcontrol-setfocus-method-access.md)|The  **SetFocus** method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.|
|[SizeToFit](navigationcontrol-sizetofit-method-access.md)|You can use the  **SizeToFit** method to size a control so it fits the text or image that it contains.|
|[Undo](navigationcontrol-undo-method-access.md)|You can use the  **Undo** method to reset a control or form when its value has been changed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](navigationcontrol-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.|
|[AutoTab](navigationcontrol-autotab-property-access.md)|You can use the  **AutoTab** property to specify whether an automatic tab occurs when the last character permitted by a text box control's input mask is entered. An automatic tab moves the focus to the next control in the form's tab order. Read/write **Boolean**.|
|[BackColor](navigationcontrol-backcolor-property-access.md)|Gets or sets the interior color of the specified object. Read/write  **Long**.|
|[BackShade](navigationcontrol-backshade-property-access.md)|Gets or sets the shade applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
|[BackStyle](navigationcontrol-backstyle-property-access.md)|You can use the  **BackStyle** property to specify whether a control will be transparent. Read/write **Byte**.|
|[BackThemeColorIndex](navigationcontrol-backthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BackColor** property of the specified object. Read/write **Long**.|
|[BackTint](navigationcontrol-backtint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
|[BorderColor](navigationcontrol-bordercolor-property-access.md)|You can use the  **BorderColor** property to specify the color of a control's border. Read/write **Long**.|
|[BorderShade](navigationcontrol-bordershade-property-access.md)|Gets or sets the shade applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderStyle](navigationcontrol-borderstyle-property-access.md)|Specifies how a control's border appears.Read/write  **Byte**.|
|[BorderThemeColorIndex](navigationcontrol-borderthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BorderColor** property of the specified object. Read/write **Long**.|
|[BorderTint](navigationcontrol-bordertint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderWidth](navigationcontrol-borderwidth-property-access.md)|You can use the  **BorderWidth** property to specify the width of a control's border. Read/write **Byte**.|
|[BottomPadding](navigationcontrol-bottompadding-property-access.md)|Gets or sets the amount of space (in inches) between the list box and its bottom gridline. Read/write  **Integer**.|
|[Controls](navigationcontrol-controls-property-access.md)|Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.|
|[ControlTipText](navigationcontrol-controltiptext-property-access.md)|You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.|
|[ControlType](navigationcontrol-controltype-property-access.md)|You can use the  **ControlType** property in Visual Basic to determine the type of a control on a form or report. Read/write **Byte**.|
|[DisplayWhen](navigationcontrol-displaywhen-property-access.md)|You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.|
|[Enabled](navigationcontrol-enabled-property-access.md)|You can use the  **Enabled** property to set or return the status of the conditional format in the **[FormatCondition](formatcondition-object-access.md)** object. Read/write **Boolean**.|
|[EventProcPrefix](navigationcontrol-eventprocprefix-property-access.md)|Gets or sets the prefix portion of an event procedure name. Read/write  **String**.|
|[FilterLookup](navigationcontrol-filterlookup-property-access.md)|You can use the  **FilterLookup** property to specify whether values appear in a boundtext box control when using the Filter By Form or Server Filter By Form window. Read/write **Byte**.|
|[FormatConditions](navigationcontrol-formatconditions-property-access.md)|You can use the  **FormatConditions** property to return a read-only reference to the **[FormatConditions](formatconditions-object-access.md)** collection and its related properties.|
|[GridlineColor](navigationcontrol-gridlinecolor-property-access.md)|Gets or sets the color of the gridline for the specified list box. Read/write  **Long**.|
|[GridlineShade](navigationcontrol-gridlineshade-property-access.md)|Gets or sets the shade applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
|[GridlineStyleBottom](navigationcontrol-gridlinestylebottom-property-access.md)|Gets or sets the bottom gridline style of the specified list box. Read/write  **Byte**.|
|[GridlineStyleLeft](navigationcontrol-gridlinestyleleft-property-access.md)|Gets or sets the width of the bottom gridline for the specified text box. Read/write  **Byte**.|
|[GridlineStyleRight](navigationcontrol-gridlinestyleright-property-access.md)|Gets or sets the right gridline style of the specified text box. Read/write  **Byte**.|
|[GridlineStyleTop](navigationcontrol-gridlinestyletop-property-access.md)|Gets or sets the top gridline style of the specified text box. Read/write  **Byte**.|
|[GridlineThemeColorIndex](navigationcontrol-gridlinethemecolorindex-property-access.md)|Gets or sets the theme color index that represents a color in the applied color theme associated with the  **GridlineColor** property of the specified object. Read/write **Long**.|
|[GridlineTint](navigationcontrol-gridlinetint-property-access.md)|Gets or sets the tint applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
|[GridlineWidthBottom](navigationcontrol-gridlinewidthbottom-property-access.md)|Gets or sets the width of the bottom gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthLeft](navigationcontrol-gridlinewidthleft-property-access.md)|Gets or sets the width of the left gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthRight](navigationcontrol-gridlinewidthright-property-access.md)|Gets or sets the width of the right gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthTop](navigationcontrol-gridlinewidthtop-property-access.md)|Gets or sets the width of the top gridline for the specified text box. Read/write  **Byte**.|
|[Height](navigationcontrol-height-property-access.md)|Gets or sets the height of the specified object in twips. Read/write  **Integer**.|
|[HelpContextId](navigationcontrol-helpcontextid-property-access.md)|The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.|
|[HorizontalAnchor](navigationcontrol-horizontalanchor-property-access.md)|Gets or sets an  **[AcHorizontalAnchor](achorizontalanchor-enumeration-access.md)** constant that indicates how the text box is anchored horizontally within its layout. Read/write.|
|[Hyperlink](navigationcontrol-hyperlink-property-access.md)|You can use the  **Hyperlink** property to return a reference to a **Hyperlink** object. You can use the **Hyperlink** property to access the properties and methods of a control's hyperlink. Read-only.|
|[InSelection](navigationcontrol-inselection-property-access.md)|You can use the  **InSelection** property to determine or specify whether a control on a form in Design view is selected. Read/write **Boolean**.|
|[IsVisible](navigationcontrol-isvisible-property-access.md)|You can use the  **IsVisible** property in to determine whether a control on a report is visible. Read/write **Boolean**.|
|[KeyboardLanguage](navigationcontrol-keyboardlanguage-property-access.md)||
|[Layout](navigationcontrol-layout-property-access.md)|Returns the type of layout for the specified text box. Read-only  **[AcLayoutType](aclayouttype-enumeration-access.md)**.|
|[LayoutID](navigationcontrol-layoutid-property-access.md)|Returns the unique identifier for the layout that contains the specified text box. Read-only  **Long**.|
|[Left](navigationcontrol-left-property-access.md)|You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.|
|[LeftPadding](navigationcontrol-leftpadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its left gridline. Read/write  **Integer**.|
|[LineSpacing](navigationcontrol-linespacing-property-access.md)|You can use the  **LineSpacing** property to specify or determine the location of information displayed within a label or text box control. Read/write **Integer**.|
|[Name](navigationcontrol-name-property-access.md)|You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.|
|[NumeralShapes](navigationcontrol-numeralshapes-property-access.md)||
|[OldBorderStyle](navigationcontrol-oldborderstyle-property-access.md)|You can use this property to set or returns the unedited value of the  **BorderStyle** property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write **Byte**.|
|[OldValue](navigationcontrol-oldvalue-property-access.md)|You can use the  **OldValue** property to determine the unedited value of a bound control. Read-only **Variant**.|
|[OnClick](navigationcontrol-onclick-property-access.md)|Sets or returns the value of the  **On Click** box in the **Properties** window. Read/write **String**.|
|[OnDblClick](navigationcontrol-ondblclick-property-access.md)|Sets or returns the value of the  **On Dbl Click** box in the **Properties** window. Read/write **String**.|
|[OnGotFocus](navigationcontrol-ongotfocus-property-access.md)|Sets or returns the value of the  **On Got Focus** box in the **Properties** window of the specified object. Read/write **String**.|
|[OnKeyDown](navigationcontrol-onkeydown-property-access.md)|Sets or returns the value of the  **On Key Down** box in the **Properties** window. Read/write **String**.|
|[OnKeyPress](navigationcontrol-onkeypress-property-access.md)|Sets or returns the value of the  **On Key Press** box in the **Properties** window. Read/write **String**.|
|[OnKeyUp](navigationcontrol-onkeyup-property-access.md)|Sets or returns the value of the  **On Key Up** box in the **Properties** window. Read/write **String**.|
|[OnLostFocus](navigationcontrol-onlostfocus-property-access.md)|Sets or returns the value of the  **On Lost Focus** box in the **Properties** window of the specified object. Read/write **String**.|
|[OnMouseDown](navigationcontrol-onmousedown-property-access.md)|Sets or returns the value of the  **On Mouse Down** box in the **Properties** window. Read/write **String**.|
|[OnMouseMove](navigationcontrol-onmousemove-property-access.md)|Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.|
|[OnMouseUp](navigationcontrol-onmouseup-property-access.md)|Sets or returns the value of the  **On Mouse Up** box in the **Properties** window. Read/write **String**.|
|[Parent](navigationcontrol-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Properties](navigationcontrol-properties-property-access.md)|Returns a reference to a control's **[Properties](properties-object-access.md)** collection object. Read-only.|
|[ReadingOrder](navigationcontrol-readingorder-property-access.md)|You can use the  **ReadingOrder** property to specify or determine the reading order of words in text. Read/write **Byte**.|
|[RightPadding](navigationcontrol-rightpadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its right gridline. Read/write  **Integer**.|
|[ScrollBarAlign](navigationcontrol-scrollbaralign-property-access.md)|You can use the  **ScrollBarAlign** to specify or determine the alignment of a vertical scroll bar. Read/write **Byte**.|
|[Section](navigationcontrol-section-property-access.md)|You can identify these controls by the section of a form or report where the control appears. Read/write  **Integer**.|
|[SelectedTab](navigationcontrol-selectedtab-property-access.md)|Gets the active tab of the navigation control. Read-only  **[NavigationButton](navigationbutton-object-access.md)**.|
|[ShortcutMenuBar](navigationcontrol-shortcutmenubar-property-access.md)|You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.|
|[SmartTags](navigationcontrol-smarttags-property-access.md)|Returns a  **[SmartTags](smarttags-object-access.md)** collection that represents the collection of smart tags that have been added to a control. .|
|[Span](navigationcontrol-span-property-access.md)|Gets or sets the orientation of the navigation buttons. Read/write  **[AcNavigationSpan](acnavigationspan-enumeration-access.md)**.|
|[SpecialEffect](navigationcontrol-specialeffect-property-access.md)|You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.|
|[StatusBarText](navigationcontrol-statusbartext-property-access.md)|You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.|
|[SubForm](navigationcontrol-subform-property-access.md)|Gets or sets the name of the  **[SubForm](subform-object-access.md)** object used to display forms. Read/write **String**.|
|[TabIndex](navigationcontrol-tabindex-property-access.md)|You can use the  **TabIndex** property to specify a control's place in the tab order on a form or report. Read/write **Integer**.|
|[Tabs](navigationcontrol-tabs-property-access.md)|Gets the collection of navigation buttons for the specified navigation control. Read-only  **Children**.|
|[TabStop](navigationcontrol-tabstop-property-access.md)|You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control. Read/write **Boolean**.|
|[Tag](navigationcontrol-tag-property-access.md)|Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write  **String**.|
|[Top](navigationcontrol-top-property-access.md)|You can use the  **Top** property to specify an object's location on a form or report. Read/write **Integer**. .|
|[TopPadding](navigationcontrol-toppadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its top gridline. Read/write  **Integer**.|
|[Value](navigationcontrol-value-property-access.md)|Determines or specifies the text in the text box. Read/write  **Variant**.|
|[VerticalAnchor](navigationcontrol-verticalanchor-property-access.md)|Gets or sets an [AcVerticalAnchor](acverticalanchor-enumeration-access.md) constant that indicates how the specified text box is anchored vertically within its layout. Read/write.|
|[Visible](navigationcontrol-visible-property-access.md)|Returns or sets whether the object is visible. Read/write  **Boolean**.|
|[Width](navigationcontrol-width-property-access.md)|Gets or sets the width of the specified object in twips. Read/write  **Integer**.|

