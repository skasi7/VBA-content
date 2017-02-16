---
title: ComboBox Members (Access)
ms.prod: ACCESS
ms.assetid: d0d83ca3-3698-295e-5335-7d0816557d6b
---


# ComboBox Members (Access)


This object corresponds to a combo box control. The combo box control combines the features of a text box and a list box. Use a combo box when you want the option of either typing a value or selecting a value from a predefined list.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterUpdate](combobox-afterupdate-event-access.md)|The  **AfterUpdate** event occurs after changed data in a control or record is updated.|
|[BeforeUpdate](combobox-beforeupdate-event-access.md)|The  **BeforeUpdate** event occurs before changed data in a control or record is updated.|
|[Change](combobox-change-event-access.md)|The  **Change** event occurs when the contents of the specified control changes.|
|[Click](combobox-click-event-access.md)|The  **Click** event occurs when the user presses and then releases a mouse button over an object.|
|[DblClick](combobox-dblclick-event-access.md)|The  **DblClick** event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.|
|[Dirty](combobox-dirty-event-access.md)|The Dirty event occurs when the contents of the specified control changes.|
|[Enter](combobox-enter-event-access.md)|The  **Enter** event occurs before a control actually receives the focus from a control on the same form or report.|
|[Exit](combobox-exit-event-access.md)|The  **Exit** event occurs just before a control loses the focus to another control on the same form or report.|
|[GotFocus](combobox-gotfocus-event-access.md)|The  **GotFocus** event occurs when the specified object receives the focus.|
|[KeyDown](combobox-keydown-event-access.md)|The  **KeyDown** event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyPress](combobox-keypress-event-access.md)|The  **KeyPress** event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyUp](combobox-keyup-event-access.md)|The  **KeyUp** event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[LostFocus](combobox-lostfocus-event-access.md)|The  **LostFocus** event occurs when the specified object loses the focus.|
|[MouseDown](combobox-mousedown-event-access.md)|The  **MouseDown** event occurs when the user presses a mouse button.|
|[MouseMove](combobox-mousemove-event-access.md)|The  **MouseMove** event occurs when the user moves the mouse.|
|[MouseUp](combobox-mouseup-event-access.md)|The  **MouseUp** event occurs when the user releases a mouse button.|
|[NotInList](combobox-notinlist-event-access.md)|The  **NotInList** event occurs when the user enters a value in the text box portion of a combo box that isn't in the combo box list.|
|[Undo](combobox-undo-event-access.md)|Occurs when the user undoes a change.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddItem](combobox-additem-method-access.md)|Adds a new item to the list of values displayed by the specified combo box control.|
|[Dropdown](combobox-dropdown-method-access.md)|You can use the  **Dropdown** method to force the list in the specified combo box to drop down.|
|[Move](combobox-move-method-access.md)|Moves the specified object to the coordinates specified by the argument values.|
|[RemoveItem](combobox-removeitem-method-access.md)|Removes an item from the list of values displayed by the specified combo box control.|
|[Requery](combobox-requery-method-access.md)|The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.|
|[SetFocus](combobox-setfocus-method-access.md)|The  **SetFocus** method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.|
|[SizeToFit](combobox-sizetofit-method-access.md)|You can use the  **SizeToFit** method to size a control so it fits the text or image that it contains.|
|[Undo](combobox-undo-method-access.md)|You can use the  **Undo** method to reset a control or form when its value has been changed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AddColon](combobox-addcolon-property-access.md)|Specifies whether a colon follows the text in labels for new controls. Read/write  **Boolean**.|
|[AfterUpdate](combobox-afterupdate-property-access.md)|Returns or sets which macro, event procedure, or user-defined function runs when the  **AfterUpdate** event occurs. Read/write **String**.|
|[AllowAutoCorrect](combobox-allowautocorrect-property-access.md)|You can use the  **AllowAutoCorrect** property to specify whetherthe specified control will automatically correct entries made by the user. Read/write **Boolean**.|
|[AllowValueListEdits](combobox-allowvaluelistedits-property-access.md)|Gets or sets whether the  **Edit List Items** command is available when the user right-clicks a combo box. Read/write **Boolean**.|
|[Application](combobox-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.|
|[AutoExpand](combobox-autoexpand-property-access.md)|You can use the  **AutoExpand** property to specify whether Microsoft Access automatically fills the text box portion of a combo box with a value from the combo box list that matches the characters you enter as you type in the combo box. This lets you quickly enter an existing value in a combo box without displaying the list box portion of the combo box. Read/write **Boolean**.|
|[AutoLabel](combobox-autolabel-property-access.md)|Specifies whether labels are automatically created and attached to new controls. Read/write  **Boolean**.|
|[BackColor](combobox-backcolor-property-access.md)|Gets or sets the interior color of the specified object. Read/write  **Long**.|
|[BackShade](combobox-backshade-property-access.md)|Gets or sets the shade that is applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
|[BackStyle](combobox-backstyle-property-access.md)|You can use the  **BackStyle** property to specify whether a control will be transparent. Read/write **Byte**.|
|[BackThemeColorIndex](combobox-backthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BackColor** property of the specified object. Read/write **Long**.|
|[BackTint](combobox-backtint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
|[BeforeUpdate](combobox-beforeupdate-property-access.md)|Returns or sets which macro, event procedure, or user-defined function runs when the  **BeforeUpdate** event occurs. Read/write **String**.|
|[BorderColor](combobox-bordercolor-property-access.md)|You can use the  **BorderColor** property to specify the color of a control's border. Read/write **Long**.|
|[BorderShade](combobox-bordershade-property-access.md)|Gets or sets the shade that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderStyle](combobox-borderstyle-property-access.md)|Specifies how a control's border appears.Read/write  **Byte**.|
|[BorderThemeColorIndex](combobox-borderthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BorderColor** property of the specified object. Read/write **Long**.|
|[BorderTint](combobox-bordertint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderWidth](combobox-borderwidth-property-access.md)|You can use the  **BorderWidth** property to specify the width of a control's border. Read/write **Byte**.|
|[BottomMargin](combobox-bottommargin-property-access.md)|Along with the  **LeftMargin**, **RightMargin**, and **TopMargin** properties, specifies the location of information displayed within a text box control. Read/write **Integer**.|
|[BottomPadding](combobox-bottompadding-property-access.md)|Gets or sets the amount of space (in inches) between the combo box and its bottom gridline. Read/write  **Integer**.|
|[BoundColumn](combobox-boundcolumn-property-access.md)|When you make a selection from a combo box, the  **BoundColumn** property tells Microsoft Access which column's values to use as the value of the control. If the control is bound to a field, the value in the column specified by the **BoundColumn** property is stored in the field named in the **ControlSource** property. Read/write **Long**.|
|[CanGrow](combobox-cangrow-property-access.md)|Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data the control contains. Read/write  **Boolean**.|
|[CanShrink](combobox-canshrink-property-access.md)|Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data the control contains. Read/write  **Boolean**.|
|[Column](combobox-column-property-access.md)|You can use the  **Column** property to refer to a specific column, or column and row combination, in a multiple-column combo box or list box. Read-only **Variant**.|
|[ColumnCount](combobox-columncount-property-access.md)|You can use the  **ColumnCount** property to specify the number of columns displayed in a list box or in the list box portion of a combo box, or sent to OLE objects in a chart control or unbound object frame . Read/write **Integer**.|
|[ColumnHeads](combobox-columnheads-property-access.md)|You can use the  **ColumnHeads** property to display a single row of column headings for list boxes, combo boxes, and OLE objects that accept column headings. You can also use this property to create a label for each entry in a chart control . What is actually displayed as the first-row column heading depends on the object's **RowSourceType** property setting. Read/write **Boolean**.|
|[ColumnHidden](combobox-columnhidden-property-access.md)|You can use the  **ColumnHidden** property to show or hide a specified column in Datasheet view. Read/write **Boolean**.|
|[ColumnOrder](combobox-columnorder-property-access.md)|You can use the  **ColumnOrder** property to specify the order of the columns in Datasheet view. Read/write **Integer**.|
|[ColumnWidth](combobox-columnwidth-property-access.md)|You can use the  **ColumnWidth** property to specify the width of a column in Datasheet view. Read/write **Integer**.|
|[ColumnWidths](combobox-columnwidths-property-access.md)|You can use the  **ColumnWidths** property to specify the width of each column in a multiple-column combo box. Read/write **String**.|
|[Controls](combobox-controls-property-access.md)|Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.|
|[ControlSource](combobox-controlsource-property-access.md)|You can use the  **ControlSource** property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write **String**.|
|[ControlTipText](combobox-controltiptext-property-access.md)|You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.|
|[ControlType](combobox-controltype-property-access.md)|You can use the  **ControlType** property in Visual Basic to determine the type of a control on a form or report. Read/write **Byte**.|
|[DecimalPlaces](combobox-decimalplaces-property-access.md)|You can use the  **DecimalPlaces** property to specify the number of decimal places Microsoft Access uses to display numbers. Read/write **Byte**.|
|[DefaultValue](combobox-defaultvalue-property-access.md)|Specifies a value that is automatically entered in a field when a new record is created. For example, in an Addresses table you can set the default value for the City field to New York. When users add a record to the table, they can either accept this value or enter the name of a different city. Read/write  **String**.|
|[DisplayAsHyperlink](combobox-displayashyperlink-property-access.md)|Gets or sets an  **[AcDisplayAsHyperlink](acdisplayashyperlink-enumeration-access.md)** constant that specifies whether to display the contents of the specified combo box as a hyperlink. Read/write.|
|[DisplayWhen](combobox-displaywhen-property-access.md)|You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.|
|[Enabled](combobox-enabled-property-access.md)|You can use the  **Enabled** property to set or return the status of the conditional format in the **[FormatCondition](formatcondition-object-access.md)** object. Read/write **Boolean**.|
|[EventProcPrefix](combobox-eventprocprefix-property-access.md)|Gets or sets the prefix portion of an event procedure name. Read/write  **String**.|
|[FontBold](combobox-fontbold-property-access.md)|You can use the  **FontBold** property to specify whether a font appears in a bold style in the following situations:|
|[FontItalic](combobox-fontitalic-property-access.md)|You can use the  **FontItalic** property to specify whether text is italic in the following situations:|
|[FontName](combobox-fontname-property-access.md)|You can use the  **FontName** property to specify the font for text in the following situations:|
|[FontSize](combobox-fontsize-property-access.md)|You can use the  **FontSize** property to specify the point size for text in the following situations:|
|[FontUnderline](combobox-fontunderline-property-access.md)|You can use the  **FontUnderline** property to specify whether text is underlined in the following situations:|
|[FontWeight](combobox-fontweight-property-access.md)|You can use the  **DatasheetFontWeight** property to specify the line width of the font used to display and print characters for field names and data in Datasheet view. Read/write **Integer**.|
|[ForeColor](combobox-forecolor-property-access.md)|You can use the  **ForeColor** property to specify the color for text in a control. Read/write **Long**.|
|[ForeShade](combobox-foreshade-property-access.md)|Gets or sets the shade that is applied to the theme color in the  **ForeColor** property of the specified object. Read/write **Single**.|
|[ForeThemeColorIndex](combobox-forethemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **ForeColor** property of the specified object. Read/write **Long**.|
|[ForeTint](combobox-foretint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **ForeColor** property of the specified object. Read/write **Single**.|
|[Format](combobox-format-property-access.md)|You can use the  **Format** property to customize the way numbers, dates, times, and text are displayed and printed. Read/write **String**.|
|[FormatConditions](combobox-formatconditions-property-access.md)|You can use the  **FormatConditions** property to return a read-only reference to the **[FormatConditions](formatconditions-object-access.md)** collection and its related properties.|
|[GridlineColor](combobox-gridlinecolor-property-access.md)|Gets or sets the color of the gridline for the specified combo box. Read/write  **Long**.|
|[GridlineShade](combobox-gridlineshade-property-access.md)|Gets or sets the shade applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
|[GridlineStyleBottom](combobox-gridlinestylebottom-property-access.md)|Gets or sets the bottom gridline style of the specified combo box. Read/write  **Byte**.|
|[GridlineStyleLeft](combobox-gridlinestyleleft-property-access.md)|Gets or sets the width of the bottom gridline for the specified combo box. Read/write  **Byte**.|
|[GridlineStyleRight](combobox-gridlinestyleright-property-access.md)|Gets or sets the right gridline style of the specified combo box. Read/write  **Byte**.|
|[GridlineStyleTop](combobox-gridlinestyletop-property-access.md)|Gets or sets the top gridline style of the specified combo box. Read/write  **Byte**.|
|[GridlineThemeColorIndex](combobox-gridlinethemecolorindex-property-access.md)|Gets or sets the theme color index that represents a color in the applied color theme associated with the  **GridlineColor** property of the specified object. Read/write **Long**.|
|[GridlineTint](combobox-gridlinetint-property-access.md)|Gets or sets the tint applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
|[GridlineWidthBottom](combobox-gridlinewidthbottom-property-access.md)|Gets or sets the width of the bottom gridline for the specified combo box. Read/write  **Byte**.|
|[GridlineWidthLeft](combobox-gridlinewidthleft-property-access.md)|Gets or sets the width of the left gridline for the specified combo box. Read/write  **Byte**.|
|[GridlineWidthRight](combobox-gridlinewidthright-property-access.md)|Gets or sets the width of the right gridline for the specified combo box. Read/write  **Byte**.|
|[GridlineWidthTop](combobox-gridlinewidthtop-property-access.md)|Gets or sets the width of the top gridline for the specified combo box. Read/write  **Byte**.|
|[Height](combobox-height-property-access.md)|Gets or sets the height of the specified object in twips. Read/write  **Integer**.|
|[HelpContextId](combobox-helpcontextid-property-access.md)|The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.|
|[HideDuplicates](combobox-hideduplicates-property-access.md)|You can use the  **HideDuplicates** property to hide a control on a report when its value is the same as in the preceding record. Read/write **Boolean**.|
|[HorizontalAnchor](combobox-horizontalanchor-property-access.md)|Gets or sets an  **[AcHorizontalAnchor](achorizontalanchor-enumeration-access.md)** constant that indicates how the combo box is anchored horizontally within its layout. Read/write.|
|[Hyperlink](combobox-hyperlink-property-access.md)|You can use the  **Hyperlink** property to return a reference to a **Hyperlink** object. You can use the **Hyperlink** property to access the properties and methods of a control's hyperlink. Read-only.|
|[IMEHold](combobox-imehold-property-access.md)|[Language-specific information](learn-about-language-specific-information-access.md)You can use the  **IMEHold/Hold KanjiConversionMode** property to show whether the Kanji Conversion Mode is maintained when the control loses the focus. Read/write **Boolean**.|
|[IMEMode](combobox-imemode-property-access.md)||
|[IMESentenceMode](combobox-imesentencemode-property-access.md)||
|[InheritValueList](combobox-inheritvaluelist-property-access.md)|Gets or sets whether a combo box's value list is inherited from its field. Read/write  **Boolean**.|
|[InputMask](combobox-inputmask-property-access.md)|You can use the  **InputMask** property to make data entry easier and to control the values users can enter in a combo box control. Read/write **String**.|
|[InSelection](combobox-inselection-property-access.md)|You can use the  **InSelection** property to determine or specify whether a control on a form in Design view is selected. Read/write **Boolean**.|
|[IsHyperlink](combobox-ishyperlink-property-access.md)|You can use the  **IsHyperlink** property to specify or determine if the data contained in a combo box is a hyperlink. Read/write **Boolean**.|
|[IsVisible](combobox-isvisible-property-access.md)|You can use the  **IsVisible** property in to determine whether a control on a report is visible. Read/write **Boolean**.|
|[ItemData](combobox-itemdata-property-access.md)|The  **ItemData** property returns the data in the bound column for the specified row in a combo box. Read-only **Variant**.|
|[ItemsSelected](combobox-itemsselected-property-access.md)|You can use the  **ItemsSelected** property to return a read-only reference to the hidden **ItemsSelected** collection. This hidden collection can be used to access data in the selected rows of a multiselect combo box control.|
|[KeyboardLanguage](combobox-keyboardlanguage-property-access.md)||
|[LabelAlign](combobox-labelalign-property-access.md)|The property specifies the text alignment within attached labels on new controls. Read/write  **Byte**.|
|[LabelX](combobox-labelx-property-access.md)|The  **LabelX** property (along with the **LabelY** property) specifies the placement of the label for a new control. Read/write **Integer**.|
|[LabelY](combobox-labely-property-access.md)|The  **LabelY** property (along with the **LabelX** property) specifies the placement of the label for a new control. Read/write **Integer**.|
|[Layout](combobox-layout-property-access.md)|Returns the type of layout for the specified combo box. Read-only  **[AcLayoutType](aclayouttype-enumeration-access.md)**.|
|[LayoutID](combobox-layoutid-property-access.md)|Returns the unique identifier for the layout that contains the specified combo box. Read-only  **Long**.|
|[Left](combobox-left-property-access.md)|You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.|
|[LeftMargin](combobox-leftmargin-property-access.md)|Along with the  **TopMargin**, **RightMargin**, and **BottomMargin** properties, specifies the location of information displayed within a text box control. Read/write **Integer**. .|
|[LeftPadding](combobox-leftpadding-property-access.md)|Gets or sets the amount of space (in inches) between the combo box and its left gridline. Read/write  **Integer**.|
|[LimitToList](combobox-limittolist-property-access.md)|You can use the  **LimitToList** property to limit a combo box's values to the listed items. Read/write **Boolean**.|
|[ListCount](combobox-listcount-property-access.md)|You can use the  **ListCount** property to determine the number of rows in the list box portion of a combo box. Read/write **Long**.|
|[ListIndex](combobox-listindex-property-access.md)|You can use the  **ListIndex** property to determine which item is selected in a combo box. Read/write **Long**.|
|[ListItemsEditForm](combobox-listitemseditform-property-access.md)|Gets or sets the name of the form that is displayed when the user clicks  **Edit List Items**. Read/write  **String**.|
|[ListRows](combobox-listrows-property-access.md)|You can use the  **ListRows** property to set the maximum number of rows to display in the list box portion of a combo box. Read/write **Integer**.|
|[ListWidth](combobox-listwidth-property-access.md)|You can use the  **ListWidth** property to set the width of the list box portion of a combo box. Read/write **String**.|
|[Locked](combobox-locked-property-access.md)|The  **Locked** property specifies whether you can edit data in a control in Form view. Read/write **Boolean**.|
|[Name](combobox-name-property-access.md)|You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.|
|[NumeralShapes](combobox-numeralshapes-property-access.md)||
|[OldBorderStyle](combobox-oldborderstyle-property-access.md)|You can use this property to set or returns the unedited value of the  **BorderStyle** property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write **Byte**.|
|[OldValue](combobox-oldvalue-property-access.md)|You can use the  **OldValue** property to determine the unedited value of a bound control. Read-only **Variant**.|
|[OnChange](combobox-onchange-property-access.md)|Sets or returns the value of the  **On Change** box in the **Properties** window of one of the objects in the Applies To list. Read/write **String**.|
|[OnClick](combobox-onclick-property-access.md)|Sets or returns the value of the  **On Click** box in the **Properties** window. Read/write **String**.|
|[OnDblClick](combobox-ondblclick-property-access.md)|Sets or returns the value of the  **On Dbl Click** box in the **Properties** window. Read/write **String**.|
|[OnDirty](combobox-ondirty-property-access.md)|Sets or returns the value of the  **On Dirty** box in the **Properties** window of a form or report. Read/write **String**.|
|[OnEnter](combobox-onenter-property-access.md)|Sets or returns the value of the  **On Enter** box in the **Properties** window of specified object. Read/write **String**. .|
|[OnExit](combobox-onexit-property-access.md)|Sets or returns the value of the  **On Exit** box in the **Properties** window of specified object. Read/write **String**. .|
|[OnGotFocus](combobox-ongotfocus-property-access.md)|Sets or returns the value of the  **On Got Focus** box in the **Properties** window of the specified object. Read/write **String**.|
|[OnKeyDown](combobox-onkeydown-property-access.md)|Sets or returns the value of the  **On Key Down** box in the **Properties** window. Read/write **String**.|
|[OnKeyPress](combobox-onkeypress-property-access.md)|Sets or returns the value of the  **On Key Press** box in the **Properties** window. Read/write **String**.|
|[OnKeyUp](combobox-onkeyup-property-access.md)|Sets or returns the value of the  **On Key Up** box in the **Properties** window. Read/write **String**.|
|[OnLostFocus](combobox-onlostfocus-property-access.md)|Sets or returns the value of the  **On Lost Focus** box in the **Properties** window of the specified object. Read/write **String**.|
|[OnMouseDown](combobox-onmousedown-property-access.md)|Sets or returns the value of the  **On Mouse Down** box in the **Properties** window. Read/write **String**.|
|[OnMouseMove](combobox-onmousemove-property-access.md)|Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.|
|[OnMouseUp](combobox-onmouseup-property-access.md)|Sets or returns the value of the  **On Mouse Up** box in the **Properties** window. Read/write **String**.|
|[OnNotInList](combobox-onnotinlist-property-access.md)|Sets or returns the value of the  **On Not in List** box in the **Properties** window of a combo box. Read/write **String**.|
|[OnUndo](combobox-onundo-property-access.md)|Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **Undo** event occurs. Read/write..|
|[Parent](combobox-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Properties](combobox-properties-property-access.md)|Returns a reference to a control's **[Properties](properties-object-access.md)** collection object. Read-only.|
|[ReadingOrder](combobox-readingorder-property-access.md)|You can use the  **ReadingOrder** property to specify or determine the reading order of words in text. Read/write **Byte**.|
|[Recordset](combobox-recordset-property-access.md)|Returns or sets the ADO  **Recordset** or DAO **[Recordset](recordset-object-dao.md)** object representing the record source for the specified object. Read/write **Object**.|
|[RightMargin](combobox-rightmargin-property-access.md)|Along with the  **TopMargin**, **Left Margin**, and **BottomMargin** properties, specifies the location of information displayed within a combo box control. Read/write **Integer**.|
|[RightPadding](combobox-rightpadding-property-access.md)|Gets or sets the amount of space (in inches) between the combo box and its right gridline. Read/write  **Integer**.|
|[RowSource](combobox-rowsource-property-access.md)|You can use the  **RowSource** property (along with the **RowSourceType** property) to tell Microsoft Access how to provide data tothe specified object. Read/write **String**.|
|[RowSourceType](combobox-rowsourcetype-property-access.md)|You can use the  **RowSourceType** property (along with the **RowSource** property) to tell Microsoft Access how to provide data tothe specified object. Read/write **String**.|
|[ScrollBarAlign](combobox-scrollbaralign-property-access.md)|You can use the  **ScrollBarAlign** to specify or determine the alignment of a vertical scroll bar. Read/write **Byte**.|
|[Section](combobox-section-property-access.md)|You can identify these controls by the section of a form or report where the control appears. Read/write  **Integer**.|
|[Selected](combobox-selected-property-access.md)|You can use the  **Selected** property in Visual Basic to determine if an item in a combo box is selected. Read/write **Long**.|
|[SelLength](combobox-sellength-property-access.md)|The  **SelLength** property specifies or determines the number of characters selected in the text box portion of a combo box. Read/write **Integer**.|
|[SelStart](combobox-selstart-property-access.md)|The  **SelStart** property specifies or determines the starting point of the selected text or the position of the insertion point if no text is selected. Read/write **Integer**.|
|[SelText](combobox-seltext-property-access.md)|The  **SelText** property returns a string containing the selected text. Read/write **String**.|
|[SeparatorCharacters](combobox-separatorcharacters-property-access.md)|Gets or sets the separator displayed between values when the combo box is bound to a multi-valued field. Read/write [AcSeparatorCharacters](acseparatorcharacters-enumeration-access.md).|
|[ShortcutMenuBar](combobox-shortcutmenubar-property-access.md)|You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.|
|[ShowOnlyRowSourceValues](combobox-showonlyrowsourcevalues-property-access.md)|Gets or sets whether the combo box can display values that aren't specified by the  **RowSource** property. Read/write **Boolean**.|
|[SmartTags](combobox-smarttags-property-access.md)|Returns a  **[SmartTags](smarttags-object-access.md)** collection that represents the collection of smart tags that have been added to a control. .|
|[SpecialEffect](combobox-specialeffect-property-access.md)|You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.|
|[StatusBarText](combobox-statusbartext-property-access.md)|You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.|
|[TabIndex](combobox-tabindex-property-access.md)|You can use the  **TabIndex** property to specify a control's place in the tab order on a form or report. Read/write **Integer**.|
|[TabStop](combobox-tabstop-property-access.md)|You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control. Read/write **Boolean**.|
|[Tag](combobox-tag-property-access.md)|Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write  **String**.|
|[Text](combobox-text-property-access.md)|You can use the  **Text** property to set or return the text contained in the text box portion of a combo box. Read/write **String**.|
|[TextAlign](combobox-textalign-property-access.md)|The  **TextAlign** property specifies the text alignment in new controls. Read/write **Byte**.|
|[ThemeFontIndex](combobox-themefontindex-property-access.md)|Gets or sets the font index that represents a font in the applied theme associated with the  **FontName** property of the specified object. Read/write **Long**.|
|[Top](combobox-top-property-access.md)|You can use the  **Top** property to specify an object's location on a form or report. Read/write **Integer**. .|
|[TopMargin](combobox-topmargin-property-access.md)|Along with the  **LeftMargin**, **RightMargin**, and **BottomMargin** properties, specifies the location of information displayed within a text box control. Read/write **Integer**.|
|[TopPadding](combobox-toppadding-property-access.md)|Gets or sets the amount of space (in inches) between the combo box and its top gridline. Read/write  **Integer**.|
|[ValidationRule](combobox-validationrule-property-access.md)|You can use the  **ValidationRule** property to specify requirements for data entered into a record, field, or control. When data is entered that violates the **ValidationRule** setting, you can use the **ValidationText** property to specify the message to be displayed to the user. Read/write **String**.|
|[ValidationText](combobox-validationtext-property-access.md)|Use the  **ValidationText** property to specify a message to be displayed to the user when data is entered that violates a **ValidationRule** setting for a record, field, or control. Read/write **String**.|
|[Value](combobox-value-property-access.md)|Determines or specifies which value or option in the combo box is selected. Read/write  **Variant**.|
|[VerticalAnchor](combobox-verticalanchor-property-access.md)|Gets or sets an [AcVerticalAnchor](acverticalanchor-enumeration-access.md) constant that indicates how the specified combo box is anchored vertically within its layout. Read/write.|
|[Visible](combobox-visible-property-access.md)|Returns or sets whether the object is visible. Read/write  **Boolean**.|
|[Width](combobox-width-property-access.md)|Gets or sets the width of the specified object in twips. Read/write  **Integer**.|

