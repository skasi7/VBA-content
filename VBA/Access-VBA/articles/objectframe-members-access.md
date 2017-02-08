---
title: ObjectFrame Members (Access)
ms.prod: ACCESS
ms.assetid: 65229083-68ec-b870-50f4-a6c329259a39
---


# ObjectFrame Members (Access)


This object corresponds to an unbound object frame. The unbound object frame control displays a picture, chart, or any OLE object not stored in a table.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Click](objectframe-click-event-access.md)|The  **Click** event occurs when the user presses and then releases a mouse button over an object.|
|[DblClick](objectframe-dblclick-event-access.md)|The  **DblClick** event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.|
|[Enter](objectframe-enter-event-access.md)|The  **Enter** event occurs before a control actually receives the focus from a control on the same form or report.|
|[Exit](objectframe-exit-event-access.md)|The  **Exit** event occurs just before a control loses the focus to another control on the same form or report.|
|[GotFocus](objectframe-gotfocus-event-access.md)|The  **GotFocus** event occurs when the specified object receives the focus.|
|[LostFocus](objectframe-lostfocus-event-access.md)|The  **LostFocus** event occurs when the specified object loses the focus.|
|[MouseDown](objectframe-mousedown-event-access.md)|The  **MouseDown** event occurs when the user presses a mouse button.|
|[MouseMove](objectframe-mousemove-event-access.md)|The  **MouseMove** event occurs when the user moves the mouse.|
|[MouseUp](objectframe-mouseup-event-access.md)|The  **MouseUp** event occurs when the user releases a mouse button.|
|[Updated](objectframe-updated-event-access.md)|The  **Updated** event occurs when an OLE object's data has been modified.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Move](objectframe-move-method-access.md)|Moves the specified object to the coordinates specified by the argument values.|
|[Requery](objectframe-requery-method-access.md)|The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.|
|[SetFocus](objectframe-setfocus-method-access.md)|The  **SetFocus** method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.|
|[SizeToFit](objectframe-sizetofit-method-access.md)|You can use the  **SizeToFit** method to size a control so it fits the text or image that it contains.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Action](objectframe-action-property-access.md)|You can use the  **Action** property in Visual Basic to specify the operation to perform on an OLE object. Read/write **Integer**.|
|[Application](objectframe-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.|
|[AutoActivate](objectframe-autoactivate-property-access.md)|You can use the  **AutoActivate** property to specify how the user can activate an OLE object. Read/write **Integer**.|
|[BackColor](objectframe-backcolor-property-access.md)|Gets or sets the interior color of the specified object. Read/write  **Long**.|
|[BackShade](objectframe-backshade-property-access.md)|Gets or sets the shade applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
|[BackStyle](objectframe-backstyle-property-access.md)|You can use the  **BackStyle** property to specify whether a control will be transparent. Read/write **Byte**.|
|[BackThemeColorIndex](objectframe-backthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BackColor** property of the specified object. Read/write **Long**.|
|[BackTint](objectframe-backtint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
|[BorderColor](objectframe-bordercolor-property-access.md)|You can use the  **BorderColor** property to specify the color of a control's border. Read/write **Long**.|
|[BorderShade](objectframe-bordershade-property-access.md)|Gets or sets the shade that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderStyle](objectframe-borderstyle-property-access.md)|Specifies how a control's border appears.Read/write  **Byte**.|
|[BorderThemeColorIndex](objectframe-borderthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BorderColor** property of the specified object. Read/write **Long**.|
|[BorderTint](objectframe-bordertint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderWidth](objectframe-borderwidth-property-access.md)|You can use the  **BorderWidth** property to specify the width of a control's border. Read/write **Byte**.|
|[BottomPadding](objectframe-bottompadding-property-access.md)|Gets or sets the amount of space (in inches) between the object frame and its bottom gridline. Read/write  **Integer**.|
|[Class](objectframe-class-property-access.md)|You can use the  **Class** property to specify or determine the class name of an embeddedOLE object. Read/write **String**.|
|[ColumnCount](objectframe-columncount-property-access.md)|You can use the  **ColumnCount** property to specify the number of columns displayed in a list box or in the list box portion of a combo box, or sent to OLE objects in a chart control or unbound object frame . Read/write **Integer**.|
|[ColumnHeads](objectframe-columnheads-property-access.md)|You can use the  **ColumnHeads** property to display a single row of column headings for list boxes, combo boxes, and OLE objects that accept column headings. You can also use this property to create a label for each entry in a chart control . What is actually displayed as the first-row column heading depends on the object's **RowSourceType** property setting. Read/write **Boolean**.|
|[Controls](objectframe-controls-property-access.md)|Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.|
|[ControlTipText](objectframe-controltiptext-property-access.md)|You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.|
|[ControlType](objectframe-controltype-property-access.md)|You can use the  **ControlType** property in Visual Basic to determine the type of a control on a form or report. Read/write **Byte**.|
|[DisplayType](objectframe-displaytype-property-access.md)|You can use the  **DisplayType** property to specify whether Microsoft Access displays an OLE object's content or an icon. Read/write **Boolean**.|
|[DisplayWhen](objectframe-displaywhen-property-access.md)|You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.|
|[Enabled](objectframe-enabled-property-access.md)|You can use the  **Enabled** property to set or return the status of the conditional format in the **[FormatCondition](formatcondition-object-access.md)** object. Read/write **Boolean**.|
|[EventProcPrefix](objectframe-eventprocprefix-property-access.md)|Gets or sets the prefix portion of an event procedure name. Read/write  **String**.|
|[GridlineColor](objectframe-gridlinecolor-property-access.md)|Gets or sets the color of the gridline for the specified object frame. Read/write  **Long**.|
|[GridlineShade](objectframe-gridlineshade-property-access.md)|Gets or sets the shade applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
|[GridlineStyleBottom](objectframe-gridlinestylebottom-property-access.md)|Gets or sets the bottom gridline style of the specified frame. Read/write  **Byte**.|
|[GridlineStyleLeft](objectframe-gridlinestyleleft-property-access.md)|Gets or sets the width of the bottom gridline for the specified frame. Read/write  **Byte**.|
|[GridlineStyleRight](objectframe-gridlinestyleright-property-access.md)|Gets or sets the right gridline style of the specified frame. Read/write  **Byte**.|
|[GridlineStyleTop](objectframe-gridlinestyletop-property-access.md)|Gets or sets the top gridline style of the specified frame. Read/write  **Byte**.|
|[GridlineThemeColorIndex](objectframe-gridlinethemecolorindex-property-access.md)|Gets or sets the theme color index that represents a color in the applied color theme associated with the  **GridlineColor** property of the specified object. Read/write **Long**.|
|[GridlineTint](objectframe-gridlinetint-property-access.md)|Gets or sets the tint applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
|[GridlineWidthBottom](objectframe-gridlinewidthbottom-property-access.md)|Gets or sets the width of the bottom gridline for the specified frame. Read/write  **Byte**.|
|[GridlineWidthLeft](objectframe-gridlinewidthleft-property-access.md)|Gets or sets the width of the left gridline for the specified frame. Read/write  **Byte**.|
|[GridlineWidthRight](objectframe-gridlinewidthright-property-access.md)|Gets or sets the width of the right gridline for the specified frame. Read/write  **Byte**.|
|[GridlineWidthTop](objectframe-gridlinewidthtop-property-access.md)|Gets or sets the width of the top gridline for the specified frame. Read/write  **Byte**.|
|[Height](objectframe-height-property-access.md)|Gets or sets the height of the specified object in twips. Read/write  **Integer**.|
|[HelpContextId](objectframe-helpcontextid-property-access.md)|The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.|
|[HorizontalAnchor](objectframe-horizontalanchor-property-access.md)|Gets or sets an  **[AcHorizontalAnchor](achorizontalanchor-enumeration-access.md)** constant that indicates how the object frame is anchored horizontally within its layout. Read/write.|
|[InSelection](objectframe-inselection-property-access.md)|You can use the  **InSelection** property to determine or specify whether a control on a form in Design view is selected. Read/write **Boolean**.|
|[IsVisible](objectframe-isvisible-property-access.md)|You can use the  **IsVisible** property in to determine whether a control on a report is visible. Read/write **Boolean**.|
|[Item](objectframe-item-property-access.md)|The  **Item** property returns or sets a specific member of a collection. Read/write **String**.|
|[Layout](objectframe-layout-property-access.md)|Returns the type of layout for the specified object frame. Read-only  **[AcLayoutType](aclayouttype-enumeration-access.md)**.|
|[LayoutID](objectframe-layoutid-property-access.md)|Returns the unique identifier for the layout that contains the specified object frame. Read-only  **Long**.|
|[Left](objectframe-left-property-access.md)|You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.|
|[LeftPadding](objectframe-leftpadding-property-access.md)|Gets or sets the amount of space (in inches) between the object frame and its left gridline. Read/write  **Integer**.|
|[LinkChildFields](objectframe-linkchildfields-property-access.md)|You can use the  **LinkChildFields** property (along with the **LinkMasterFields** property) together to specify how Microsoft Access links records in a form or report to records in a subform, subreport, or embedded object, such as a chart. If these properties are set, Microsoft Access automatically updates the related record in the subform when you change to a new record in a main form. Read/write **String**.|
|[LinkMasterFields](objectframe-linkmasterfields-property-access.md)|You can use the  **LinkMasterFields** property (along with the **LinkChildFields** property) together to specify how Microsoft Access links records in a form or report to records in a subform, subreport, or embedded object, such as a chart. If these properties are set, Microsoft Access automatically updates the related record in the subform when you change to a new record in a main form. Read/write **String**.|
|[Locked](objectframe-locked-property-access.md)|The  **Locked** property specifies whether you can edit data in a control in Form view. Read/write **Boolean**.|
|[Name](objectframe-name-property-access.md)|You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.|
|[Object](objectframe-object-property-access.md)|You can use the  **Object** property in Visual Basic to return a reference to the ActiveX object that is associated with a linked or embedded OLE object in a control. By using this reference, you can access the properties or invoke the methods of the OLE object. Read-only **Object**.|
|[ObjectPalette](objectframe-objectpalette-property-access.md)|The  **ObjectPalette** property specifies the palette in the application used to create an OLE object. Read/write **Variant**.|
|[ObjectVerbs](objectframe-objectverbs-property-access.md)|You can use the  **ObjectVerbs** property in Visual Basic to determine the list of verbs an OLE object supports. Read-only **String**.|
|[ObjectVerbsCount](objectframe-objectverbscount-property-access.md)|You can use the  **ObjectVerbsCount** property in Visual Basic to determine the number of verbs supported by an OLE object. Read-only **Long**.|
|[OldBorderStyle](objectframe-oldborderstyle-property-access.md)|You can use this property to set or returns the unedited value of the  **BorderStyle** property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write **Byte**.|
|[OldValue](objectframe-oldvalue-property-access.md)|You can use the  **OldValue** property to determine the unedited value of a bound control. Read-only **Variant**.|
|[OLEClass](objectframe-oleclass-property-access.md)|You can use the  **OLEClass** property to obtain a description of the kind of OLE object contained in a chart control or an unbound object frame. Read-only **String**.|
|[OLEType](objectframe-oletype-property-access.md)|You can use the  **OLEType** property to determine if a control contains an OLE object, and, if so, whether the object is linked or embedded. Read/write **Byte**.|
|[OLETypeAllowed](objectframe-oletypeallowed-property-access.md)|You can use the  **OLETypeAllowed** property to specify the type of OLE object a control can contain. Read/write **Byte**.|
|[OnClick](objectframe-onclick-property-access.md)|Sets or returns the value of the  **On Click** box in the **Properties** window. Read/write **String**.|
|[OnDblClick](objectframe-ondblclick-property-access.md)|Sets or returns the value of the  **On Dbl Click** box in the **Properties** window. Read/write **String**.|
|[OnEnter](objectframe-onenter-property-access.md)|Sets or returns the value of the  **On Enter** box in the **Properties** window of specified object. Read/write **String**. .|
|[OnExit](objectframe-onexit-property-access.md)|Sets or returns the value of the  **On Exit** box in the **Properties** window of specified object. Read/write **String**. .|
|[OnGotFocus](objectframe-ongotfocus-property-access.md)|Sets or returns the value of the  **On Got Focus** box in the **Properties** window of the specified object. Read/write **String**.|
|[OnLostFocus](objectframe-onlostfocus-property-access.md)|Sets or returns the value of the  **On Lost Focus** box in the **Properties** window of the specified object. Read/write **String**.|
|[OnMouseDown](objectframe-onmousedown-property-access.md)|Sets or returns the value of the  **On Mouse Down** box in the **Properties** window. Read/write **String**.|
|[OnMouseMove](objectframe-onmousemove-property-access.md)|Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.|
|[OnMouseUp](objectframe-onmouseup-property-access.md)|Sets or returns the value of the  **On Mouse Up** box in the **Properties** window. Read/write **String**.|
|[OnUpdated](objectframe-onupdated-property-access.md)|Sets or returns the value of the  **On Updated** box in the **Properties** window of a form or report. Read/write **String**.|
|[Parent](objectframe-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Properties](objectframe-properties-property-access.md)|Returns a reference to a control's **[Properties](properties-object-access.md)** collection object. Read-only.|
|[RightPadding](objectframe-rightpadding-property-access.md)|Gets or sets the amount of space (in inches) between the object frame and its right gridline. Read/write  **Integer**.|
|[RowSource](objectframe-rowsource-property-access.md)|You can use the  **RowSource** property (along with the **RowSourceType** property) to tell Microsoft Access how to provide data tothe specified object. Read/write **String**.|
|[RowSourceType](objectframe-rowsourcetype-property-access.md)|You can use the  **RowSourceType** property (along with the **RowSource** property) to tell Microsoft Access how to provide data tothe specified object. Read/write **String**.|
|[Scaling](objectframe-scaling-property-access.md)|Controls how the contents of an object frame control are displayed. Read/write  **Byte**.|
|[Section](objectframe-section-property-access.md)|You can identify these controls by the section of a form or report where the control appears. Read/write  **Integer**.|
|[ShortcutMenuBar](objectframe-shortcutmenubar-property-access.md)|You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.|
|[SizeMode](objectframe-sizemode-property-access.md)|You can use the  **SizeMode** property to specify how to size a picture or other object in a bound object frame, an unbound object frame, or an image control.|
|[SourceDoc](objectframe-sourcedoc-property-access.md)|You can use the  **SourceDoc** property to specify the file to create a link to or to embed when you create a linked object or embedded object by using the **Action** property in Visual Basic. Read/write **String**.|
|[SourceItem](objectframe-sourceitem-property-access.md)|You can use the  **SourceItem** property to specify the data within a file to be linked when you create a linked OLE object. Read/write **String**.|
|[SourceObject](objectframe-sourceobject-property-access.md)|You can use this property for linked unbound object frames to determine the complete path and file name of the file that contains the data linked to the object frame. Read-only  **String**.|
|[SpecialEffect](objectframe-specialeffect-property-access.md)|You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.|
|[StatusBarText](objectframe-statusbartext-property-access.md)|You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.|
|[TabIndex](objectframe-tabindex-property-access.md)|You can use the  **TabIndex** property to specify a control's place in the tab order on a form or report. Read/write **Integer**.|
|[TabStop](objectframe-tabstop-property-access.md)|You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control. Read/write **Boolean**.|
|[Tag](objectframe-tag-property-access.md)|Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write  **String**.|
|[Top](objectframe-top-property-access.md)|You can use the  **Top** property to specify an object's location on a form or report. Read/write **Integer**. .|
|[TopPadding](objectframe-toppadding-property-access.md)|Gets or sets the amount of space (in inches) between the object frame and its top gridline. Read/write  **Integer**.|
|[UpdateMethod](objectframe-updatemethod-property-access.md)|This property has been deprecated. Use the  **[UpdateOptions](objectframe-updateoptions-property-access.md)** property to specify how a linkedOLE object is updated.|
|[UpdateOptions](objectframe-updateoptions-property-access.md)|You can use the  **UpdateOptions** property to specify how a linkedOLE object is updated. Read/write **Integer**.|
|[VarOleObject](objectframe-varoleobject-property-access.md)| Gets a pointer to an **IOLEObject** that represents the memory address of an OLE object. Read-only **Variant**.|
|[Verb](objectframe-verb-property-access.md)|You can use the  **Verb** property to specify the operation to perform when an OLE object is activated, which is permitted when the control's **Action** property is set to **acOLEActivate**. Read/write **Long**.|
|[VerticalAnchor](objectframe-verticalanchor-property-access.md)|Gets or sets an [AcVerticalAnchor](acverticalanchor-enumeration-access.md) constant that indicates how the specified object frame is anchored vertically within its layout. Read/write.|
|[Visible](objectframe-visible-property-access.md)|Returns or sets whether the object is visible. Read/write  **Boolean**.|
|[Width](objectframe-width-property-access.md)|Gets or sets the width of the specified object in twips. Read/write  **Integer**.|

