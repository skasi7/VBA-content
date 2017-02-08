---
title: CustomControl Members (Access)
ms.prod: ACCESS
ms.assetid: 3093550b-7994-fb58-044c-90e8da535f9d
---


# CustomControl Members (Access)


When setting the properties of an ActiveX control, you may need or prefer to use the control's custom properties dialog box. This custom properties dialog box provides an alternative to the list of properties in the Microsoft Access property sheet for setting ActiveX control properties in Design view.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Enter](customcontrol-enter-event-access.md)|The  **Enter** event occurs before a control actually receives the focus from a control on the same form or report.|
|[Exit](customcontrol-exit-event-access.md)|The  **Exit** event occurs just before a control loses the focus to another control on the same form or report.|
|[GotFocus](customcontrol-gotfocus-event-access.md)|The  **GotFocus** event occurs when the specified object receives the focus.|
|[LostFocus](customcontrol-lostfocus-event-access.md)|The  **LostFocus** event occurs when the specified object loses the focus.|
|[Updated](customcontrol-updated-event-access.md)|The  **Updated** event occurs when an OLE object's data has been modified.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Move](customcontrol-move-method-access.md)|Moves the specified object to the coordinates specified by the argument values.|
|[Requery](customcontrol-requery-method-access.md)|The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.|
|[SetFocus](customcontrol-setfocus-method-access.md)|The  **SetFocus** method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.|
|[SizeToFit](customcontrol-sizetofit-method-access.md)|You can use the  **SizeToFit** method to size a control so it fits the text or image that it contains.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[About](customcontrol-about-property-access.md)|Returns or sets a  **String** representing version and copyright information for an ActiveX control. Read/write.|
|[Application](customcontrol-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.|
|[BorderColor](customcontrol-bordercolor-property-access.md)|You can use the  **BorderColor** property to specify the color of a control's border. Read/write **Long**.|
|[BorderShade](customcontrol-bordershade-property-access.md)|Gets or sets the shade that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderStyle](customcontrol-borderstyle-property-access.md)|Specifies how a control's border appears.Read/write  **Byte**.|
|[BorderThemeColorIndex](customcontrol-borderthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BorderColor** property of the specified object. Read/write **Long**.|
|[BorderTint](customcontrol-bordertint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderWidth](customcontrol-borderwidth-property-access.md)|You can use the  **BorderWidth** property to specify the width of a control's border. Read/write **Byte**.|
|[BottomPadding](customcontrol-bottompadding-property-access.md)|Gets or sets the amount of space (in inches) between the list box and its bottom gridline. Read/write  **Integer**.|
|[Cancel](customcontrol-cancel-property-access.md)|You can use the  **Cancel** property to specify whether a command button is also the Cancel button on a form. Read/write **Boolean**.|
|[Class](customcontrol-class-property-access.md)|You can use the  **Class** property to specify or determine the class name of an embeddedOLE object. Read/write **String**.|
|[Controls](customcontrol-controls-property-access.md)|Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.|
|[ControlSource](customcontrol-controlsource-property-access.md)|You can use the  **ControlSource** property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write **String**.|
|[ControlTipText](customcontrol-controltiptext-property-access.md)|You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.|
|[ControlType](customcontrol-controltype-property-access.md)|You can use the  **ControlType** property in Visual Basic to determine the type of a control on a form or report. Read/write **Byte**.|
|[Custom](customcontrol-custom-property-access.md)|Returns or sets a  **String** representing the custom properties dialog box for an ActiveX control. Read/write.|
|[Default](customcontrol-default-property-access.md)|You can use the  **Default** property to specify whether a command button is the default button on a form. Read/write **Boolean**.|
|[DisplayWhen](customcontrol-displaywhen-property-access.md)|You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.|
|[Enabled](customcontrol-enabled-property-access.md)|You can use the  **Enabled** property to set or return the status of the conditional format in the **[FormatCondition](formatcondition-object-access.md)** object. Read/write **Boolean**.|
|[EventProcPrefix](customcontrol-eventprocprefix-property-access.md)|Gets or sets the prefix portion of an event procedure name. Read/write  **String**.|
|[GridlineColor](customcontrol-gridlinecolor-property-access.md)|Gets or sets the color of the gridline for the specified list box. Read/write  **Long**.|
|[GridlineStyleBottom](customcontrol-gridlinestylebottom-property-access.md)|Gets or sets the bottom gridline style of the specified list box. Read/write  **Byte**.|
|[GridlineStyleLeft](customcontrol-gridlinestyleleft-property-access.md)|Gets or sets the width of the bottom gridline for the specified text box. Read/write  **Byte**.|
|[GridlineStyleRight](customcontrol-gridlinestyleright-property-access.md)|Gets or sets the right gridline style of the specified text box. Read/write  **Byte**.|
|[GridlineStyleTop](customcontrol-gridlinestyletop-property-access.md)|Gets or sets the top gridline style of the specified text box. Read/write  **Byte**.|
|[GridlineWidthBottom](customcontrol-gridlinewidthbottom-property-access.md)|Gets or sets the width of the bottom gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthLeft](customcontrol-gridlinewidthleft-property-access.md)|Gets or sets the width of the left gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthRight](customcontrol-gridlinewidthright-property-access.md)|Gets or sets the width of the right gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthTop](customcontrol-gridlinewidthtop-property-access.md)|Gets or sets the width of the top gridline for the specified text box. Read/write  **Byte**.|
|[Height](customcontrol-height-property-access.md)|Gets or sets the height of the specified object in twips. Read/write  **Integer**.|
|[HelpContextId](customcontrol-helpcontextid-property-access.md)|The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.|
|[HorizontalAnchor](customcontrol-horizontalanchor-property-access.md)|Gets or sets an  **[AcHorizontalAnchor](achorizontalanchor-enumeration-access.md)** constant that indicates how the text box is anchored horizontally within its layout. Read/write.|
|[InSelection](customcontrol-inselection-property-access.md)|You can use the  **InSelection** property to determine or specify whether a control on a form in Design view is selected. Read/write **Boolean**.|
|[IsVisible](customcontrol-isvisible-property-access.md)|You can use the  **IsVisible** property in to determine whether a control on a report is visible. Read/write **Boolean**.|
|[Layout](customcontrol-layout-property-access.md)|Returns the type of layout for the specified text box. Read-only  **[AcLayoutType](aclayouttype-enumeration-access.md)**.|
|[LayoutID](customcontrol-layoutid-property-access.md)|Returns the unique identifier for the layout that contains the specified text box. Read-only  **Long**.|
|[Left](customcontrol-left-property-access.md)|You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.|
|[LeftPadding](customcontrol-leftpadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its left gridline. Read/write  **Integer**.|
|[Locked](customcontrol-locked-property-access.md)|The  **Locked** property specifies whether you can edit data in a control in Form view. Read/write **Boolean**.|
|[Name](customcontrol-name-property-access.md)|You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.|
|[Object](customcontrol-object-property-access.md)|You can use the  **Object** property in Visual Basic to return a reference to the ActiveX object that is associated with a linked or embedded OLE object in a control. By using this reference, you can access the properties or invoke the methods of the OLE object. Read-only **Object**.|
|[ObjectPalette](customcontrol-objectpalette-property-access.md)|The  **ObjectPalette** property specifies the palette in the application used to create an OLE object. Read/write **Variant**.|
|[ObjectVerbs](customcontrol-objectverbs-property-access.md)|You can use the  **ObjectVerbs** property in Visual Basic to determine the list of verbs an OLE object supports. Read-only **String**.|
|[ObjectVerbsCount](customcontrol-objectverbscount-property-access.md)|You can use the  **ObjectVerbsCount** property in Visual Basic to determine the number of verbs supported by an OLE object. Read-only **Long**.|
|[OldBorderStyle](customcontrol-oldborderstyle-property-access.md)|You can use this property to set or returns the unedited value of the  **BorderStyle** property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write **Byte**.|
|[OldValue](customcontrol-oldvalue-property-access.md)|You can use the  **OldValue** property to determine the unedited value of a bound control. Read-only **Variant**.|
|[OLEClass](customcontrol-oleclass-property-access.md)|You can use the  **OLEClass** property to obtain a description of the kind of OLE object contained in a chart control or an unbound object frame. Read-only **String**.|
|[OnEnter](customcontrol-onenter-property-access.md)|Sets or returns the value of the  **On Enter** box in the **Properties** window of specified object. Read/write **String**. .|
|[OnExit](customcontrol-onexit-property-access.md)|Sets or returns the value of the  **On Exit** box in the **Properties** window of specified object. Read/write **String**. .|
|[OnGotFocus](customcontrol-ongotfocus-property-access.md)|Sets or returns the value of the  **On Got Focus** box in the **Properties** window of the specified object. Read/write **String**.|
|[OnLostFocus](customcontrol-onlostfocus-property-access.md)|Sets or returns the value of the  **On Lost Focus** box in the **Properties** window of the specified object. Read/write **String**.|
|[OnUpdated](customcontrol-onupdated-property-access.md)|Sets or returns the value of the  **On Updated** box in the **Properties** window of a form or report. Read/write **String**.|
|[Parent](customcontrol-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Properties](customcontrol-properties-property-access.md)|Returns a reference to a control's **[Properties](properties-object-access.md)** collection object. Read-only.|
|[RightPadding](customcontrol-rightpadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its right gridline. Read/write  **Integer**.|
|[Section](customcontrol-section-property-access.md)|You can identify these controls by the section of a form or report where the control appears. Read/write  **Integer**.|
|[SpecialEffect](customcontrol-specialeffect-property-access.md)|You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.|
|[TabIndex](customcontrol-tabindex-property-access.md)|You can use the  **TabIndex** property to specify a control's place in the tab order on a form or report. Read/write **Integer**.|
|[TabStop](customcontrol-tabstop-property-access.md)|You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control. Read/write **Boolean**.|
|[Tag](customcontrol-tag-property-access.md)|Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write  **String**.|
|[Top](customcontrol-top-property-access.md)|You can use the  **Top** property to specify an object's location on a form or report. Read/write **Integer**. .|
|[TopPadding](customcontrol-toppadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its top gridline. Read/write  **Integer**.|
|[Value](customcontrol-value-property-access.md)|Gets or sets the value displayed in the specified control. Read/write  **Variant**.|
|[VarOleObject](customcontrol-varoleobject-property-access.md)| Gets a pointer to an **IOLEObject** that represents the memory address of an OLE object. Read-only **Variant**.|
|[Verb](customcontrol-verb-property-access.md)|You can use the  **Verb** property to specify the operation to perform when an OLE object is activated, which is permitted when the control's **Action** property is set to **acOLEActivate**. Read/write **Long**.|
|[VerticalAnchor](customcontrol-verticalanchor-property-access.md)|Gets or sets an [AcVerticalAnchor](acverticalanchor-enumeration-access.md) constant that indicates how the specified text box is anchored vertically within its layout. Read/write.|
|[Visible](customcontrol-visible-property-access.md)|Returns or sets whether the object is visible. Read/write  **Boolean**.|
|[Width](customcontrol-width-property-access.md)|Gets or sets the width of the specified object in twips. Read/write  **Integer**.|

