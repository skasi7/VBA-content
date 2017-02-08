---
title: WebBrowserControl Members (Access)
ms.prod: ACCESS
ms.assetid: bd19a10a-fbbc-5fd6-0818-23a377be9583
---


# WebBrowserControl Members (Access)


Represents a Web browser control on a form.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterUpdate](webbrowsercontrol-afterupdate-event-access.md)|The  **AfterUpdate** event occurs after changed data in a control or record is updated.|
|[BeforeNavigate2](webbrowsercontrol-beforenavigate2-event-access.md)|Occurs before navigation occurs in the given  **WebBrowserControl**.|
|[BeforeUpdate](webbrowsercontrol-beforeupdate-event-access.md)|The  **BeforeUpdate** event occurs before changed data in a control or record is updated.|
|[Change](webbrowsercontrol-change-event-access.md)|The  **Change** event occurs when the contents of the specified control changes.|
|[Click](webbrowsercontrol-click-event-access.md)|The  **Click** event occurs when the user presses and then releases a mouse button over an object.|
|[DblClick](webbrowsercontrol-dblclick-event-access.md)|The  **DblClick** event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.|
|[Dirty](webbrowsercontrol-dirty-event-access.md)|The  **Dirty** event occurs when the contents of the specified control changes.|
|[DocumentComplete](webbrowsercontrol-documentcomplete-event-access.md)|Occurs when a document is completely loaded and initialized.|
|[Enter](webbrowsercontrol-enter-event-access.md)|The  **Enter** event occurs before a control actually receives the focus from a control on the same form or report.|
|[Exit](webbrowsercontrol-exit-event-access.md)|The  **Exit** event occurs just before a control loses the focus to another control on the same form or report.|
|[GotFocus](webbrowsercontrol-gotfocus-event-access.md)|The  **GotFocus** event occurs when the specified object receives the focus.|
|[KeyDown](webbrowsercontrol-keydown-event-access.md)|The  **KeyDown** event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyPress](webbrowsercontrol-keypress-event-access.md)|The  **KeyPress** event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyUp](webbrowsercontrol-keyup-event-access.md)|The  **KeyUp** event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[LostFocus](webbrowsercontrol-lostfocus-event-access.md)|The  **LostFocus** event occurs when the specified object loses the focus.|
|[MouseDown](webbrowsercontrol-mousedown-event-access.md)|The  **MouseDown** event occurs when the user presses a mouse button.|
|[MouseMove](webbrowsercontrol-mousemove-event-access.md)|The  **MouseMove** event occurs when the user moves the mouse.|
|[MouseUp](webbrowsercontrol-mouseup-event-access.md)|The  **MouseUp** event occurs when the user releases a mouse button.|
|[NavigateError](webbrowsercontrol-navigateerror-event-access.md)|Occurs when an error occurs during navigation.|
|[ProgressChange](webbrowsercontrol-progresschange-event-access.md)|Occurs when the progress of a download operation is updated.|
|[Updated](webbrowsercontrol-updated-event-access.md)|The  **Updated** event occurs when an OLE object's data has been modified.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Move](webbrowsercontrol-move-method-access.md)|Moves the specified object to the coordinates specified by the argument values.|
|[Requery](webbrowsercontrol-requery-method-access.md)|The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.|
|[SetFocus](webbrowsercontrol-setfocus-method-access.md)|The  **SetFocus** method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.|
|[SizeToFit](webbrowsercontrol-sizetofit-method-access.md)|You can use the  **SizeToFit** method to size a control so it fits the text or image that it contains.|
|[Undo](webbrowsercontrol-undo-method-access.md)|You can use the  **Undo** method to reset a control or form when its value has been changed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](webbrowsercontrol-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access[Application](application-object-access.md)object and its related properties. Read-only  **Application** object.|
|[BorderColor](webbrowsercontrol-bordercolor-property-access.md)|You can use the  **BorderColor** property to specify the color of a control's border. Read/write **Long**.|
|[BorderShade](webbrowsercontrol-bordershade-property-access.md)|Gets or sets the shade applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderStyle](webbrowsercontrol-borderstyle-property-access.md)|Specifies how a control's border appears.Read/write  **Byte**.|
|[BorderThemeColorIndex](webbrowsercontrol-borderthemecolorindex-property-access.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BorderColor** property of the specified object. Read/write **Long**.|
|[BorderTint](webbrowsercontrol-bordertint-property-access.md)|Gets or sets the tint that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
|[BorderWidth](webbrowsercontrol-borderwidth-property-access.md)|You can use the  **BorderWidth** property to specify the width of a control's border. Read/write **Byte**.|
|[BottomPadding](webbrowsercontrol-bottompadding-property-access.md)|Gets or sets the amount of space (in inches) between the list box and its bottom gridline. Read/write  **Integer**.|
|[Controls](webbrowsercontrol-controls-property-access.md)|Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.|
|[ControlSource](webbrowsercontrol-controlsource-property-access.md)|You can use the  **ControlSource** property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write **String**.|
|[ControlTipText](webbrowsercontrol-controltiptext-property-access.md)|You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.|
|[ControlType](webbrowsercontrol-controltype-property-access.md)|You can use the  **ControlType** property in Visual Basic to determine the type of a control on a form or report. Read/write **Byte**.|
|[DisplayWhen](webbrowsercontrol-displaywhen-property-access.md)|You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.|
|[Enabled](webbrowsercontrol-enabled-property-access.md)|You can use the  **Enabled** property to set or return the status of the conditional format in the[FormatCondition](formatcondition-object-access.md)object. Read/write  **Boolean**.|
|[EventProcPrefix](webbrowsercontrol-eventprocprefix-property-access.md)|Gets or sets the prefix portion of an event procedure name. Read/write  **String**.|
|[GridlineColor](webbrowsercontrol-gridlinecolor-property-access.md)|Gets or sets the color of the gridline for the specified list box. Read/write  **Long**.|
|[GridlineShade](webbrowsercontrol-gridlineshade-property-access.md)|Gets or sets the shade applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
|[GridlineStyleBottom](webbrowsercontrol-gridlinestylebottom-property-access.md)|Gets or sets the bottom gridline style of the specified list box. Read/write  **Byte**.|
|[GridlineStyleLeft](webbrowsercontrol-gridlinestyleleft-property-access.md)|Gets or sets the width of the bottom gridline for the specified text box. Read/write  **Byte**.|
|[GridlineStyleRight](webbrowsercontrol-gridlinestyleright-property-access.md)|Gets or sets the right gridline style of the specified text box. Read/write  **Byte**.|
|[GridlineStyleTop](webbrowsercontrol-gridlinestyletop-property-access.md)|Gets or sets the top gridline style of the specified text box. Read/write  **Byte**.|
|[GridlineThemeColorIndex](webbrowsercontrol-gridlinethemecolorindex-property-access.md)|Gets or sets the theme color index that represents a color in the applied color theme associated with the  **GridlineColor** property of the specified object. Read/write **Long**.|
|[GridlineTint](webbrowsercontrol-gridlinetint-property-access.md)|Gets or sets the tint applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
|[GridlineWidthBottom](webbrowsercontrol-gridlinewidthbottom-property-access.md)|Gets or sets the width of the bottom gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthLeft](webbrowsercontrol-gridlinewidthleft-property-access.md)|Gets or sets the width of the left gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthRight](webbrowsercontrol-gridlinewidthright-property-access.md)|Gets or sets the width of the right gridline for the specified text box. Read/write  **Byte**.|
|[GridlineWidthTop](webbrowsercontrol-gridlinewidthtop-property-access.md)|Gets or sets the width of the top gridline for the specified text box. Read/write  **Byte**.|
|[Height](webbrowsercontrol-height-property-access.md)|Gets or sets the height of the specified object in twips. Read/write  **Integer**.|
|[HelpContextId](webbrowsercontrol-helpcontextid-property-access.md)|The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.|
|[HorizontalAnchor](webbrowsercontrol-horizontalanchor-property-access.md)|Gets or sets an [AcHorizontalAnchor](achorizontalanchor-enumeration-access.md) constant that indicates how the text box is anchored horizontally within its layout. Read/write.|
|[Hyperlink](webbrowsercontrol-hyperlink-property-access.md)|You can use the  **Hyperlink** property to return a reference to a **Hyperlink** object. You can use the **Hyperlink** property to access the properties and methods of a control's hyperlink. Read-only.|
|[InSelection](webbrowsercontrol-inselection-property-access.md)|You can use the  **InSelection** property to determine or specify whether a control on a form in Design view is selected. Read/write **Boolean**.|
|[Layout](webbrowsercontrol-layout-property-access.md)|Returns the type of layout for the specified text box. Read-only [AcLayoutType](aclayouttype-enumeration-access.md).|
|[LayoutID](webbrowsercontrol-layoutid-property-access.md)|Returns the unique identifier for the layout that contains the specified text box. Read-only  **Long**.|
|[Left](webbrowsercontrol-left-property-access.md)|You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.|
|[LeftPadding](webbrowsercontrol-leftpadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its left gridline. Read/write  **Integer**.|
|[LocationURL](webbrowsercontrol-locationurl-property-access.md)|Gets the Uniform Resource Locator (URL) of the current document. Read-only  **String**.|
|[Name](webbrowsercontrol-name-property-access.md)|You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.|
|[Object](webbrowsercontrol-object-property-access.md)|You can use the  **Object** property in Visual Basic to return a reference to the ActiveX object that is associated with a linked or embedded OLE object in a control. By using this reference, you can access the properties or invoke the methods of the OLE object. Read-only **Object**.|
|[OldValue](webbrowsercontrol-oldvalue-property-access.md)|You can use the  **OldValue** property to determine the unedited value of a bound control. Read-only **Variant**.|
|[OnBeforeNavigate](webbrowsercontrol-onbeforenavigate-property-access.md)|Gets or sets the value of the  **On Before Navigate** box in the property sheet os a Web Browser control. Read/write **String**.|
|[OnDocumentComplete](webbrowsercontrol-ondocumentcomplete-property-access.md)|Gets or sets the value of the  **On Document Complete** box in the property sheet os a Web Browser control. Read/write **String**.|
|[OnKeyDown](webbrowsercontrol-onkeydown-property-access.md)|Sets or returns the value of the  **On Key Down** box in the **Properties** window. Read/write **String**.|
|[OnKeyPress](webbrowsercontrol-onkeypress-property-access.md)|Sets or returns the value of the  **On Key Press** box in the **Properties** window. Read/write **String**.|
|[OnKeyUp](webbrowsercontrol-onkeyup-property-access.md)|Sets or returns the value of the  **On Key Up** box in the **Properties** window. Read/write **String**.|
|[OnMouseDown](webbrowsercontrol-onmousedown-property-access.md)|Sets or returns the value of the  **On Mouse Down** box in the **Properties** window. Read/write **String**.|
|[OnMouseMove](webbrowsercontrol-onmousemove-property-access.md)|Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.|
|[OnMouseUp](webbrowsercontrol-onmouseup-property-access.md)|Sets or returns the value of the  **On Mouse Up** box in the **Properties** window. Read/write **String**.|
|[OnNavigateError](webbrowsercontrol-onnavigateerror-property-access.md)|Gets or sets the value of the  **On Navigate Error** box in the property sheet os a Web Browser control. Read/write **String**.|
|[OnProgressChange](webbrowsercontrol-onprogresschange-property-access.md)|Gets or sets the value of the  **On Progress Change** box in the property sheet os a Web Browser control. Read/write **String**.|
|[OnUpdated](webbrowsercontrol-onupdated-property-access.md)|Sets or returns the value of the  **On Updated** box in the **Properties** window of a form or report. Read/write **String**.|
|[Parent](webbrowsercontrol-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Progress](webbrowsercontrol-progress-property-access.md)|Specifies the amount of total progress of a download operation. Read-only  **Long**.|
|[Properties](webbrowsercontrol-properties-property-access.md)|Returns a reference to a control's[Properties](properties-object-access.md)collection object. Read-only.|
|[ReadyState](webbrowsercontrol-readystate-property-access.md)|Gets the status of the specified Web Browser control. Read-only [AcWebBrowserState](acwebbrowserstate-enumeration-access.md).|
|[RightPadding](webbrowsercontrol-rightpadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its right gridline. Read/write  **Integer**.|
|[ScrollBars](webbrowsercontrol-scrollbars-property-access.md)|You can use the  **ScrollBars** property to specify whether scroll bars appear on a text box control. Read/write **Byte**.|
|[ScrollLeft](webbrowsercontrol-scrollleft-property-access.md)|Gets or sets the distance, in pixels, between the left edge of the  **WebBrowser** object and the leftmost portion of the content currently visible in the control. Read/write **Long**.|
|[ScrollTop](webbrowsercontrol-scrolltop-property-access.md)|Gets or sets the distance, in pixels, between the top edge of the  **WebBrowser** object and the topmost portion of the content currently visible in the control. Read/write **Long**.|
|[Section](webbrowsercontrol-section-property-access.md)|You can identify these controls by the section of a form or report where the control appears. Read/write  **Integer**.|
|[SpecialEffect](webbrowsercontrol-specialeffect-property-access.md)|You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.|
|[StatusBarText](webbrowsercontrol-statusbartext-property-access.md)|You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.|
|[TabIndex](webbrowsercontrol-tabindex-property-access.md)|You can use the  **TabIndex** property to specify a control's place in the tab order on a form or report. Read/write **Integer**.|
|[TabStop](webbrowsercontrol-tabstop-property-access.md)|You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control. Read/write **Boolean**.|
|[Tag](webbrowsercontrol-tag-property-access.md)|Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write  **String**.|
|[Top](webbrowsercontrol-top-property-access.md)|You can use the  **Top** property to specify an object's location on a form or report. Read/write **Integer**.|
|[TopPadding](webbrowsercontrol-toppadding-property-access.md)|Gets or sets the amount of space (in inches) between the text box and its top gridline. Read/write  **Integer**.|
|[Transform](webbrowsercontrol-transform-property-access.md)|Read/write|
|[Value](webbrowsercontrol-value-property-access.md)|Determines or specifies the text in the text box. Read/write  **Variant**.|
|[VerticalAnchor](webbrowsercontrol-verticalanchor-property-access.md)|Gets or sets an [AcVerticalAnchor](acverticalanchor-enumeration-access.md) constant that indicates how the specified text box is anchored vertically within its layout. Read/write.|
|[Visible](webbrowsercontrol-visible-property-access.md)|Returns or sets whether the object is visible. Read/write  **Boolean**.|
|[Width](webbrowsercontrol-width-property-access.md)|Gets or sets the width of the specified object in twips. Read/write  **Integer**.|

