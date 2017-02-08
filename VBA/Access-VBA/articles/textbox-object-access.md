---
title: TextBox Object (Access)
keywords: vbaac10.chm11201
f1_keywords:
- vbaac10.chm11201
ms.prod: ACCESS
api_name:
- Access.TextBox
ms.assetid: d74fbe9a-0d40-7d28-956f-a2bfd0cfee45
---


# TextBox Object (Access)

This object represents a text box control on a form or report. Text boxes are used to either display data from a record source, or to display the results of a calculation, or to accept input from a user.


## Example

The following code example uses a form with a text box to receive user input. The code displays a message when the user inputs data and then presses Return


```

Private Sub txtValue1_BeforeUpdate(Cancel As Integer)

MsgBox "The Text box is being updated."

End Sub

```


## Remarks

Text boxes can be either bound or unbound. You use a bound text box to display data from a particular field. You use an unbound text box to display the results of a calculation, or to accept input from a user (as in the code example above).


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|
![Text box control](/images/t-txtbox_ZA06054010.gif)

|
![Text box tool](/images/textbox_ZA06044637.gif)

|

## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/textbox-afterupdate-event-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/textbox-beforeupdate-event-access%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/textbox-change-event-access%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/textbox-click-event-access%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/textbox-dblclick-event-access%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/textbox-dirty-event-access%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/textbox-enter-event-access%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/textbox-exit-event-access%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/textbox-gotfocus-event-access%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/textbox-keydown-event-access%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/textbox-keypress-event-access%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/textbox-keyup-event-access%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/textbox-lostfocus-event-access%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/textbox-mousedown-event-access%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/textbox-mousemove-event-access%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/textbox-mouseup-event-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/textbox-undo-event-access%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/textbox-move-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/textbox-requery-method-access%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/textbox-setfocus-method-access%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/textbox-sizetofit-method-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/textbox-undo-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/textbox-addcolon-property-access%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/textbox-afterupdate-property-access%28Office.15%29.aspx)|
|[AllowAutoCorrect](http://msdn.microsoft.com/library/textbox-allowautocorrect-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/textbox-application-property-access%28Office.15%29.aspx)|
|[AsianLineBreak](http://msdn.microsoft.com/library/textbox-asianlinebreak-property-access%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/textbox-autolabel-property-access%28Office.15%29.aspx)|
|[AutoTab](http://msdn.microsoft.com/library/textbox-autotab-property-access%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/textbox-backcolor-property-access%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/textbox-backshade-property-access%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/textbox-backstyle-property-access%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/textbox-backthemecolorindex-property-access%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/textbox-backtint-property-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/textbox-beforeupdate-property-access%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/textbox-bordercolor-property-access%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/textbox-bordershade-property-access%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/textbox-borderstyle-property-access%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/textbox-borderthemecolorindex-property-access%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/textbox-bordertint-property-access%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/textbox-borderwidth-property-access%28Office.15%29.aspx)|
|[BottomMargin](http://msdn.microsoft.com/library/textbox-bottommargin-property-access%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/textbox-bottompadding-property-access%28Office.15%29.aspx)|
|[CanGrow](http://msdn.microsoft.com/library/textbox-cangrow-property-access%28Office.15%29.aspx)|
|[CanShrink](http://msdn.microsoft.com/library/textbox-canshrink-property-access%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/textbox-columnhidden-property-access%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/textbox-columnorder-property-access%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/textbox-columnwidth-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/textbox-controls-property-access%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/textbox-controlsource-property-access%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/textbox-controltiptext-property-access%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/textbox-controltype-property-access%28Office.15%29.aspx)|
|[DecimalPlaces](http://msdn.microsoft.com/library/textbox-decimalplaces-property-access%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/textbox-defaultvalue-property-access%28Office.15%29.aspx)|
|[DisplayAsHyperlink](http://msdn.microsoft.com/library/textbox-displayashyperlink-property-access%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/textbox-displaywhen-property-access%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/textbox-enabled-property-access%28Office.15%29.aspx)|
|[EnterKeyBehavior](http://msdn.microsoft.com/library/textbox-enterkeybehavior-property-access%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/textbox-eventprocprefix-property-access%28Office.15%29.aspx)|
|[FilterLookup](http://msdn.microsoft.com/library/textbox-filterlookup-property-access%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/textbox-fontbold-property-access%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/textbox-fontitalic-property-access%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/textbox-fontname-property-access%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/textbox-fontsize-property-access%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/textbox-fontunderline-property-access%28Office.15%29.aspx)|
|[FontWeight](http://msdn.microsoft.com/library/textbox-fontweight-property-access%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/textbox-forecolor-property-access%28Office.15%29.aspx)|
|[ForeShade](http://msdn.microsoft.com/library/textbox-foreshade-property-access%28Office.15%29.aspx)|
|[ForeThemeColorIndex](http://msdn.microsoft.com/library/textbox-forethemecolorindex-property-access%28Office.15%29.aspx)|
|[ForeTint](http://msdn.microsoft.com/library/textbox-foretint-property-access%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/textbox-format-property-access%28Office.15%29.aspx)|
|[FormatConditions](http://msdn.microsoft.com/library/textbox-formatconditions-property-access%28Office.15%29.aspx)|
|[FuriganaControl](http://msdn.microsoft.com/library/textbox-furiganacontrol-property-access%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/textbox-gridlinecolor-property-access%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/textbox-gridlineshade-property-access%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/textbox-gridlinestylebottom-property-access%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/textbox-gridlinestyleleft-property-access%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/textbox-gridlinestyleright-property-access%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/textbox-gridlinestyletop-property-access%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/textbox-gridlinethemecolorindex-property-access%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/textbox-gridlinetint-property-access%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/textbox-gridlinewidthbottom-property-access%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/textbox-gridlinewidthleft-property-access%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/textbox-gridlinewidthright-property-access%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/textbox-gridlinewidthtop-property-access%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/textbox-height-property-access%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/textbox-helpcontextid-property-access%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/textbox-hideduplicates-property-access%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/textbox-horizontalanchor-property-access%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/textbox-hyperlink-property-access%28Office.15%29.aspx)|
|[IMEHold](http://msdn.microsoft.com/library/textbox-imehold-property-access%28Office.15%29.aspx)|
|[IMEMode](http://msdn.microsoft.com/library/textbox-imemode-property-access%28Office.15%29.aspx)|
|[IMESentenceMode](http://msdn.microsoft.com/library/textbox-imesentencemode-property-access%28Office.15%29.aspx)|
|[InputMask](http://msdn.microsoft.com/library/textbox-inputmask-property-access%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/textbox-inselection-property-access%28Office.15%29.aspx)|
|[IsHyperlink](http://msdn.microsoft.com/library/textbox-ishyperlink-property-access%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/textbox-isvisible-property-access%28Office.15%29.aspx)|
|[KeyboardLanguage](http://msdn.microsoft.com/library/textbox-keyboardlanguage-property-access%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/textbox-labelalign-property-access%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/textbox-labelx-property-access%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/textbox-labely-property-access%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/textbox-layout-property-access%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/textbox-layoutid-property-access%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/textbox-left-property-access%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/textbox-leftmargin-property-access%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/textbox-leftpadding-property-access%28Office.15%29.aspx)|
|[LineSpacing](http://msdn.microsoft.com/library/textbox-linespacing-property-access%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/textbox-locked-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/textbox-name-property-access%28Office.15%29.aspx)|
|[NumeralShapes](http://msdn.microsoft.com/library/textbox-numeralshapes-property-access%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/textbox-oldborderstyle-property-access%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/textbox-oldvalue-property-access%28Office.15%29.aspx)|
|[OnChange](http://msdn.microsoft.com/library/textbox-onchange-property-access%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/textbox-onclick-property-access%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/textbox-ondblclick-property-access%28Office.15%29.aspx)|
|[OnDirty](http://msdn.microsoft.com/library/textbox-ondirty-property-access%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/textbox-onenter-property-access%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/textbox-onexit-property-access%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/textbox-ongotfocus-property-access%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/textbox-onkeydown-property-access%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/textbox-onkeypress-property-access%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/textbox-onkeyup-property-access%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/textbox-onlostfocus-property-access%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/textbox-onmousedown-property-access%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/textbox-onmousemove-property-access%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/textbox-onmouseup-property-access%28Office.15%29.aspx)|
|[OnUndo](http://msdn.microsoft.com/library/textbox-onundo-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/textbox-parent-property-access%28Office.15%29.aspx)|
|[PostalAddress](http://msdn.microsoft.com/library/textbox-postaladdress-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/textbox-properties-property-access%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/textbox-readingorder-property-access%28Office.15%29.aspx)|
|[RightMargin](http://msdn.microsoft.com/library/textbox-rightmargin-property-access%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/textbox-rightpadding-property-access%28Office.15%29.aspx)|
|[RunningSum](http://msdn.microsoft.com/library/textbox-runningsum-property-access%28Office.15%29.aspx)|
|[ScrollBarAlign](http://msdn.microsoft.com/library/textbox-scrollbaralign-property-access%28Office.15%29.aspx)|
|[ScrollBars](http://msdn.microsoft.com/library/textbox-scrollbars-property-access%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/textbox-section-property-access%28Office.15%29.aspx)|
|[SelLength](http://msdn.microsoft.com/library/textbox-sellength-property-access%28Office.15%29.aspx)|
|[SelStart](http://msdn.microsoft.com/library/textbox-selstart-property-access%28Office.15%29.aspx)|
|[SelText](http://msdn.microsoft.com/library/textbox-seltext-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/textbox-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[ShowDatePicker](http://msdn.microsoft.com/library/textbox-showdatepicker-property-access%28Office.15%29.aspx)|
|[SmartTags](http://msdn.microsoft.com/library/textbox-smarttags-property-access%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/textbox-specialeffect-property-access%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/textbox-statusbartext-property-access%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/textbox-tabindex-property-access%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/textbox-tabstop-property-access%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/textbox-tag-property-access%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/textbox-text-property-access%28Office.15%29.aspx)|
|[TextAlign](http://msdn.microsoft.com/library/textbox-textalign-property-access%28Office.15%29.aspx)|
|[TextFormat](http://msdn.microsoft.com/library/textbox-textformat-property-access%28Office.15%29.aspx)|
|[ThemeFontIndex](http://msdn.microsoft.com/library/textbox-themefontindex-property-access%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/textbox-top-property-access%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/textbox-topmargin-property-access%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/textbox-toppadding-property-access%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/textbox-validationrule-property-access%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/textbox-validationtext-property-access%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/textbox-value-property-access%28Office.15%29.aspx)|
|[Vertical](http://msdn.microsoft.com/library/textbox-vertical-property-access%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/textbox-verticalanchor-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/textbox-visible-property-access%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/textbox-width-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[TextBox Object Members](http://msdn.microsoft.com/library/textbox-members-access%28Office.15%29.aspx)
