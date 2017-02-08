---
title: ListBox Object (Access)
keywords: vbaac10.chm11354
f1_keywords:
- vbaac10.chm11354
ms.prod: ACCESS
api_name:
- Access.ListBox
ms.assetid: 6bc00755-34e7-4fc2-8e72-40dae2010dd8
---


# ListBox Object (Access)

This object corresponds to a list box control. The list box control displays a list of values or alternatives.


## Remarks

In many cases, it's quicker and easier to select a value from a list than to remember a value to type. A list of choices also helps ensure that the value that's entered in a field is correct.


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|
![List box control](/images/t-lstbox_ZA06053984.gif)

|
![List box tool](/images/listbox_ZA06044481.gif)

|

 **Note**  The list in a list box consists of rows of data. Rows can have one or more columns, which can appear with or without headings.


![Multi-column list box](/images/cfrmlst2_ZA06047456.gif)

If a multiple-column list box is bound, Microsoft Access stores the values from one of the columns.

You can use an unbound list box to store a value that you can use with another control. For example, you could use an unbound list box to limit the values in another list box or in a custom dialog box. You could also use an unbound list box to find a record based on the value you select in the list box.

If you don't have room on your form to display a list box, or if you want to be able to type new values as well as select values from a list, use a combo box instead of a list box.

 **Links provided by:**
![Community Member Icon](/images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community



- [Create a Query that uses a Multi-select Listbox as Criteria](http://www.utteraccess.com/forum/Creating-Query-Multi-t414388.mdl)
    
- [ListBox Picker](http://www.utteraccess.com/forum/ListBox-Picker-t426483.mdl)
    
- [Move/Change Order of List Box Items with Up/Down Buttons](http://www.utteraccess.com/wiki/index.php/List_Box:_Reorder_Items)
    
- [Populate a Listbox with Files from a Directory](http://www.utteraccess.com/forum/Populate-Listbox-Files-t1209291.mdl)
    

## Example

This example demonstrates how to filter the contents of a list box while you are typing in a text box.

In this example, a list box named ColorID displays a list of colors stored in the Colors table. As you type in the FilterBy text box, the items in ColorID are filtered dynamically

To do this, use the Change event of the text box to build a SQL statement that will serve as the new RowSource of the list box.

 **Sample code provided by:**
![Community Member Icon](/images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```
Private Sub FilterBy_Change()

    Dim sql As String
    
    'This will match any entry in the list that begins with what the user 
    'has typed in the FilterBy control
    sql = "SELECT ColorID, ColorName FROM Colors WHERE ColorName Like '" &amp; Me.FilterBy.Text &amp; "*' ORDER BY ColorName"
    
    'If you want to match any part of the string then add wildcard (*) before
    'the FilterBy.Text, too:
    'sql = "SELECT ColorID, ColorName FROM Colors WHERE ColorName Like '*" &amp; Me.FilterBy.Text &amp; "*' ORDER BY ColorName"
    
    Me.ColorID.RowSource = sql
    
End Sub
```


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/listbox-afterupdate-event-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/listbox-beforeupdate-event-access%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/listbox-click-event-access%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/listbox-dblclick-event-access%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/listbox-enter-event-access%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/listbox-exit-event-access%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/listbox-gotfocus-event-access%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/listbox-keydown-event-access%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/listbox-keypress-event-access%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/listbox-keyup-event-access%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/listbox-lostfocus-event-access%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/listbox-mousedown-event-access%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/listbox-mousemove-event-access%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/listbox-mouseup-event-access%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddItem](http://msdn.microsoft.com/library/listbox-additem-method-access%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/listbox-move-method-access%28Office.15%29.aspx)|
|[RemoveItem](http://msdn.microsoft.com/library/listbox-removeitem-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/listbox-requery-method-access%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/listbox-setfocus-method-access%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/listbox-sizetofit-method-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/listbox-undo-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/listbox-addcolon-property-access%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/listbox-afterupdate-property-access%28Office.15%29.aspx)|
|[AllowValueListEdits](http://msdn.microsoft.com/library/listbox-allowvaluelistedits-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/listbox-application-property-access%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/listbox-autolabel-property-access%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/listbox-backcolor-property-access%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/listbox-backshade-property-access%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/listbox-backthemecolorindex-property-access%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/listbox-backtint-property-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/listbox-beforeupdate-property-access%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/listbox-bordercolor-property-access%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/listbox-bordershade-property-access%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/listbox-borderstyle-property-access%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/listbox-borderthemecolorindex-property-access%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/listbox-bordertint-property-access%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/listbox-borderwidth-property-access%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/listbox-bottompadding-property-access%28Office.15%29.aspx)|
|[BoundColumn](http://msdn.microsoft.com/library/listbox-boundcolumn-property-access%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/listbox-column-property-access%28Office.15%29.aspx)|
|[ColumnCount](http://msdn.microsoft.com/library/listbox-columncount-property-access%28Office.15%29.aspx)|
|[ColumnHeads](http://msdn.microsoft.com/library/listbox-columnheads-property-access%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/listbox-columnhidden-property-access%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/listbox-columnorder-property-access%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/listbox-columnwidth-property-access%28Office.15%29.aspx)|
|[ColumnWidths](http://msdn.microsoft.com/library/listbox-columnwidths-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/listbox-controls-property-access%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/listbox-controlsource-property-access%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/listbox-controltiptext-property-access%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/listbox-controltype-property-access%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/listbox-defaultvalue-property-access%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/listbox-displaywhen-property-access%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/listbox-enabled-property-access%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/listbox-eventprocprefix-property-access%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/listbox-fontbold-property-access%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/listbox-fontitalic-property-access%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/listbox-fontname-property-access%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/listbox-fontsize-property-access%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/listbox-fontunderline-property-access%28Office.15%29.aspx)|
|[FontWeight](http://msdn.microsoft.com/library/listbox-fontweight-property-access%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/listbox-forecolor-property-access%28Office.15%29.aspx)|
|[ForeShade](http://msdn.microsoft.com/library/listbox-foreshade-property-access%28Office.15%29.aspx)|
|[ForeThemeColorIndex](http://msdn.microsoft.com/library/listbox-forethemecolorindex-property-access%28Office.15%29.aspx)|
|[ForeTint](http://msdn.microsoft.com/library/listbox-foretint-property-access%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/listbox-gridlinecolor-property-access%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/listbox-gridlineshade-property-access%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/listbox-gridlinestylebottom-property-access%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/listbox-gridlinestyleleft-property-access%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/listbox-gridlinestyleright-property-access%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/listbox-gridlinestyletop-property-access%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/listbox-gridlinethemecolorindex-property-access%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/listbox-gridlinetint-property-access%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/listbox-gridlinewidthbottom-property-access%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/listbox-gridlinewidthleft-property-access%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/listbox-gridlinewidthright-property-access%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/listbox-gridlinewidthtop-property-access%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/listbox-height-property-access%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/listbox-helpcontextid-property-access%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/listbox-hideduplicates-property-access%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/listbox-horizontalanchor-property-access%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/listbox-hyperlink-property-access%28Office.15%29.aspx)|
|[IMEHold](http://msdn.microsoft.com/library/listbox-imehold-property-access%28Office.15%29.aspx)|
|[IMEMode](http://msdn.microsoft.com/library/listbox-imemode-property-access%28Office.15%29.aspx)|
|[IMESentenceMode](http://msdn.microsoft.com/library/listbox-imesentencemode-property-access%28Office.15%29.aspx)|
|[InheritValueList](http://msdn.microsoft.com/library/listbox-inheritvaluelist-property-access%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/listbox-inselection-property-access%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/listbox-isvisible-property-access%28Office.15%29.aspx)|
|[ItemData](http://msdn.microsoft.com/library/listbox-itemdata-property-access%28Office.15%29.aspx)|
|[ItemsSelected](http://msdn.microsoft.com/library/listbox-itemsselected-property-access%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/listbox-labelalign-property-access%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/listbox-labelx-property-access%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/listbox-labely-property-access%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/listbox-layout-property-access%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/listbox-layoutid-property-access%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/listbox-left-property-access%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/listbox-leftpadding-property-access%28Office.15%29.aspx)|
|[ListCount](http://msdn.microsoft.com/library/listbox-listcount-property-access%28Office.15%29.aspx)|
|[ListIndex](http://msdn.microsoft.com/library/listbox-listindex-property-access%28Office.15%29.aspx)|
|[ListItemsEditForm](http://msdn.microsoft.com/library/listbox-listitemseditform-property-access%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/listbox-locked-property-access%28Office.15%29.aspx)|
|[MultiSelect](http://msdn.microsoft.com/library/listbox-multiselect-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/listbox-name-property-access%28Office.15%29.aspx)|
|[NumeralShapes](http://msdn.microsoft.com/library/listbox-numeralshapes-property-access%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/listbox-oldborderstyle-property-access%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/listbox-oldvalue-property-access%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/listbox-onclick-property-access%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/listbox-ondblclick-property-access%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/listbox-onenter-property-access%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/listbox-onexit-property-access%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/listbox-ongotfocus-property-access%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/listbox-onkeydown-property-access%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/listbox-onkeypress-property-access%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/listbox-onkeyup-property-access%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/listbox-onlostfocus-property-access%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/listbox-onmousedown-property-access%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/listbox-onmousemove-property-access%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/listbox-onmouseup-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/listbox-parent-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/listbox-properties-property-access%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/listbox-readingorder-property-access%28Office.15%29.aspx)|
|[Recordset](http://msdn.microsoft.com/library/listbox-recordset-property-access%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/listbox-rightpadding-property-access%28Office.15%29.aspx)|
|[RowSource](http://msdn.microsoft.com/library/listbox-rowsource-property-access%28Office.15%29.aspx)|
|[RowSourceType](http://msdn.microsoft.com/library/listbox-rowsourcetype-property-access%28Office.15%29.aspx)|
|[ScrollBarAlign](http://msdn.microsoft.com/library/listbox-scrollbaralign-property-access%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/listbox-section-property-access%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/listbox-selected-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/listbox-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[ShowOnlyRowSourceValues](http://msdn.microsoft.com/library/listbox-showonlyrowsourcevalues-property-access%28Office.15%29.aspx)|
|[SmartTags](http://msdn.microsoft.com/library/listbox-smarttags-property-access%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/listbox-specialeffect-property-access%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/listbox-statusbartext-property-access%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/listbox-tabindex-property-access%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/listbox-tabstop-property-access%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/listbox-tag-property-access%28Office.15%29.aspx)|
|[ThemeFontIndex](http://msdn.microsoft.com/library/listbox-themefontindex-property-access%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/listbox-top-property-access%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/listbox-toppadding-property-access%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/listbox-validationrule-property-access%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/listbox-validationtext-property-access%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/listbox-value-property-access%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/listbox-verticalanchor-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/listbox-visible-property-access%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/listbox-width-property-access%28Office.15%29.aspx)|

## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


<<<<<<< HEAD
#### Other resources

[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
=======
#### Concepts


[Access Object Model Reference](object-model-access-vba-reference.md)
>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

