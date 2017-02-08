---
title: ComboBox Object (Access)
keywords: vbaac10.chm11545
f1_keywords:
- vbaac10.chm11545
ms.prod: ACCESS
api_name:
- Access.ComboBox
ms.assetid: 1cf508d5-023e-eb38-3991-71e82b2a4e7e
---


# ComboBox Object (Access)

This object corresponds to a combo box control. The combo box control combines the features of a text box and a list box. Use a combo box when you want the option of either typing a value or selecting a value from a predefined list.


## Remarks


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|
![Combo box control](/images/t-combox_ZA06053980.gif)

|
![Combo box tool](/images/a_combobox_ZA06047114.gif)

|
In Form view, Microsoft Access doesn't display the list until you click the combo box's arrow.

If you have Control Wizards on before you select the combo box tool, you can create a combo box with a wizard. To turn Control Wizards on or off, click the  **Control Wizards** tool in the toolbox.

The setting of the  **LimitToList** property determines whether you can enter values that aren't in the list.

The list can be single- or multiple-column, and the columns can appear with or without headings.

 **Link provided by:** Luke Chung,[FMS, Inc.](http://www.fmsinc.com/)


- [Tips and Techniques for Using and Validating Combo Boxes](http://www.fmsinc.com/free/NewTips/Access/ComboBox/AccessComboBox.asp)
    
 **Links provided by:**
![Community Member Icon](/images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Combo Box](http://www.utteraccess.com/wiki/index.php/Combo_Box)
    
- [Cascading Combo Boxes](http://www.utteraccess.com/wiki/index.php/Cascading_Combo_Boxes)
    
- [Cascading Combo Boxes: Demo](http://www.utteraccess.com/wiki/index.php/Cascading_Combo_Boxes:_Demo)
    
- [Cascading Combo Boxes - Leaving Null Values](http://www.utteraccess.com/wiki/index.php/Cascade_Combo_Leaving_Null_Values)
    
- [Forms: Populate Controls/Text Boxes Based on Combobox Selection](http://www.utteraccess.com/wiki/index.php/Forms:_Populate_Controls/Text_Boxes_Based_on_Combobox_Selection)
    

## Example

The following example shows how to use multiple  **ComboBox** controls to supply criteria for a query.

 **Sample code provided by:**
![Community Member Icon](/images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community

UtterAccess members can download a database that contains this example from [here](http://www.utteraccess.com/forum/Dynamic-Query-Examples-t1405533.mdl).




```
Private Sub cmdSearch_Click()
    Dim db As Database
    Dim qd As QueryDef
    Dim vWhere As Variant
    
    Set db = CurrentDb()
    
    On Error Resume Next
    db.QueryDefs.Delete "Query1"
    On Error GoTo 0
    
    vWhere = Null
    vWhere = vWhere &amp; " AND [PymtTypeID]=" + Me.cboPaymentTypes
    vWhere = vWhere &amp; " AND [RefundTypeID]=" + Me.cboRefundType
    vWhere = vWhere &amp; " AND [RefundCDMID]=" + Me.cboRefundCDM
    vWhere = vWhere &amp; " AND [RefundOptionID]=" + Me.cboRefundOption
    vWhere = vWhere &amp; " AND [RefundCodeID]=" + Me.cboRefundCode
    
    If Nz(vWhere, "") = "" Then
        MsgBox "There are no search criteria selected." &amp; vbCrLf &amp; vbCrLf &amp; _
        "Search Cancelled.", vbInformation, "Search Canceled."
        
    Else
        Set qd = db.CreateQueryDef("Query1", "SELECT * FROM tblRefundData WHERE " &amp; _
        Mid(vWhere, 6))
        
        db.Close
        Set db = Nothing
        
        DoCmd.OpenQuery "Query1", acViewNormal, acReadOnly
    End If
End Sub
```



The following example shows how to set the  **RowSource** property of a combo box when a form is loaded. When the form is displayed, the items stored in the **Departments** field of the **tblDepartment** combo box are displayed in the **cboDept** combo box.

 **Sample code provided by:**
![MVP Contributor](/images/odc_OfficeTA_33px_MVPContrib.jpg) Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)




```
Private Sub Form_Load()
    Me.Caption = "Today is " &amp; Format$(Date, "dddd mmm-d-yyyy")
    Me.RecordSource = "tblDepartments"
    DoCmd.Maximize  
    txtDept.ControlSource = "Department"
    cmdClose.Caption = "&amp;Close"
    cboDept.RowSourceType = "Table/Query"
    cboDept.RowSource = "SELECT Department FROM tblDepartments"
End Sub
```



The following example show how to create a combo box that is bound to one column while displaying another. Setting the  **ColumnCount** property to 2 specifies that the **cboDept** combo box will display the first two columns of the data source specified by the **RowSource** property. Setting the **BoundColumn** property to 1 specifies that the value stored in the first column will be returned when you inspect the value of the combo box.

The  **ColumnWidths** property specifies the width of the two columns. By setting the width of the first column to **0in.**, the first column is not displayed in the combo box.

 **Sample code provided by:**
![MVP Contributor](/images/odc_OfficeTA_33px_MVPContrib.jpg) Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)




```
Private Sub cboDept_Enter()
    With cboDept
        .RowSource = "SELECT * FROM tblDepartments ORDER BY Department"
        .ColumnCount = 2
        .BoundColumn = 1
        .ColumnWidths = "0in.;1in."
    End With
End Sub
```

The following example shows how to add an item to a bound combo box.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.mdl)




```
Private Sub cboMainCategory_NotInList(NewData As String, Response As Integer)

    On Error GoTo Error_Handler
    Dim intAnswer As Integer
    intAnswer = MsgBox("""" &amp; NewData &amp; """ is not an approved category. " &amp; vbcrlf _
        &amp; "Do you want to add it now?" _ vbYesNo + vbQuestion, "Invalid Category")

    Select Case intAnswer
        Case vbYes
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO tlkpCategoryNotInList (Category) "
                &amp; _ "Select """ &amp; NewData &amp; """;"
            DoCmd.SetWarnings True
            Response = acDataErrAdded
        Case vbNo
            MsgBox "Please select an item from the list.", _
                vbExclamation + vbOKOnly, "Invalid Entry"
            Response = acDataErrContinue

    End Select

    Exit_Procedure:
        DoCmd.SetWarnings True
        Exit Sub

    Error_Handler:
        MsgBox Err.Number &amp; ", " &amp; Error Description
        Resume Exit_Procedure
        Resume

End Sub
```


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/combobox-afterupdate-event-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/combobox-beforeupdate-event-access%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/combobox-change-event-access%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/combobox-click-event-access%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/combobox-dblclick-event-access%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/combobox-dirty-event-access%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/combobox-enter-event-access%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/combobox-exit-event-access%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/combobox-gotfocus-event-access%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/combobox-keydown-event-access%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/combobox-keypress-event-access%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/combobox-keyup-event-access%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/combobox-lostfocus-event-access%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/combobox-mousedown-event-access%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/combobox-mousemove-event-access%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/combobox-mouseup-event-access%28Office.15%29.aspx)|
|[NotInList](http://msdn.microsoft.com/library/combobox-notinlist-event-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/combobox-undo-event-access%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddItem](http://msdn.microsoft.com/library/combobox-additem-method-access%28Office.15%29.aspx)|
|[Dropdown](http://msdn.microsoft.com/library/combobox-dropdown-method-access%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/combobox-move-method-access%28Office.15%29.aspx)|
|[RemoveItem](http://msdn.microsoft.com/library/combobox-removeitem-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/combobox-requery-method-access%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/combobox-setfocus-method-access%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/combobox-sizetofit-method-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/combobox-undo-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/combobox-addcolon-property-access%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/combobox-afterupdate-property-access%28Office.15%29.aspx)|
|[AllowAutoCorrect](http://msdn.microsoft.com/library/combobox-allowautocorrect-property-access%28Office.15%29.aspx)|
|[AllowValueListEdits](http://msdn.microsoft.com/library/combobox-allowvaluelistedits-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/combobox-application-property-access%28Office.15%29.aspx)|
|[AutoExpand](http://msdn.microsoft.com/library/combobox-autoexpand-property-access%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/combobox-autolabel-property-access%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/combobox-backcolor-property-access%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/combobox-backshade-property-access%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/combobox-backstyle-property-access%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/combobox-backthemecolorindex-property-access%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/combobox-backtint-property-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/combobox-beforeupdate-property-access%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/combobox-bordercolor-property-access%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/combobox-bordershade-property-access%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/combobox-borderstyle-property-access%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/combobox-borderthemecolorindex-property-access%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/combobox-bordertint-property-access%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/combobox-borderwidth-property-access%28Office.15%29.aspx)|
|[BottomMargin](http://msdn.microsoft.com/library/combobox-bottommargin-property-access%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/combobox-bottompadding-property-access%28Office.15%29.aspx)|
|[BoundColumn](http://msdn.microsoft.com/library/combobox-boundcolumn-property-access%28Office.15%29.aspx)|
|[CanGrow](http://msdn.microsoft.com/library/combobox-cangrow-property-access%28Office.15%29.aspx)|
|[CanShrink](http://msdn.microsoft.com/library/combobox-canshrink-property-access%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/combobox-column-property-access%28Office.15%29.aspx)|
|[ColumnCount](http://msdn.microsoft.com/library/combobox-columncount-property-access%28Office.15%29.aspx)|
|[ColumnHeads](http://msdn.microsoft.com/library/combobox-columnheads-property-access%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/combobox-columnhidden-property-access%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/combobox-columnorder-property-access%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/combobox-columnwidth-property-access%28Office.15%29.aspx)|
|[ColumnWidths](http://msdn.microsoft.com/library/combobox-columnwidths-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/combobox-controls-property-access%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/combobox-controlsource-property-access%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/combobox-controltiptext-property-access%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/combobox-controltype-property-access%28Office.15%29.aspx)|
|[DecimalPlaces](http://msdn.microsoft.com/library/combobox-decimalplaces-property-access%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/combobox-defaultvalue-property-access%28Office.15%29.aspx)|
|[DisplayAsHyperlink](http://msdn.microsoft.com/library/combobox-displayashyperlink-property-access%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/combobox-displaywhen-property-access%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/combobox-enabled-property-access%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/combobox-eventprocprefix-property-access%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/combobox-fontbold-property-access%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/combobox-fontitalic-property-access%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/combobox-fontname-property-access%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/combobox-fontsize-property-access%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/combobox-fontunderline-property-access%28Office.15%29.aspx)|
|[FontWeight](http://msdn.microsoft.com/library/combobox-fontweight-property-access%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/combobox-forecolor-property-access%28Office.15%29.aspx)|
|[ForeShade](http://msdn.microsoft.com/library/combobox-foreshade-property-access%28Office.15%29.aspx)|
|[ForeThemeColorIndex](http://msdn.microsoft.com/library/combobox-forethemecolorindex-property-access%28Office.15%29.aspx)|
|[ForeTint](http://msdn.microsoft.com/library/combobox-foretint-property-access%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/combobox-format-property-access%28Office.15%29.aspx)|
|[FormatConditions](http://msdn.microsoft.com/library/combobox-formatconditions-property-access%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/combobox-gridlinecolor-property-access%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/combobox-gridlineshade-property-access%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/combobox-gridlinestylebottom-property-access%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/combobox-gridlinestyleleft-property-access%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/combobox-gridlinestyleright-property-access%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/combobox-gridlinestyletop-property-access%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/combobox-gridlinethemecolorindex-property-access%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/combobox-gridlinetint-property-access%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/combobox-gridlinewidthbottom-property-access%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/combobox-gridlinewidthleft-property-access%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/combobox-gridlinewidthright-property-access%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/combobox-gridlinewidthtop-property-access%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/combobox-height-property-access%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/combobox-helpcontextid-property-access%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/combobox-hideduplicates-property-access%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/combobox-horizontalanchor-property-access%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/combobox-hyperlink-property-access%28Office.15%29.aspx)|
|[IMEHold](http://msdn.microsoft.com/library/combobox-imehold-property-access%28Office.15%29.aspx)|
|[IMEMode](http://msdn.microsoft.com/library/combobox-imemode-property-access%28Office.15%29.aspx)|
|[IMESentenceMode](http://msdn.microsoft.com/library/combobox-imesentencemode-property-access%28Office.15%29.aspx)|
|[InheritValueList](http://msdn.microsoft.com/library/combobox-inheritvaluelist-property-access%28Office.15%29.aspx)|
|[InputMask](http://msdn.microsoft.com/library/combobox-inputmask-property-access%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/combobox-inselection-property-access%28Office.15%29.aspx)|
|[IsHyperlink](http://msdn.microsoft.com/library/combobox-ishyperlink-property-access%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/combobox-isvisible-property-access%28Office.15%29.aspx)|
|[ItemData](http://msdn.microsoft.com/library/combobox-itemdata-property-access%28Office.15%29.aspx)|
|[ItemsSelected](http://msdn.microsoft.com/library/combobox-itemsselected-property-access%28Office.15%29.aspx)|
|[KeyboardLanguage](http://msdn.microsoft.com/library/combobox-keyboardlanguage-property-access%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/combobox-labelalign-property-access%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/combobox-labelx-property-access%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/combobox-labely-property-access%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/combobox-layout-property-access%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/combobox-layoutid-property-access%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/combobox-left-property-access%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/combobox-leftmargin-property-access%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/combobox-leftpadding-property-access%28Office.15%29.aspx)|
|[LimitToList](http://msdn.microsoft.com/library/combobox-limittolist-property-access%28Office.15%29.aspx)|
|[ListCount](http://msdn.microsoft.com/library/combobox-listcount-property-access%28Office.15%29.aspx)|
|[ListIndex](http://msdn.microsoft.com/library/combobox-listindex-property-access%28Office.15%29.aspx)|
|[ListItemsEditForm](http://msdn.microsoft.com/library/combobox-listitemseditform-property-access%28Office.15%29.aspx)|
|[ListRows](http://msdn.microsoft.com/library/combobox-listrows-property-access%28Office.15%29.aspx)|
|[ListWidth](http://msdn.microsoft.com/library/combobox-listwidth-property-access%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/combobox-locked-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/combobox-name-property-access%28Office.15%29.aspx)|
|[NumeralShapes](http://msdn.microsoft.com/library/combobox-numeralshapes-property-access%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/combobox-oldborderstyle-property-access%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/combobox-oldvalue-property-access%28Office.15%29.aspx)|
|[OnChange](http://msdn.microsoft.com/library/combobox-onchange-property-access%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/combobox-onclick-property-access%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/combobox-ondblclick-property-access%28Office.15%29.aspx)|
|[OnDirty](http://msdn.microsoft.com/library/combobox-ondirty-property-access%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/combobox-onenter-property-access%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/combobox-onexit-property-access%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/combobox-ongotfocus-property-access%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/combobox-onkeydown-property-access%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/combobox-onkeypress-property-access%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/combobox-onkeyup-property-access%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/combobox-onlostfocus-property-access%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/combobox-onmousedown-property-access%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/combobox-onmousemove-property-access%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/combobox-onmouseup-property-access%28Office.15%29.aspx)|
|[OnNotInList](http://msdn.microsoft.com/library/combobox-onnotinlist-property-access%28Office.15%29.aspx)|
|[OnUndo](http://msdn.microsoft.com/library/combobox-onundo-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/combobox-parent-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/combobox-properties-property-access%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/combobox-readingorder-property-access%28Office.15%29.aspx)|
|[Recordset](http://msdn.microsoft.com/library/combobox-recordset-property-access%28Office.15%29.aspx)|
|[RightMargin](http://msdn.microsoft.com/library/combobox-rightmargin-property-access%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/combobox-rightpadding-property-access%28Office.15%29.aspx)|
|[RowSource](http://msdn.microsoft.com/library/combobox-rowsource-property-access%28Office.15%29.aspx)|
|[RowSourceType](http://msdn.microsoft.com/library/combobox-rowsourcetype-property-access%28Office.15%29.aspx)|
|[ScrollBarAlign](http://msdn.microsoft.com/library/combobox-scrollbaralign-property-access%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/combobox-section-property-access%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/combobox-selected-property-access%28Office.15%29.aspx)|
|[SelLength](http://msdn.microsoft.com/library/combobox-sellength-property-access%28Office.15%29.aspx)|
|[SelStart](http://msdn.microsoft.com/library/combobox-selstart-property-access%28Office.15%29.aspx)|
|[SelText](http://msdn.microsoft.com/library/combobox-seltext-property-access%28Office.15%29.aspx)|
|[SeparatorCharacters](http://msdn.microsoft.com/library/combobox-separatorcharacters-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/combobox-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[ShowOnlyRowSourceValues](http://msdn.microsoft.com/library/combobox-showonlyrowsourcevalues-property-access%28Office.15%29.aspx)|
|[SmartTags](http://msdn.microsoft.com/library/combobox-smarttags-property-access%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/combobox-specialeffect-property-access%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/combobox-statusbartext-property-access%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/combobox-tabindex-property-access%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/combobox-tabstop-property-access%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/combobox-tag-property-access%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/combobox-text-property-access%28Office.15%29.aspx)|
|[TextAlign](http://msdn.microsoft.com/library/combobox-textalign-property-access%28Office.15%29.aspx)|
|[ThemeFontIndex](http://msdn.microsoft.com/library/combobox-themefontindex-property-access%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/combobox-top-property-access%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/combobox-topmargin-property-access%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/combobox-toppadding-property-access%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/combobox-validationrule-property-access%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/combobox-validationtext-property-access%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/combobox-value-property-access%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/combobox-verticalanchor-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/combobox-visible-property-access%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/combobox-width-property-access%28Office.15%29.aspx)|

## About the Contributors
<a name="AboutContributors"> </a>

Luke Chung is the founder and president of FMS, Inc., a leading provider of custom database solutions and developer tools. 

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[ComboBox Object Members](http://msdn.microsoft.com/library/combobox-members-access%28Office.15%29.aspx)
