---
title: Form Object (Access)
keywords: vbaac10.chm13686
f1_keywords:
- vbaac10.chm13686
ms.prod: ACCESS
api_name:
- Access.Form
ms.assetid: 72ef9219-142b-b690-b696-3eba9a5d4522
---


# Form Object (Access)

A  **Form** object refers to a particular Microsoft Access form.


## Remarks

A  **Form** object is a member of the **Forms** collection, which is a collection of all currently open forms. Within the **Forms** collection, individual forms are indexed beginning with zero. You can refer to an individual **Form** object in the **Forms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific form in the **Forms** collection, it's better to refer to the form by name because a form's collection index may change. If the form name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**AllForms** ! _formname_|AllForms!OrderForm|
|**AllForms** ![ _form name_]|AllForms![Order Form]|
|**AllForms** (" _formname_")|AllForms("OrderForm")|
|**AllForms** ( _formname_)|AllForms(0)|
Each  **Form** object has a **Controls** collection, which contains all controls on the form. You can refer to a control on a form either by implicitly or explicitly referring to the **Controls** collection. Your code will be faster if you refer to the **Controls** collection implicitly. The following examples show two of the ways you might refer to a control named **NewData** on the form called **OrderForm**:




```
' Implicit reference. 
Forms!OrderForm!NewData
```




```
' Explicit reference. 
Forms!OrderForm.Controls!NewData
```

The next two examples show how you might refer to a control named  **NewData** on a subform `ctlSubForm` contained in the form called **OrderForm**:




```
Forms!OrderForm.ctlSubForm.Form!Controls.NewData
```




```
Forms!OrderForm.ctlSubForm!NewData
```

 **Links provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) Luke Chung,[FMS, Inc.](http://www.fmsinc.com/)


- [Microsoft Access Form Tips and Avoiding Common Mistakes](http://www.fmsinc.com/tpapers/genaccess/formtips.mdl)
    
- [Microsoft Office Access 2007 Form Design Tips](http://www.fmsinc.com/tpapers/access/Forms/Access2007FormTips.mdl)
    
 **Links provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Display Pictures on a Form](http://www.utteraccess.com/wiki/index.php/Display_Pictures_on_a_Form)
    
- [Display Related Data](http://www.utteraccess.com/wiki/index.php/Display_Related_Data)
    
- [Opening a Detail Form to Related Information](http://www.utteraccess.com/wiki/index.php/Forms:_Open_a_Detail_Form_to_Related_Information)
    
- [Forms: Populate Controls/Text Boxes Based on Combobox Selection](http://www.utteraccess.com/wiki/index.php/Forms:_Populate_Controls/Text_Boxes_Based_on_Combobox_Selection)
    
- [Referring To Properties And Controls On Subforms](http://www.utteraccess.com/wiki/index.php/Referring_To_Properties_And_Controls_On_Subforms)
    

## Example

The following example shows how to use  **TextBox** controls to supply date criteria for a query.

UtterAccess members can download a database that contains this example from [here](http://www.utteraccess.com/forum/Dynamic-Query-Examples-t1405533.mdl).

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```
Private Sub cmdSearch_Click()

   Dim db As DAO.Database
   Dim qd As QueryDef
   Dim vWhere As Variant

   Set db = CurrentDb()

   On Error Resume Next
   db.QueryDefs.Delete "Query1"
   On Error GoTo 0

   vWhere = Null

   vWhere = vWhere &amp; " AND [PayeeID]=" + Me.cboPayeeID

   If Nz(Me.txtEndDate, "") <> "" And Nz(Me.txtStartDate, "") <> "" Then
      vWhere = vWhere &amp; " AND [RefundProcessed] Between #" &amp; _
      Me.txtStartDate &amp; "# AND #" &amp; Me.txtEndDate &amp; "#"
   Else
      If Nz(Me.txtEndDate, "") = "" And Nz(Me.txtStartDate, "") <> "" Then
         vWhere = vWhere &amp; " AND [RefundProcessed]>=#" _
                  + Me.txtStartDate &amp; "#"
      Else
         If Nz(Me.txtEndDate, "") <> "" And Nz(Me.txtStartDate, "") = "" Then
            vWhere = vWhere &amp; " AND [RefundProcessed] <=#" _
                     + Me.txtEndDate &amp; "#"
      End If
     End If
   End If

   If Nz(vWhere, "") = "" Then
      MsgBox "There are no search criteria selected." &amp; vbCrLf &amp; vbCrLf &amp; _
             "Search Cancelled.", vbInformation, "Search Canceled."
   Else
      Set qd = db.CreateQueryDef("Query1", "SELECT * FROM tblRefundData? &amp; _
               " WHERE " &amp; Mid(vWhere, 6))
      db.Close
      Set db = Nothing

      DoCmd.OpenQuery "Query1", acViewNormal, acReadOnly
   End If
End Sub
```

The following example shows how to use the  **BeforeUpdate** event of a form to require that a value be entered into one control when another control also has data.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer?s Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.mdl)




```
Private Sub Form_BeforeUpdate(Cancel As Integer)
If (IsNull(Me.FieldOne)) Or (Me.FieldOne.Value =  "") Then
    ' No action required
Else
    If (IsNull(Me.FieldTwo)) or (Me.FieldTwo.Value = "") Then
        MsgBox "You must provide data for field 'FieldTwo', " &amp; _
            "if a value is entered in FieldOne", _
            vbOKOnly, "Required Field"
        Me.FieldTwo.SetFocus
        Cancel = True
        Exit Sub
    End If
End If

End Sub
```

The following example shows how to use the  **OpenArgs** property to prevent a form from being opened from the Navigation Pane.




```
Private Sub Form_Open(Cancel As Integer)

If Me.OpenArgs() <> "Valid User" Then
    MsgBox "You are not authorized to use this form!", _
        vbExclamation + vbOKOnly, "Invalid Access"
    Cancel = True
End If
End Sub
```

The following example shows how to use the  _WhereCondition_ argument of the **OpenForm** method to filter the records displayed on a form as it is opened.




```
Private Sub cmdShowOrders_Click()
If Not Me.NewRecord Then
    DoCmd.OpenForm "frmOrder", _
        WhereCondition:="CustomerID=" &amp; Me.txtCustomerID
End If
End Sub
```


## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/form-activate-event-access%28Office.15%29.aspx)|
|[AfterDelConfirm](http://msdn.microsoft.com/library/form-afterdelconfirm-event-access%28Office.15%29.aspx)|
|[AfterFinalRender](http://msdn.microsoft.com/library/form-afterfinalrender-event-access%28Office.15%29.aspx)|
|[AfterInsert](http://msdn.microsoft.com/library/form-afterinsert-event-access%28Office.15%29.aspx)|
|[AfterLayout](http://msdn.microsoft.com/library/form-afterlayout-event-access%28Office.15%29.aspx)|
|[AfterRender](http://msdn.microsoft.com/library/form-afterrender-event-access%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/form-afterupdate-event-access%28Office.15%29.aspx)|
|[ApplyFilter](http://msdn.microsoft.com/library/form-applyfilter-event-access%28Office.15%29.aspx)|
|[BeforeDelConfirm](http://msdn.microsoft.com/library/form-beforedelconfirm-event-access%28Office.15%29.aspx)|
|[BeforeInsert](http://msdn.microsoft.com/library/form-beforeinsert-event-access%28Office.15%29.aspx)|
|[BeforeQuery](http://msdn.microsoft.com/library/form-beforequery-event-access%28Office.15%29.aspx)|
|[BeforeRender](http://msdn.microsoft.com/library/form-beforerender-event-access%28Office.15%29.aspx)|
|[BeforeScreenTip](http://msdn.microsoft.com/library/form-beforescreentip-event-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/form-beforeupdate-event-access%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/form-click-event-access%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/form-close-event-access%28Office.15%29.aspx)|
|[CommandBeforeExecute](http://msdn.microsoft.com/library/form-commandbeforeexecute-event-access%28Office.15%29.aspx)|
|[CommandChecked](http://msdn.microsoft.com/library/form-commandchecked-event-access%28Office.15%29.aspx)|
|[CommandEnabled](http://msdn.microsoft.com/library/form-commandenabled-event-access%28Office.15%29.aspx)|
|[CommandExecute](http://msdn.microsoft.com/library/form-commandexecute-event-access%28Office.15%29.aspx)|
|[Current](http://msdn.microsoft.com/library/form-current-event-access%28Office.15%29.aspx)|
|[DataChange](http://msdn.microsoft.com/library/form-datachange-event-access%28Office.15%29.aspx)|
|[DataSetChange](http://msdn.microsoft.com/library/form-datasetchange-event-access%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/form-dblclick-event-access%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/form-deactivate-event-access%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/form-delete-event-access%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/form-dirty-event-access%28Office.15%29.aspx)|
|[Error](http://msdn.microsoft.com/library/form-error-event-access%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/form-filter-event-access%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/form-gotfocus-event-access%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/form-keydown-event-access%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/form-keypress-event-access%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/form-keyup-event-access%28Office.15%29.aspx)|
|[Load](http://msdn.microsoft.com/library/form-load-event-access%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/form-lostfocus-event-access%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/form-mousedown-event-access%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/form-mousemove-event-access%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/form-mouseup-event-access%28Office.15%29.aspx)|
|[MouseWheel](http://msdn.microsoft.com/library/form-mousewheel-event-access%28Office.15%29.aspx)|
|[OnConnect](http://msdn.microsoft.com/library/form-onconnect-event-access%28Office.15%29.aspx)|
|[OnDisconnect](http://msdn.microsoft.com/library/form-ondisconnect-event-access%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/form-open-event-access%28Office.15%29.aspx)|
|[PivotTableChange](http://msdn.microsoft.com/library/form-pivottablechange-event-access%28Office.15%29.aspx)|
|[Query](http://msdn.microsoft.com/library/form-query-event-access%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/form-resize-event-access%28Office.15%29.aspx)|
|[SelectionChange](http://msdn.microsoft.com/library/form-selectionchange-event-access%28Office.15%29.aspx)|
|[Timer](http://msdn.microsoft.com/library/form-timer-event-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/form-undo-event-access%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/form-unload-event-access%28Office.15%29.aspx)|
|[ViewChange](http://msdn.microsoft.com/library/form-viewchange-event-access%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[GoToPage](http://msdn.microsoft.com/library/form-gotopage-method-access%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/form-move-method-access%28Office.15%29.aspx)|
|[Recalc](http://msdn.microsoft.com/library/form-recalc-method-access%28Office.15%29.aspx)|
|[Refresh](http://msdn.microsoft.com/library/form-refresh-method-access%28Office.15%29.aspx)|
|[Repaint](http://msdn.microsoft.com/library/form-repaint-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/form-requery-method-access%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/form-setfocus-method-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/form-undo-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveControl](http://msdn.microsoft.com/library/form-activecontrol-property-access%28Office.15%29.aspx)|
|[AfterDelConfirm](http://msdn.microsoft.com/library/form-afterdelconfirm-property-access%28Office.15%29.aspx)|
|[AfterFinalRender](http://msdn.microsoft.com/library/form-afterfinalrender-property-access%28Office.15%29.aspx)|
|[AfterInsert](http://msdn.microsoft.com/library/form-afterinsert-property-access%28Office.15%29.aspx)|
|[AfterLayout](http://msdn.microsoft.com/library/form-afterlayout-property-access%28Office.15%29.aspx)|
|[AfterRender](http://msdn.microsoft.com/library/form-afterrender-property-access%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/form-afterupdate-property-access%28Office.15%29.aspx)|
|[AllowAdditions](http://msdn.microsoft.com/library/form-allowadditions-property-access%28Office.15%29.aspx)|
|[AllowDatasheetView](http://msdn.microsoft.com/library/form-allowdatasheetview-property-access%28Office.15%29.aspx)|
|[AllowDeletions](http://msdn.microsoft.com/library/form-allowdeletions-property-access%28Office.15%29.aspx)|
|[AllowEdits](http://msdn.microsoft.com/library/form-allowedits-property-access%28Office.15%29.aspx)|
|[AllowFilters](http://msdn.microsoft.com/library/form-allowfilters-property-access%28Office.15%29.aspx)|
|[AllowFormView](http://msdn.microsoft.com/library/form-allowformview-property-access%28Office.15%29.aspx)|
|[AllowLayoutView](http://msdn.microsoft.com/library/form-allowlayoutview-property-access%28Office.15%29.aspx)|
|[AllowPivotChartView](http://msdn.microsoft.com/library/form-allowpivotchartview-property-access%28Office.15%29.aspx)|
|[AllowPivotTableView](http://msdn.microsoft.com/library/form-allowpivottableview-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/form-application-property-access%28Office.15%29.aspx)|
|[AutoCenter](http://msdn.microsoft.com/library/form-autocenter-property-access%28Office.15%29.aspx)|
|[AutoResize](http://msdn.microsoft.com/library/form-autoresize-property-access%28Office.15%29.aspx)|
|[BeforeDelConfirm](http://msdn.microsoft.com/library/form-beforedelconfirm-property-access%28Office.15%29.aspx)|
|[BeforeInsert](http://msdn.microsoft.com/library/form-beforeinsert-property-access%28Office.15%29.aspx)|
|[BeforeQuery](http://msdn.microsoft.com/library/form-beforequery-property-access%28Office.15%29.aspx)|
|[BeforeRender](http://msdn.microsoft.com/library/form-beforerender-property-access%28Office.15%29.aspx)|
|[BeforeScreenTip](http://msdn.microsoft.com/library/form-beforescreentip-property-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/form-beforeupdate-property-access%28Office.15%29.aspx)|
|[Bookmark](http://msdn.microsoft.com/library/form-bookmark-property-access%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/form-borderstyle-property-access%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/form-caption-property-access%28Office.15%29.aspx)|
|[ChartSpace](http://msdn.microsoft.com/library/form-chartspace-property-access%28Office.15%29.aspx)|
|[CloseButton](http://msdn.microsoft.com/library/form-closebutton-property-access%28Office.15%29.aspx)|
|[CommandBeforeExecute](http://msdn.microsoft.com/library/form-commandbeforeexecute-property-access%28Office.15%29.aspx)|
|[CommandChecked](http://msdn.microsoft.com/library/form-commandchecked-property-access%28Office.15%29.aspx)|
|[CommandEnabled](http://msdn.microsoft.com/library/form-commandenabled-property-access%28Office.15%29.aspx)|
|[CommandExecute](http://msdn.microsoft.com/library/form-commandexecute-property-access%28Office.15%29.aspx)|
|[ControlBox](http://msdn.microsoft.com/library/form-controlbox-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/form-controls-property-access%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/form-count-property-access%28Office.15%29.aspx)|
|[CurrentRecord](http://msdn.microsoft.com/library/form-currentrecord-property-access%28Office.15%29.aspx)|
|[CurrentSectionLeft](http://msdn.microsoft.com/library/form-currentsectionleft-property-access%28Office.15%29.aspx)|
|[CurrentSectionTop](http://msdn.microsoft.com/library/form-currentsectiontop-property-access%28Office.15%29.aspx)|
|[CurrentView](http://msdn.microsoft.com/library/form-currentview-property-access%28Office.15%29.aspx)|
|[Cycle](http://msdn.microsoft.com/library/form-cycle-property-access%28Office.15%29.aspx)|
|[DataChange](http://msdn.microsoft.com/library/form-datachange-property-access%28Office.15%29.aspx)|
|[DataEntry](http://msdn.microsoft.com/library/form-dataentry-property-access%28Office.15%29.aspx)|
|[DataSetChange](http://msdn.microsoft.com/library/form-datasetchange-property-access%28Office.15%29.aspx)|
|[DatasheetAlternateBackColor](http://msdn.microsoft.com/library/form-datasheetalternatebackcolor-property-access%28Office.15%29.aspx)|
|[DatasheetBackColor](http://msdn.microsoft.com/library/form-datasheetbackcolor-property-access%28Office.15%29.aspx)|
|[DatasheetBorderLineStyle](http://msdn.microsoft.com/library/form-datasheetborderlinestyle-property-access%28Office.15%29.aspx)|
|[DatasheetCellsEffect](http://msdn.microsoft.com/library/form-datasheetcellseffect-property-access%28Office.15%29.aspx)|
|[DatasheetColumnHeaderUnderlineStyle](http://msdn.microsoft.com/library/form-datasheetcolumnheaderunderlinestyle-property-access%28Office.15%29.aspx)|
|[DatasheetFontHeight](http://msdn.microsoft.com/library/form-datasheetfontheight-property-access%28Office.15%29.aspx)|
|[DatasheetFontItalic](http://msdn.microsoft.com/library/form-datasheetfontitalic-property-access%28Office.15%29.aspx)|
|[DatasheetFontName](http://msdn.microsoft.com/library/form-datasheetfontname-property-access%28Office.15%29.aspx)|
|[DatasheetFontUnderline](http://msdn.microsoft.com/library/form-datasheetfontunderline-property-access%28Office.15%29.aspx)|
|[DatasheetFontWeight](http://msdn.microsoft.com/library/form-datasheetfontweight-property-access%28Office.15%29.aspx)|
|[DatasheetForeColor](http://msdn.microsoft.com/library/form-datasheetforecolor-property-access%28Office.15%29.aspx)|
|[DatasheetGridlinesBehavior](http://msdn.microsoft.com/library/form-datasheetgridlinesbehavior-property-access%28Office.15%29.aspx)|
|[DatasheetGridlinesColor](http://msdn.microsoft.com/library/form-datasheetgridlinescolor-property-access%28Office.15%29.aspx)|
|[DefaultControl](http://msdn.microsoft.com/library/form-defaultcontrol-property-access%28Office.15%29.aspx)|
|[DefaultView](http://msdn.microsoft.com/library/form-defaultview-property-access%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/form-dirty-property-access%28Office.15%29.aspx)|
|[DisplayOnSharePointSite](http://msdn.microsoft.com/library/form-displayonsharepointsite-property-access%28Office.15%29.aspx)|
|[DividingLines](http://msdn.microsoft.com/library/form-dividinglines-property-access%28Office.15%29.aspx)|
|[FastLaserPrinting](http://msdn.microsoft.com/library/form-fastlaserprinting-property-access%28Office.15%29.aspx)|
|[FetchDefaults](http://msdn.microsoft.com/library/form-fetchdefaults-property-access%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/form-filter-property-access%28Office.15%29.aspx)|
|[FilterOn](http://msdn.microsoft.com/library/form-filteron-property-access%28Office.15%29.aspx)|
|[FilterOnLoad](http://msdn.microsoft.com/library/form-filteronload-property-access%28Office.15%29.aspx)|
|[FitToScreen](http://msdn.microsoft.com/library/form-fittoscreen-property-access%28Office.15%29.aspx)|
|[Form](http://msdn.microsoft.com/library/form-form-property-access%28Office.15%29.aspx)|
|[FrozenColumns](http://msdn.microsoft.com/library/form-frozencolumns-property-access%28Office.15%29.aspx)|
|[GridX](http://msdn.microsoft.com/library/form-gridx-property-access%28Office.15%29.aspx)|
|[GridY](http://msdn.microsoft.com/library/form-gridy-property-access%28Office.15%29.aspx)|
|[HasModule](http://msdn.microsoft.com/library/form-hasmodule-property-access%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/form-helpcontextid-property-access%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/form-helpfile-property-access%28Office.15%29.aspx)|
|[HorizontalDatasheetGridlineStyle](http://msdn.microsoft.com/library/form-horizontaldatasheetgridlinestyle-property-access%28Office.15%29.aspx)|
|[Hwnd](http://msdn.microsoft.com/library/form-hwnd-property-access%28Office.15%29.aspx)|
|[InputParameters](http://msdn.microsoft.com/library/form-inputparameters-property-access%28Office.15%29.aspx)|
|[InsideHeight](http://msdn.microsoft.com/library/form-insideheight-property-access%28Office.15%29.aspx)|
|[InsideWidth](http://msdn.microsoft.com/library/form-insidewidth-property-access%28Office.15%29.aspx)|
|[KeyPreview](http://msdn.microsoft.com/library/form-keypreview-property-access%28Office.15%29.aspx)|
|[LayoutForPrint](http://msdn.microsoft.com/library/form-layoutforprint-property-access%28Office.15%29.aspx)|
|[MaxRecButton](http://msdn.microsoft.com/library/form-maxrecbutton-property-access%28Office.15%29.aspx)|
|[MaxRecords](http://msdn.microsoft.com/library/form-maxrecords-property-access%28Office.15%29.aspx)|
|[MenuBar](http://msdn.microsoft.com/library/form-menubar-property-access%28Office.15%29.aspx)|
|[MinMaxButtons](http://msdn.microsoft.com/library/form-minmaxbuttons-property-access%28Office.15%29.aspx)|
|[Modal](http://msdn.microsoft.com/library/form-modal-property-access%28Office.15%29.aspx)|
|[Module](http://msdn.microsoft.com/library/form-module-property-access%28Office.15%29.aspx)|
|[MouseWheel](http://msdn.microsoft.com/library/form-mousewheel-property-access%28Office.15%29.aspx)|
|[Moveable](http://msdn.microsoft.com/library/form-moveable-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/form-name-property-access%28Office.15%29.aspx)|
|[NavigationButtons](http://msdn.microsoft.com/library/form-navigationbuttons-property-access%28Office.15%29.aspx)|
|[NavigationCaption](http://msdn.microsoft.com/library/form-navigationcaption-property-access%28Office.15%29.aspx)|
|[NewRecord](http://msdn.microsoft.com/library/form-newrecord-property-access%28Office.15%29.aspx)|
|[OnActivate](http://msdn.microsoft.com/library/form-onactivate-property-access%28Office.15%29.aspx)|
|[OnApplyFilter](http://msdn.microsoft.com/library/form-onapplyfilter-property-access%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/form-onclick-property-access%28Office.15%29.aspx)|
|[OnClose](http://msdn.microsoft.com/library/form-onclose-property-access%28Office.15%29.aspx)|
|[OnConnect](http://msdn.microsoft.com/library/form-onconnect-property-access%28Office.15%29.aspx)|
|[OnCurrent](http://msdn.microsoft.com/library/form-oncurrent-property-access%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/form-ondblclick-property-access%28Office.15%29.aspx)|
|[OnDeactivate](http://msdn.microsoft.com/library/form-ondeactivate-property-access%28Office.15%29.aspx)|
|[OnDelete](http://msdn.microsoft.com/library/form-ondelete-property-access%28Office.15%29.aspx)|
|[OnDirty](http://msdn.microsoft.com/library/form-ondirty-property-access%28Office.15%29.aspx)|
|[OnDisconnect](http://msdn.microsoft.com/library/form-ondisconnect-property-access%28Office.15%29.aspx)|
|[OnError](http://msdn.microsoft.com/library/form-onerror-property-access%28Office.15%29.aspx)|
|[OnFilter](http://msdn.microsoft.com/library/form-onfilter-property-access%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/form-ongotfocus-property-access%28Office.15%29.aspx)|
|[OnInsert](http://msdn.microsoft.com/library/form-oninsert-property-access%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/form-onkeydown-property-access%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/form-onkeypress-property-access%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/form-onkeyup-property-access%28Office.15%29.aspx)|
|[OnLoad](http://msdn.microsoft.com/library/form-onload-property-access%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/form-onlostfocus-property-access%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/form-onmousedown-property-access%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/form-onmousemove-property-access%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/form-onmouseup-property-access%28Office.15%29.aspx)|
|[OnOpen](http://msdn.microsoft.com/library/form-onopen-property-access%28Office.15%29.aspx)|
|[OnResize](http://msdn.microsoft.com/library/form-onresize-property-access%28Office.15%29.aspx)|
|[OnTimer](http://msdn.microsoft.com/library/form-ontimer-property-access%28Office.15%29.aspx)|
|[OnUndo](http://msdn.microsoft.com/library/form-onundo-property-access%28Office.15%29.aspx)|
|[OnUnload](http://msdn.microsoft.com/library/form-onunload-property-access%28Office.15%29.aspx)|
|[OpenArgs](http://msdn.microsoft.com/library/form-openargs-property-access%28Office.15%29.aspx)|
|[OrderBy](http://msdn.microsoft.com/library/form-orderby-property-access%28Office.15%29.aspx)|
|[OrderByOn](http://msdn.microsoft.com/library/form-orderbyon-property-access%28Office.15%29.aspx)|
|[OrderByOnLoad](http://msdn.microsoft.com/library/form-orderbyonload-property-access%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/form-orientation-property-access%28Office.15%29.aspx)|
|[Page](http://msdn.microsoft.com/library/form-page-property-access%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/form-pages-property-access%28Office.15%29.aspx)|
|[Painting](http://msdn.microsoft.com/library/form-painting-property-access%28Office.15%29.aspx)|
|[PaintPalette](http://msdn.microsoft.com/library/form-paintpalette-property-access%28Office.15%29.aspx)|
|[PaletteSource](http://msdn.microsoft.com/library/form-palettesource-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/form-parent-property-access%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/form-picture-property-access%28Office.15%29.aspx)|
|[PictureAlignment](http://msdn.microsoft.com/library/form-picturealignment-property-access%28Office.15%29.aspx)|
|[PictureData](http://msdn.microsoft.com/library/form-picturedata-property-access%28Office.15%29.aspx)|
|[PicturePalette](http://msdn.microsoft.com/library/form-picturepalette-property-access%28Office.15%29.aspx)|
|[PictureSizeMode](http://msdn.microsoft.com/library/form-picturesizemode-property-access%28Office.15%29.aspx)|
|[PictureTiling](http://msdn.microsoft.com/library/form-picturetiling-property-access%28Office.15%29.aspx)|
|[PictureType](http://msdn.microsoft.com/library/form-picturetype-property-access%28Office.15%29.aspx)|
|[PivotTable](http://msdn.microsoft.com/library/form-pivottable-property-access%28Office.15%29.aspx)|
|[PivotTableChange](http://msdn.microsoft.com/library/form-pivottablechange-property-access%28Office.15%29.aspx)|
|[PopUp](http://msdn.microsoft.com/library/form-popup-property-access%28Office.15%29.aspx)|
|[Printer](http://msdn.microsoft.com/library/form-printer-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/form-properties-property-access%28Office.15%29.aspx)|
|[PrtDevMode](http://msdn.microsoft.com/library/form-prtdevmode-property-access%28Office.15%29.aspx)|
|[PrtDevNames](http://msdn.microsoft.com/library/form-prtdevnames-property-access%28Office.15%29.aspx)|
|[PrtMip](http://msdn.microsoft.com/library/form-prtmip-property-access%28Office.15%29.aspx)|
|[Query](http://msdn.microsoft.com/library/form-query-property-access%28Office.15%29.aspx)|
|[RecordLocks](http://msdn.microsoft.com/library/form-recordlocks-property-access%28Office.15%29.aspx)|
|[RecordSelectors](http://msdn.microsoft.com/library/form-recordselectors-property-access%28Office.15%29.aspx)|
|[Recordset](http://msdn.microsoft.com/library/form-recordset-property-access%28Office.15%29.aspx)|
|[RecordsetClone](http://msdn.microsoft.com/library/form-recordsetclone-property-access%28Office.15%29.aspx)|
|[RecordsetType](http://msdn.microsoft.com/library/form-recordsettype-property-access%28Office.15%29.aspx)|
|[RecordSource](http://msdn.microsoft.com/library/form-recordsource-property-access%28Office.15%29.aspx)|
|[RecordSourceQualifier](http://msdn.microsoft.com/library/form-recordsourcequalifier-property-access%28Office.15%29.aspx)|
|[ResyncCommand](http://msdn.microsoft.com/library/form-resynccommand-property-access%28Office.15%29.aspx)|
|[RibbonName](http://msdn.microsoft.com/library/form-ribbonname-property-access%28Office.15%29.aspx)|
|[RowHeight](http://msdn.microsoft.com/library/form-rowheight-property-access%28Office.15%29.aspx)|
|[ScrollBars](http://msdn.microsoft.com/library/form-scrollbars-property-access%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/form-section-property-access%28Office.15%29.aspx)|
|[SelectionChange](http://msdn.microsoft.com/library/form-selectionchange-property-access%28Office.15%29.aspx)|
|[SelHeight](http://msdn.microsoft.com/library/form-selheight-property-access%28Office.15%29.aspx)|
|[SelLeft](http://msdn.microsoft.com/library/form-selleft-property-access%28Office.15%29.aspx)|
|[SelTop](http://msdn.microsoft.com/library/form-seltop-property-access%28Office.15%29.aspx)|
|[SelWidth](http://msdn.microsoft.com/library/form-selwidth-property-access%28Office.15%29.aspx)|
|[ServerFilter](http://msdn.microsoft.com/library/form-serverfilter-property-access%28Office.15%29.aspx)|
|[ServerFilterByForm](http://msdn.microsoft.com/library/form-serverfilterbyform-property-access%28Office.15%29.aspx)|
|[ShortcutMenu](http://msdn.microsoft.com/library/form-shortcutmenu-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/form-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[SplitFormDatasheet](http://msdn.microsoft.com/library/form-splitformdatasheet-property-access%28Office.15%29.aspx)|
|[SplitFormOrientation](http://msdn.microsoft.com/library/form-splitformorientation-property-access%28Office.15%29.aspx)|
|[SplitFormPrinting](http://msdn.microsoft.com/library/form-splitformprinting-property-access%28Office.15%29.aspx)|
|[SplitFormSize](http://msdn.microsoft.com/library/form-splitformsize-property-access%28Office.15%29.aspx)|
|[SplitFormSplitterBar](http://msdn.microsoft.com/library/form-splitformsplitterbar-property-access%28Office.15%29.aspx)|
|[SplitFormSplitterBarSave](http://msdn.microsoft.com/library/form-splitformsplitterbarsave-property-access%28Office.15%29.aspx)|
|[SubdatasheetExpanded](http://msdn.microsoft.com/library/form-subdatasheetexpanded-property-access%28Office.15%29.aspx)|
|[SubdatasheetHeight](http://msdn.microsoft.com/library/form-subdatasheetheight-property-access%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/form-tag-property-access%28Office.15%29.aspx)|
|[TimerInterval](http://msdn.microsoft.com/library/form-timerinterval-property-access%28Office.15%29.aspx)|
|[Toolbar](http://msdn.microsoft.com/library/form-toolbar-property-access%28Office.15%29.aspx)|
|[UniqueTable](http://msdn.microsoft.com/library/form-uniquetable-property-access%28Office.15%29.aspx)|
|[UseDefaultPrinter](http://msdn.microsoft.com/library/form-usedefaultprinter-property-access%28Office.15%29.aspx)|
|[VerticalDatasheetGridlineStyle](http://msdn.microsoft.com/library/form-verticaldatasheetgridlinestyle-property-access%28Office.15%29.aspx)|
|[ViewChange](http://msdn.microsoft.com/library/form-viewchange-property-access%28Office.15%29.aspx)|
|[ViewsAllowed](http://msdn.microsoft.com/library/form-viewsallowed-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/form-visible-property-access%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/form-width-property-access%28Office.15%29.aspx)|
|[WindowHeight](http://msdn.microsoft.com/library/form-windowheight-property-access%28Office.15%29.aspx)|
|[WindowLeft](http://msdn.microsoft.com/library/form-windowleft-property-access%28Office.15%29.aspx)|
|[WindowTop](http://msdn.microsoft.com/library/form-windowtop-property-access%28Office.15%29.aspx)|
|[WindowWidth](http://msdn.microsoft.com/library/form-windowwidth-property-access%28Office.15%29.aspx)|

## About the Contributors
<a name="AboutContributors"> </a>

Luke Chung is the founder and president of FMS, Inc., a leading provider of custom database solutions and developer tools. 

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


<<<<<<< HEAD
#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
=======
#### Concepts


[Access Object Model Reference](object-model-access-vba-reference.md)
>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

