---
title: Report Object (Access)
keywords: vbaac10.chm13901
f1_keywords:
- vbaac10.chm13901
ms.prod: ACCESS
api_name:
- Access.Report
ms.assetid: 6f77c1b4-a9ce-7caa-204c-fe0755c6f9df
---


# Report Object (Access)

A  **Report** object refers to a particular Microsoft Access report.


## Remarks

A  **Report** object is a member of the **Reports** collection, which is a collection of all currently open reports. Within the **Reports** collection, individual reports are indexed beginning with zero. You can refer to an individual **Report** object in the **Reports** collection either by referring to the report by name, or by referring to its index within the collection. If the report name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**Reports** ! _reportname_|Reports!OrderReport|
|**Reports** ![ _report name_]|Reports![Order Report]|
|**Reports** (" _reportname_")|Reports("OrderReport")|
|**Reports** ( _index_)|Reports(0)|

 **Note**  Each  **Report** object has a **Controls** collection, which contains all controls on the report. You can refer to a control on a report either by implicitly or explicitly referring to the **Controls** collection. Your code will be faster if you refer to the **Controls** collection implicitly. The following examples show two of the ways you might refer to a control named **NewData** on a report called **OrderReport**. 


```
' Implicit reference. 
Reports!OrderReport!NewData
```


```
' Explicit reference. 
Reports!OrderReport.Controls!NewData
```


## Example

The following example shows how to use the  **NoData** event of a report to prevent the report form opening when there is no data to be displayed.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer?s Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.mdl)




```
Private Sub Report_NoData(Cancel As Integer)

    'Add code here that will be executed if no data
    'was returned by the Report's RecordSource
    MsgBox "No customers ordered this product this month. " &amp; _
        "The report will now close."
    Cancel = True

End Sub
```

The following example shows how to use the  **Page** event to add a watermark to a report before it is printed.




```
Private Sub Report_Page()
    Dim strWatermarkText As String
    Dim sizeHor As Single
    Dim sizeVer As Single

#If RUN_PAGE_EVENT = True Then
    With Me
        '// Print page border
        Me.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vbBlack, B
    
        '// Print watermark
        strWatermarkText = "Confidential"
        
        .ScaleMode = 3
        .FontName = "Segoe UI"
        .FontSize = 48
        .ForeColor = RGB(255, 0, 0)

        '// Calculate text metrics
        sizeHor = .TextWidth(strWatermarkText)
        sizeVer = .TextHeight(strWatermarkText)
        
        '// Set the print location
        .CurrentX = (.ScaleWidth / 2) - (sizeHor / 2)
        .CurrentY = (.ScaleHeight / 2) - (sizeVer / 2)
    
        '// Print the watermark
        .Print strWatermarkText
    End With
#End If

End Sub
```

The following example shows how to set the  **BackColor** property of a control based on its value.




```
Private Sub SetControlFormatting()
    If (Me.AvgOfRating >= 8) Then
        Me.AvgOfRating.BackColor = vbGreen
    ElseIf (Me.AvgOfRating >= 5) Then
        Me.AvgOfRating.BackColor = vbYellow
    Else
        Me.AvgOfRating.BackColor = vbRed
    End If
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' size the width of the rectangle
    Dim lngOffset As Long
    lngOffset = (Me.boxInside.Left - Me.boxOutside.Left) * 2
    Me.boxInside.Width = (Me.boxOutside.Width * (Me.AvgOfRating / 10)) - lngOffset
    
    ' do conditional formatting for the control in print preview
    SetControlFormatting
End Sub

Private Sub Detail_Paint()
    ' do conditional formatting for the control in report view
    SetControlFormatting
End Sub
```

The following example shows how to format a report to show progress bars. The example uses a pair of rectangle controls,  **boxInside** and **boxOutside**, to create a progress bar based on the value of  **AvgOfRating**. The progress bars are visible only when the report is opened in  **Print Preview** mode or it is printed.




```
Private Sub Report_Load()
    If (Me.CurrentView = AcCurrentView.acCurViewPreview) Then
        Me.boxInside.Visible = True
        Me.boxOutside.Visible = True
    Else
        Me.boxInside.Visible = False
        Me.boxOutside.Visible = False
    End If
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' size the width of the rectangle
    Dim lngOffset As Long
    lngOffset = (Me.boxInside.Left - Me.boxOutside.Left) * 2
    Me.boxInside.Width = (Me.boxOutside.Width * (Me.AvgOfRating / 10)) - lngOffset
    
    ' do conditional formatting for the control in print preview
    SetControlFormatting
End Sub
```


## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/report-activate-event-access%28Office.15%29.aspx)|
|[ApplyFilter](http://msdn.microsoft.com/library/report-applyfilter-event-access%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/report-click-event-access%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/report-close-event-access%28Office.15%29.aspx)|
|[Current](http://msdn.microsoft.com/library/report-current-event-access%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/report-dblclick-event-access%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/report-deactivate-event-access%28Office.15%29.aspx)|
|[Error](http://msdn.microsoft.com/library/report-error-event-access%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/report-filter-event-access%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/report-gotfocus-event-access%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/report-keydown-event-access%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/report-keypress-event-access%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/report-keyup-event-access%28Office.15%29.aspx)|
|[Load](http://msdn.microsoft.com/library/report-load-event-access%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/report-lostfocus-event-access%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/report-mousedown-event-access%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/report-mousemove-event-access%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/report-mouseup-event-access%28Office.15%29.aspx)|
|[MouseWheel](http://msdn.microsoft.com/library/report-mousewheel-event-access%28Office.15%29.aspx)|
|[NoData](http://msdn.microsoft.com/library/report-nodata-event-access%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/report-open-event-access%28Office.15%29.aspx)|
|[Page](http://msdn.microsoft.com/library/report-page-event-access%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/report-resize-event-access%28Office.15%29.aspx)|
|[Timer](http://msdn.microsoft.com/library/report-timer-event-access%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/report-unload-event-access%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Circle](http://msdn.microsoft.com/library/report-circle-method-access%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/report-line-method-access%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/report-move-method-access%28Office.15%29.aspx)|
|[Print](http://msdn.microsoft.com/library/report-print-method-access%28Office.15%29.aspx)|
|[PSet](http://msdn.microsoft.com/library/report-pset-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/report-requery-method-access%28Office.15%29.aspx)|
|[Scale](http://msdn.microsoft.com/library/report-scale-method-access%28Office.15%29.aspx)|
|[TextHeight](http://msdn.microsoft.com/library/report-textheight-method-access%28Office.15%29.aspx)|
|[TextWidth](http://msdn.microsoft.com/library/report-textwidth-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveControl](http://msdn.microsoft.com/library/report-activecontrol-property-access%28Office.15%29.aspx)|
|[AllowLayoutView](http://msdn.microsoft.com/library/report-allowlayoutview-property-access%28Office.15%29.aspx)|
|[AllowReportView](http://msdn.microsoft.com/library/report-allowreportview-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/report-application-property-access%28Office.15%29.aspx)|
|[AutoCenter](http://msdn.microsoft.com/library/report-autocenter-property-access%28Office.15%29.aspx)|
|[AutoResize](http://msdn.microsoft.com/library/report-autoresize-property-access%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/report-borderstyle-property-access%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/report-caption-property-access%28Office.15%29.aspx)|
|[CloseButton](http://msdn.microsoft.com/library/report-closebutton-property-access%28Office.15%29.aspx)|
|[ControlBox](http://msdn.microsoft.com/library/report-controlbox-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/report-controls-property-access%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/report-count-property-access%28Office.15%29.aspx)|
|[CurrentRecord](http://msdn.microsoft.com/library/report-currentrecord-property-access%28Office.15%29.aspx)|
|[CurrentView](http://msdn.microsoft.com/library/report-currentview-property-access%28Office.15%29.aspx)|
|[CurrentX](http://msdn.microsoft.com/library/report-currentx-property-access%28Office.15%29.aspx)|
|[CurrentY](http://msdn.microsoft.com/library/report-currenty-property-access%28Office.15%29.aspx)|
|[Cycle](http://msdn.microsoft.com/library/report-cycle-property-access%28Office.15%29.aspx)|
|[DateGrouping](http://msdn.microsoft.com/library/report-dategrouping-property-access%28Office.15%29.aspx)|
|[DefaultControl](http://msdn.microsoft.com/library/report-defaultcontrol-property-access%28Office.15%29.aspx)|
|[DefaultView](http://msdn.microsoft.com/library/report-defaultview-property-access%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/report-dirty-property-access%28Office.15%29.aspx)|
|[DisplayOnSharePointSite](http://msdn.microsoft.com/library/report-displayonsharepointsite-property-access%28Office.15%29.aspx)|
|[DrawMode](http://msdn.microsoft.com/library/report-drawmode-property-access%28Office.15%29.aspx)|
|[DrawStyle](http://msdn.microsoft.com/library/report-drawstyle-property-access%28Office.15%29.aspx)|
|[DrawWidth](http://msdn.microsoft.com/library/report-drawwidth-property-access%28Office.15%29.aspx)|
|[FastLaserPrinting](http://msdn.microsoft.com/library/report-fastlaserprinting-property-access%28Office.15%29.aspx)|
|[FillColor](http://msdn.microsoft.com/library/report-fillcolor-property-access%28Office.15%29.aspx)|
|[FillStyle](http://msdn.microsoft.com/library/report-fillstyle-property-access%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/report-filter-property-access%28Office.15%29.aspx)|
|[FilterOn](http://msdn.microsoft.com/library/report-filteron-property-access%28Office.15%29.aspx)|
|[FilterOnLoad](http://msdn.microsoft.com/library/report-filteronload-property-access%28Office.15%29.aspx)|
|[FitToPage](http://msdn.microsoft.com/library/report-fittopage-property-access%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/report-fontbold-property-access%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/report-fontitalic-property-access%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/report-fontname-property-access%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/report-fontsize-property-access%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/report-fontunderline-property-access%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/report-forecolor-property-access%28Office.15%29.aspx)|
|[FormatCount](http://msdn.microsoft.com/library/report-formatcount-property-access%28Office.15%29.aspx)|
|[GridX](http://msdn.microsoft.com/library/report-gridx-property-access%28Office.15%29.aspx)|
|[GridY](http://msdn.microsoft.com/library/report-gridy-property-access%28Office.15%29.aspx)|
|[GroupLevel](http://msdn.microsoft.com/library/report-grouplevel-property-access%28Office.15%29.aspx)|
|[GrpKeepTogether](http://msdn.microsoft.com/library/report-grpkeeptogether-property-access%28Office.15%29.aspx)|
|[HasData](http://msdn.microsoft.com/library/report-hasdata-property-access%28Office.15%29.aspx)|
|[HasModule](http://msdn.microsoft.com/library/report-hasmodule-property-access%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/report-height-property-access%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/report-helpcontextid-property-access%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/report-helpfile-property-access%28Office.15%29.aspx)|
|[Hwnd](http://msdn.microsoft.com/library/report-hwnd-property-access%28Office.15%29.aspx)|
|[InputParameters](http://msdn.microsoft.com/library/report-inputparameters-property-access%28Office.15%29.aspx)|
|[KeyPreview](http://msdn.microsoft.com/library/report-keypreview-property-access%28Office.15%29.aspx)|
|[LayoutForPrint](http://msdn.microsoft.com/library/report-layoutforprint-property-access%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/report-left-property-access%28Office.15%29.aspx)|
|[MenuBar](http://msdn.microsoft.com/library/report-menubar-property-access%28Office.15%29.aspx)|
|[MinMaxButtons](http://msdn.microsoft.com/library/report-minmaxbuttons-property-access%28Office.15%29.aspx)|
|[Modal](http://msdn.microsoft.com/library/report-modal-property-access%28Office.15%29.aspx)|
|[Module](http://msdn.microsoft.com/library/report-module-property-access%28Office.15%29.aspx)|
|[MouseWheel](http://msdn.microsoft.com/library/report-mousewheel-property-access%28Office.15%29.aspx)|
|[Moveable](http://msdn.microsoft.com/library/report-moveable-property-access%28Office.15%29.aspx)|
|[MoveLayout](http://msdn.microsoft.com/library/report-movelayout-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/report-name-property-access%28Office.15%29.aspx)|
|[NextRecord](http://msdn.microsoft.com/library/report-nextrecord-property-access%28Office.15%29.aspx)|
|[OnActivate](http://msdn.microsoft.com/library/report-onactivate-property-access%28Office.15%29.aspx)|
|[OnApplyFilter](http://msdn.microsoft.com/library/report-onapplyfilter-property-access%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/report-onclick-property-access%28Office.15%29.aspx)|
|[OnClose](http://msdn.microsoft.com/library/report-onclose-property-access%28Office.15%29.aspx)|
|[OnCurrent](http://msdn.microsoft.com/library/report-oncurrent-property-access%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/report-ondblclick-property-access%28Office.15%29.aspx)|
|[OnDeactivate](http://msdn.microsoft.com/library/report-ondeactivate-property-access%28Office.15%29.aspx)|
|[OnError](http://msdn.microsoft.com/library/report-onerror-property-access%28Office.15%29.aspx)|
|[OnFilter](http://msdn.microsoft.com/library/report-onfilter-property-access%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/report-ongotfocus-property-access%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/report-onkeydown-property-access%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/report-onkeypress-property-access%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/report-onkeyup-property-access%28Office.15%29.aspx)|
|[OnLoad](http://msdn.microsoft.com/library/report-onload-property-access%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/report-onlostfocus-property-access%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/report-onmousedown-property-access%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/report-onmousemove-property-access%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/report-onmouseup-property-access%28Office.15%29.aspx)|
|[OnNoData](http://msdn.microsoft.com/library/report-onnodata-property-access%28Office.15%29.aspx)|
|[OnOpen](http://msdn.microsoft.com/library/report-onopen-property-access%28Office.15%29.aspx)|
|[OnPage](http://msdn.microsoft.com/library/report-onpage-property-access%28Office.15%29.aspx)|
|[OnResize](http://msdn.microsoft.com/library/report-onresize-property-access%28Office.15%29.aspx)|
|[OnTimer](http://msdn.microsoft.com/library/report-ontimer-property-access%28Office.15%29.aspx)|
|[OnUnload](http://msdn.microsoft.com/library/report-onunload-property-access%28Office.15%29.aspx)|
|[OpenArgs](http://msdn.microsoft.com/library/report-openargs-property-access%28Office.15%29.aspx)|
|[OrderBy](http://msdn.microsoft.com/library/report-orderby-property-access%28Office.15%29.aspx)|
|[OrderByOn](http://msdn.microsoft.com/library/report-orderbyon-property-access%28Office.15%29.aspx)|
|[OrderByOnLoad](http://msdn.microsoft.com/library/report-orderbyonload-property-access%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/report-orientation-property-access%28Office.15%29.aspx)|
|[Page](http://msdn.microsoft.com/library/report-page-property-access%28Office.15%29.aspx)|
|[PageFooter](http://msdn.microsoft.com/library/report-pagefooter-property-access%28Office.15%29.aspx)|
|[PageHeader](http://msdn.microsoft.com/library/report-pageheader-property-access%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/report-pages-property-access%28Office.15%29.aspx)|
|[Painting](http://msdn.microsoft.com/library/report-painting-property-access%28Office.15%29.aspx)|
|[PaintPalette](http://msdn.microsoft.com/library/report-paintpalette-property-access%28Office.15%29.aspx)|
|[PaletteSource](http://msdn.microsoft.com/library/report-palettesource-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/report-parent-property-access%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/report-picture-property-access%28Office.15%29.aspx)|
|[PictureAlignment](http://msdn.microsoft.com/library/report-picturealignment-property-access%28Office.15%29.aspx)|
|[PictureData](http://msdn.microsoft.com/library/report-picturedata-property-access%28Office.15%29.aspx)|
|[PicturePages](http://msdn.microsoft.com/library/report-picturepages-property-access%28Office.15%29.aspx)|
|[PicturePalette](http://msdn.microsoft.com/library/report-picturepalette-property-access%28Office.15%29.aspx)|
|[PictureSizeMode](http://msdn.microsoft.com/library/report-picturesizemode-property-access%28Office.15%29.aspx)|
|[PictureTiling](http://msdn.microsoft.com/library/report-picturetiling-property-access%28Office.15%29.aspx)|
|[PictureType](http://msdn.microsoft.com/library/report-picturetype-property-access%28Office.15%29.aspx)|
|[PopUp](http://msdn.microsoft.com/library/report-popup-property-access%28Office.15%29.aspx)|
|[PrintCount](http://msdn.microsoft.com/library/report-printcount-property-access%28Office.15%29.aspx)|
|[Printer](http://msdn.microsoft.com/library/report-printer-property-access%28Office.15%29.aspx)|
|[PrintSection](http://msdn.microsoft.com/library/report-printsection-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/report-properties-property-access%28Office.15%29.aspx)|
|[PrtDevMode](http://msdn.microsoft.com/library/report-prtdevmode-property-access%28Office.15%29.aspx)|
|[PrtDevNames](http://msdn.microsoft.com/library/report-prtdevnames-property-access%28Office.15%29.aspx)|
|[PrtMip](http://msdn.microsoft.com/library/report-prtmip-property-access%28Office.15%29.aspx)|
|[RecordLocks](http://msdn.microsoft.com/library/report-recordlocks-property-access%28Office.15%29.aspx)|
|[Recordset](http://msdn.microsoft.com/library/report-recordset-property-access%28Office.15%29.aspx)|
|[RecordSource](http://msdn.microsoft.com/library/report-recordsource-property-access%28Office.15%29.aspx)|
|[RecordSourceQualifier](http://msdn.microsoft.com/library/report-recordsourcequalifier-property-access%28Office.15%29.aspx)|
|[Report](http://msdn.microsoft.com/library/report-report-property-access%28Office.15%29.aspx)|
|[RibbonName](http://msdn.microsoft.com/library/report-ribbonname-property-access%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/report-scaleheight-property-access%28Office.15%29.aspx)|
|[ScaleLeft](http://msdn.microsoft.com/library/report-scaleleft-property-access%28Office.15%29.aspx)|
|[ScaleMode](http://msdn.microsoft.com/library/report-scalemode-property-access%28Office.15%29.aspx)|
|[ScaleTop](http://msdn.microsoft.com/library/report-scaletop-property-access%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/report-scalewidth-property-access%28Office.15%29.aspx)|
|[ScrollBars](http://msdn.microsoft.com/library/report-scrollbars-property-access%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/report-section-property-access%28Office.15%29.aspx)|
|[ServerFilter](http://msdn.microsoft.com/library/report-serverfilter-property-access%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/report-shape-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/report-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[ShowPageMargins](http://msdn.microsoft.com/library/report-showpagemargins-property-access%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/report-tag-property-access%28Office.15%29.aspx)|
|[TimerInterval](http://msdn.microsoft.com/library/report-timerinterval-property-access%28Office.15%29.aspx)|
|[Toolbar](http://msdn.microsoft.com/library/report-toolbar-property-access%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/report-top-property-access%28Office.15%29.aspx)|
|[UseDefaultPrinter](http://msdn.microsoft.com/library/report-usedefaultprinter-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/report-visible-property-access%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/report-width-property-access%28Office.15%29.aspx)|
|[WindowHeight](http://msdn.microsoft.com/library/report-windowheight-property-access%28Office.15%29.aspx)|
|[WindowLeft](http://msdn.microsoft.com/library/report-windowleft-property-access%28Office.15%29.aspx)|
|[WindowTop](http://msdn.microsoft.com/library/report-windowtop-property-access%28Office.15%29.aspx)|
|[WindowWidth](http://msdn.microsoft.com/library/report-windowwidth-property-access%28Office.15%29.aspx)|

## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Other resources


[Report Object Members](http://msdn.microsoft.com/library/report-members-access%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
