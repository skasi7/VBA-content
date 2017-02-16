---
title: Report Members (Access)
ms.prod: ACCESS
ms.assetid: 73370a33-1ca0-da4d-9e36-88011bc2b93e
---


# Report Members (Access)


A  **Report** object refers to a particular Microsoft Access report.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Activate](report-activate-event-access.md)|The Activate event occurs when a report receives the focus and becomes the active window.|
|[ApplyFilter](report-applyfilter-event-access.md)|Occurs when a filter is applied to a report.|
|[Click](report-click-event-access.md)|The  **Click** event occurs when the user presses and then releases a mouse button over a report.|
|[Close](report-close-event-access.md)|The  **Close** event occurs when a report is closed and removed from the screen.|
|[Current](report-current-event-access.md)|Occurs when the focus moves to a record, making it the current record, or when the report is refreshed or requeried.|
|[DblClick](report-dblclick-event-access.md)|The  **DblClick** event occurs when the user presses and releases the left mouse button twice over an report within the double-click time limit of the system.|
|[Deactivate](report-deactivate-event-access.md)|The  **Deactivate** event occurs when a report loses the focus to a Table, Query, Form, Report, Macro, or Module window, or to the Database window.|
|[Error](report-error-event-access.md)|The Error event occurs when a run-time error is produced in Microsoft Access when a report has the focus.|
|[Filter](report-filter-event-access.md)|Occurs when the user opens a filter window by clicking  **Advanced Filter/Sort**.|
|[GotFocus](report-gotfocus-event-access.md)|The  **GotFocus** event occurs when the report receives the focus.|
|[KeyDown](report-keydown-event-access.md)|The  **KeyDown** event occurs when the user presses a key while a report has the focus. This event also occurs if you send a keystroke to a report by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyPress](report-keypress-event-access.md)|The  **KeyPress** event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a report has the focus. This event also occurs if you send an ANSI keystroke to a report by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyUp](report-keyup-event-access.md)|The  **KeyUp** event occurs when the user releases a key while a report has the focus. This event also occurs if you send a keystroke to a report by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[Load](report-load-event-access.md)|Occurs when a report is opened and its records are displayed.|
|[LostFocus](report-lostfocus-event-access.md)|The  **LostFocus** event occurs when the specified object loses the focus.|
|[MouseDown](report-mousedown-event-access.md)|The  **MouseDown** event occurs when the user presses a mouse button.|
|[MouseMove](report-mousemove-event-access.md)|The  **MouseMove** event occurs when the user moves the mouse.|
|[MouseUp](report-mouseup-event-access.md)|The  **MouseUp** event occurs when the user releases a mouse button.|
|[MouseWheel](report-mousewheel-event-access.md)|Occurs when the user rolls the mouse wheel in Report view or Layout view.|
|[NoData](report-nodata-event-access.md)|The  **NoData** event occurs after Microsoft Access formats a report for printing that has no data (the report is bound to an empty recordset), but before the report is printed. You can use this event to cancel printing of a blank report.|
|[Open](report-open-event-access.md)|The  **Open** occurs before a report is previewed or printed.|
|[Page](report-page-event-access.md)|The  **Page** event occurs after Microsoft Access formats a page of a report for printing, but before the page is printed. You can use this event to draw a border around the page, or add other graphic elements to the page.|
|[Resize](report-resize-event-access.md)|The  **Resize** event occurs when a report is opened and whenever the size of a report changes.|
|[Timer](report-timer-event-access.md)|The  **Timer** event occurs for a report at regular intervals as specified by the report's **[TimerInterval](report-timerinterval-property-access.md)** property.|
|[Unload](report-unload-event-access.md)|The  **Unload** event occurs after a report is closed but before it's removed from the screen.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Circle](report-circle-method-access.md)|The  **Circle** method draws a circle, an ellipse, or an arc on a **Report** object when the Print event occurs.|
|[Line](report-line-method-access.md)|The  **Line** method draws lines and rectangles on a **Report** object when the Print event occurs.|
|[Move](report-move-method-access.md)|Moves the specified object to the coordinates specified by the argument values.|
|[Print](report-print-method-access.md)|The  **Print** method prints text on a **[Report](report-object-access.md)** object using the current color and font.|
|[PSet](report-pset-method-access.md)|The  **PSet** method sets a point on a **[Report](report-object-access.md)** object to a specified color when the **Print** event occurs.|
|[Requery](report-requery-method-access.md)|The  **Requery** method updates the data underlying the specified report by requerying the source of data for the control.|
|[Scale](report-scale-method-access.md)|The  **Scale** method defines the coordinate system for a **[Report](report-object-access.md)** object.|
|[TextHeight](report-textheight-method-access.md)|The  **TextHeight** method returns the height of a text string as it would be printed in the current font of a **[Report](report-object-access.md)** object.|
|[TextWidth](report-textwidth-method-access.md)|The  **TextWidth** method returns the width of a text string as it would be printed in the current font of a **[Report](report-object-access.md)** object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActiveControl](report-activecontrol-property-access.md)|You can use the  **ActiveControl** property together with the **[Screen](screen-object-access.md)** object to identify or refer to the control that has the focus. Read-only **Control** object.|
|[AllowLayoutView](report-allowlayoutview-property-access.md)|Gets or sets whether the specified report can be used in Layout View. Read/write  **Boolean**.|
|[AllowReportView](report-allowreportview-property-access.md)|Gets or sets whether the user is allowed to enter Report view while using the specified report. Read/write  **Boolean**.|
|[Application](report-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.|
|[AutoCenter](report-autocenter-property-access.md)|Returns or sets a  **Boolean** indicating whether a report will be centered automatically in the application window when the form is opened. Read/write.|
|[AutoResize](report-autoresize-property-access.md)|Returns or sets a  **Boolean** indicating whether a Report window opens automatically sized to display complete records. Read/write.|
|[BorderStyle](report-borderstyle-property-access.md)|Specifies how a control's border appears.Read/write  **Byte**.|
|[Caption](report-caption-property-access.md)|Gets or sets the title of the report in Print Preview. Read/write  **String**.|
|[CloseButton](report-closebutton-property-access.md)|Specifies whether the  **Close** button on a form is enabled. Read/write **Boolean**.|
|[ControlBox](report-controlbox-property-access.md)|Specifies whether a report has a  **Control** menu in Report view. Read/write **Boolean**.|
|[Controls](report-controls-property-access.md)|Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.|
|[Count](report-count-property-access.md)|You can use the  **Count** property to determine the number of items in a specified collection. Read-only **Integer**.|
|[CurrentRecord](report-currentrecord-property-access.md)|You can use the  **CurrentRecord** property to identify the current record in the recordset being viewed. Read/write **Long**.|
|[CurrentView](report-currentview-property-access.md)|You can use the  **CurrentView** property to determine how a report is currently displayed. Read/write **Integer**.|
|[CurrentX](report-currentx-property-access.md)|You can use the  **CurrentX** property (along with the **CurrentY** property) to specify the horizontal and vertical coordinates for the starting position of the next printing and drawing method on a report. Read/write **Single**.|
|[CurrentY](report-currenty-property-access.md)|You can use the  **CurrentY** property (along with the **CurrentX** property) to specify the horizontal and vertical coordinates for the starting position of the next printing and drawing method on a report. Read/write **Single**.|
|[Cycle](report-cycle-property-access.md)|You can use the  **Cycle** property to specify what happens when you press the TAB key and the focus is in the last control on a report. Read/write **Byte**.|
|[DateGrouping](report-dategrouping-property-access.md)|You can use the  **DateGrouping** property to specify how you want to group dates in a report. Read/write **Byte**.|
|[DefaultControl](report-defaultcontrol-property-access.md)|The  **DefaultControl** property returns a **[Control](control-object-access.md)** object with which you can set the default properties for a particular type of control on a particular report. Read-only.|
|[DefaultView](report-defaultview-property-access.md)|You can use the  **DefaultView** property to specify the opening view of a report. Read/write **Byte**.|
|[Dirty](report-dirty-property-access.md)|You can use the  **Dirty** property to determine whether the current record has been modified since it was last saved. Read/write **Boolean**.|
|[DisplayOnSharePointSite](report-displayonsharepointsite-property-access.md)|Gets or sets whether the specified report can be made available as a view on a Microsoft SharePoint Foundation site. Read/write  **Byte**.|
|[DrawMode](report-drawmode-property-access.md)|You can use the  **DrawMode** property to specify how the pen (the color used in drawing) interacts with existing background colors on a report when the **[Line](report-line-method-access.md)**, **[Circle](report-circle-method-access.md)**, or **[Pset](report-pset-method-access.md)** method is used to draw on a report when printing. Read/write **Integer**.|
|[DrawStyle](report-drawstyle-property-access.md)|You can use the  **DrawStyle** property to specify the line style when using the **[Line](report-line-method-access.md)** and **[Circle](report-circle-method-access.md)** methods to print lines on reports. Read/write **Integer**.|
|[DrawWidth](report-drawwidth-property-access.md)|You can use the  **DrawWidth** property to specify the line width for the **[Line](report-line-method-access.md)**, **[Circle](report-circle-method-access.md)**, and **[Pset](report-pset-method-access.md)** methods to print lines on reports. Read/write **Integer**.|
|[FastLaserPrinting](report-fastlaserprinting-property-access.md)|You can use the  **FastLaserPrinting** property to specify whether lines and rectangles are replaced by text character lines — similar to the underscore (_) and vertical bar (|) characters — when you print a report using most laser printers. Replacing lines and rectangles with text character lines can make printing much faster. Read/write **Boolean**.|
|[FillColor](report-fillcolor-property-access.md)|You use the  **FillColor** property to specify the color that fills in boxes and circles drawn on reports with the **[Line](report-line-method-access.md)** and **[Circle](report-circle-method-access.md)** methods. You can also use this property with[Visual Basic](set-properties-by-using-visual-basic.md)to create special visual effects on custom reports when you print using a color printer or preview the reports on a color monitor. Read/write  **Long**.|
|[FillStyle](report-fillstyle-property-access.md)|You can use the  **FillStyle** property to specify whether a circle or line drawn by the **[Circle](report-circle-method-access.md)** or **[Line](report-line-method-access.md)** method on a report is transparent, opaque, or filled with a pattern. Read/write **Integer**.|
|[Filter](report-filter-property-access.md)|You can use the  **Filter** property to specify a subset of records to be displayed when a filter is applied to a form, reportquery, or table. Read/write **String**.|
|[FilterOn](report-filteron-property-access.md)|You can use the  **FilterOn** property to specify or determine whether the **Filter** property for a form or report is applied. Read/write **Boolean**.|
|[FilterOnLoad](report-filteronload-property-access.md)|Gets or sets whether the filter specified by the  **[Filter](report-filter-property-access.md)** property is applied when the report is loaded. Read/write **Boolean**.|
|[FitToPage](report-fittopage-property-access.md)|Gets or sets whether the width of the specified report is sized to automatically fit the page. Read/write  **Boolean**.|
|[FontBold](report-fontbold-property-access.md)|You can use the  **FontBold** property to specify whether a font appears in a bold style in the following situations:|
|[FontItalic](report-fontitalic-property-access.md)|You can use the  **FontItalic** property to specify whether text is italic in the following situations:|
|[FontName](report-fontname-property-access.md)|You can use the  **FontName** property to specify the font for text in the following situations:|
|[FontSize](report-fontsize-property-access.md)|You can use the  **FontSize** property to specify the point size for text in the following situations:|
|[FontUnderline](report-fontunderline-property-access.md)|You can use the  **FontUnderline** property to specify whether text is underlined in the following situations:|
|[ForeColor](report-forecolor-property-access.md)|You can use the  **ForeColor** property to specify the color for text in a control. Read/write **Long**.|
|[FormatCount](report-formatcount-property-access.md)|You can use the  **FormatCount** property to determine the number of times the **[OnFormat](section-onformat-property-access.md)** property has been evaluated for the current section on a report. Read/write **Integer**.|
|[GridX](report-gridx-property-access.md)|You can use the  **GridX** property (along with the **GridY** property) to specify the horizontal and vertical divisions of the alignment grid in report Design view. Read/write **Integer**.|
|[GridY](report-gridy-property-access.md)|You can use the  **GridY** property (along with the **GridX** property) to specify the horizontal and vertical divisions of the alignment grid in report Design view. Read/write **Integer**.|
|[GroupLevel](report-grouplevel-property-access.md)|You can use the  **GroupLevel** property in Visual Basic to refer to the group level you are grouping or sorting on in a report. Read-only **GroupLevel** object.|
|[GrpKeepTogether](report-grpkeeptogether-property-access.md)|You can use the  **GrpKeepTogether** property to specify whether groups in a multiple column report that have their **[KeepTogether](grouplevel-keeptogether-property-access.md)** property for a group set to Whole Group or With First Detail will be kept together by page or by column. Read/write **Byte**.|
|[HasData](report-hasdata-property-access.md)|You can use the  **HasData** property to determine if a report is bound to an empty recordset. Read/write **Long**.|
|[HasModule](report-hasmodule-property-access.md)|You can use the  **HasModule** property to specify or determine whether a form or report has a class module. Read/write **Boolean**.|
|[Height](report-height-property-access.md)|Gets or sets the height of the specified object in twips. Read/write  **Long**.|
|[HelpContextId](report-helpcontextid-property-access.md)|The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.|
|[HelpFile](report-helpfile-property-access.md)|The name of a help file associated with a report. Read/write  **String**.|
|[Hwnd](report-hwnd-property-access.md)|You can use the  **hWnd** property to determine the handle (a unique **Long Integer** value) assigned by Microsoft Windows to the current window. Read/write **Long**.|
|[InputParameters](report-inputparameters-property-access.md)|You can use the  **InputParameters** property to specify or determine the input parameters that are passed to a SQL statement in the **RecordSource** property of a form or report or a stored procedure when used as the record source within a Microsoft Access project (.adp). Read/write **String**.|
|[KeyPreview](report-keypreview-property-access.md)|You can use the  **KeyPreview** property to specify whether the report-level keyboard event procedures are invoked before a control's keyboard event procedures. Read/write **Boolean**.|
|[LayoutForPrint](report-layoutforprint-property-access.md)|You can use the  **LayoutForPrint** property to specify whether the report uses printer or screen fonts. Read/write **Boolean**.|
|[Left](report-left-property-access.md)|You can use the  **Left** property to specify an object's location on a form or report. Read/write **Long**.|
|[MenuBar](report-menubar-property-access.md)|Specifies a custom menu to display for a report. Read/write  **String**.|
|[MinMaxButtons](report-minmaxbuttons-property-access.md)|You can use the  **MinMaxButtons** property to specify whether the **Maximize** and **Minimize** buttons will be visible on a report. Read/write **Byte**.|
|[Modal](report-modal-property-access.md)|You can use the  **Modal** property to specify whether a report opens as a modal window. When a report opens as a modal window, you must close the window before you can move the focus to another object. Read/write **Boolean**.|
|[Module](report-module-property-access.md)|You can use the  **Module** property to specify a report module. Read-only **Module** object.|
|[MouseWheel](report-mousewheel-property-access.md)|Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **MouseWheel** event occurs. Read/write.|
|[Moveable](report-moveable-property-access.md)|Returns or sets a  **Boolean** indicating whether the specified report can be moved by the user; **True** if it can be moved. Read/write.|
|[MoveLayout](report-movelayout-property-access.md)|The  **MoveLayout** property specifies whether Microsoft Access should move to the next printing location on the page. Read/write **Boolean**.|
|[Name](report-name-property-access.md)|You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.|
|[NextRecord](report-nextrecord-property-access.md)|The  **NextRecord** property specifies whether a section should advance to the next record. Read/write **Boolean**.|
|[OnActivate](report-onactivate-property-access.md)|Sets or returns the value of the  **On Activate** box in the **Properties** window of a form or report. Read/write **String**.|
|[OnApplyFilter](report-onapplyfilter-property-access.md)|Sets or returns the value of the  **On Apply Filter** box in the **Properties** window of a report. Read/write **String**.|
|[OnClick](report-onclick-property-access.md)|Sets or returns the value of the  **On Click** box in the **Properties** window. Read/write **String**.|
|[OnClose](report-onclose-property-access.md)|Sets or returns the value of the  **On Close** box in the **Properties** window of a form or report. Read/write **String**.|
|[OnCurrent](report-oncurrent-property-access.md)|Sets or returns the value of the  **On Current** property on the Report. Read/write **String**.|
|[OnDblClick](report-ondblclick-property-access.md)|Sets or returns the value of the  **On Dbl Click** box in the **Properties** window. Read/write **String**.|
|[OnDeactivate](report-ondeactivate-property-access.md)|Sets or returns the value of the  **On Deactivate** box in the **Properties** window of a form or report. Read/write **String**.|
|[OnError](report-onerror-property-access.md)|Sets or returns the value of the  **OnError** box in the **Properties** window of a form or report. Read/write **String**.|
|[OnFilter](report-onfilter-property-access.md)|Sets or returns the value of the  **On Filter** box in the **Properties** window of a report. Read/write **String**.|
|[OnGotFocus](report-ongotfocus-property-access.md)|Sets or returns the value of the  **On Got Focus** box in the **Properties** window of the specified report. Read/write **String**.|
|[OnKeyDown](report-onkeydown-property-access.md)|Sets or returns the value of the  **On Key Down** box in the **Properties** window. Read/write **String**.|
|[OnKeyPress](report-onkeypress-property-access.md)|Sets or returns the value of the  **On Key Press** box in the **Properties** window. Read/write **String**.|
|[OnKeyUp](report-onkeyup-property-access.md)|Sets or returns the value of the  **On Key Up** box in the **Properties** window. Read/write **String**.|
|[OnLoad](report-onload-property-access.md)|Sets or returns the value of the  **On Load** box in the **Properties** window of a report. Read/write **String**.|
|[OnLostFocus](report-onlostfocus-property-access.md)|Sets or returns the value of the  **On Lost Focus** box in the **Properties** window of the specified report. Read/write **String**.|
|[OnMouseDown](report-onmousedown-property-access.md)|Sets or returns the value of the  **On Mouse Down** box in the **Properties** window. Read/write **String**.|
|[OnMouseMove](report-onmousemove-property-access.md)|Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.|
|[OnMouseUp](report-onmouseup-property-access.md)|Sets or returns the value of the  **On Mouse Up** box in the **Properties** window. Read/write **String**.|
|[OnNoData](report-onnodata-property-access.md)|Sets or returns the value of the  **On No Data** box in the **Properties** window of a report. Read/write **String**.|
|[OnOpen](report-onopen-property-access.md)|Sets or returns the value of the  **On Open** box in the **Properties** window of a form or report. Read/write **String**.|
|[OnPage](report-onpage-property-access.md)|Sets or returns the value of the  **On Page** box in the **Properties** window of a report. Read/write **String**.|
|[OnResize](report-onresize-property-access.md)|Sets or returns the value of the  **On Resize** box in the **Properties** window of a report. Read/write **String**.|
|[OnTimer](report-ontimer-property-access.md)|Sets or returns the value of the  **On Timer** box in the **Properties** window of a form. Read/write **String**.|
|[OnUnload](report-onunload-property-access.md)|Sets or returns the value of the  **On Unload** box in the **Properties** window of a form. Read/write **String**.|
|[OpenArgs](report-openargs-property-access.md)|Determines the string expression specified by the  _OpenArgs_ argument of the **OpenReport** method that opened a Report. Read/write **Variant**.|
|[OrderBy](report-orderby-property-access.md)|You can use the  **OrderBy** property to specify how you want to sort records in a report. Read/write **String**.|
|[OrderByOn](report-orderbyon-property-access.md)|You can use the  **OrderByOn** property to specify whether an object's **OrderBy** property setting is applied. Read/write **Boolean**.|
|[OrderByOnLoad](report-orderbyonload-property-access.md)|Gets or sets whether the sorting specified by the  **[OrderBy](report-orderby-property-access.md)** property is applied when the report is loaded. Read/write **Boolean**.|
|[Orientation](report-orientation-property-access.md)|You can use the  **Orientation** property to specify or determine the view orientation. Read/write **Byte**.|
|[Page](report-page-property-access.md)|The  **Page** property specifies the current page number when a report is being printed. Read/write **Long**.|
|[PageFooter](report-pagefooter-property-access.md)|You can use the  **PageFooter** property to specify whether a report's page footer is printed on the same page as a report footer. Read/write **Byte**.|
|[PageHeader](report-pageheader-property-access.md)|You can use the  **PageHeader** property to specify whether a report's page header is printed on the same page as a report header. Read/write **Byte**.|
|[Pages](report-pages-property-access.md)|You can use the  **Pages** property to return information needed to print page numbers in a report. Read/write **Integer**.|
|[Painting](report-painting-property-access.md)|You can use the  **Painting** property to specify whether a report is repainted. Read/write **Boolean**.|
|[PaintPalette](report-paintpalette-property-access.md)|You can use the  **PaintPalette** property to specify a palette to be used by a report. Read/write **Variant**.|
|[PaletteSource](report-palettesource-property-access.md)|You can use the  **PaletteSource** property to specify the palette for a report. Read/write **String**.|
|[Parent](report-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Picture](report-picture-property-access.md)|You can use the  **Picture** property to specify a bitmap or other type of graphic to be used as a background picture on a report. Read/write **String**.|
|[PictureAlignment](report-picturealignment-property-access.md)|You can use the  **PictureAlignment** property to specify where a background picture will appear in an image control or on a form or report. Read/write **Byte**.Read/write.|
|[PictureData](report-picturedata-property-access.md)|You can use the  **PictureData** property to copy the picture to another object that supports the **Picture** property. Read/write **Variant**.|
|[PicturePages](report-picturepages-property-access.md)|You can use the  **PicturePages** property to specify on which page or pages of a report a picture will be displayed. Read/write **Byte**.|
|[PicturePalette](report-picturepalette-property-access.md)|You can use the  **PicturePalette** property to specify a palette to be used by a report. Read/write **Variant**.|
|[PictureSizeMode](report-picturesizemode-property-access.md)|You can use the  **PictureSizeMode** property to specify how a picture for a form or report is sized. Read/write **Byte**.|
|[PictureTiling](report-picturetiling-property-access.md)|You can use the  **PictureTiling** property to specify whether a background picture is tiled across the entire image control, Form window, form, or page of a report. Read/write **Boolean**.|
|[PictureType](report-picturetype-property-access.md)|You can use the  **PictureType** property to specify whether Microsoft Access stores an object's picture as a linked or an embedded object. Read/write **Byte**.|
|[PopUp](report-popup-property-access.md)|Specifies whether a report opens as a pop-up window. Read/write  **Boolean**.|
|[PrintCount](report-printcount-property-access.md)|You can use the  **PrintCount** property to identify the number of times the **OnPrint** property has been evaluated for the current section of a report. Read/write **Integer**.|
|[Printer](report-printer-property-access.md)|Returns or sets a  **[Printer](printer-object-access.md)** object representing the default printer on the current system. Read/write.|
|[PrintSection](report-printsection-property-access.md)|The  **PrintSection** property specifies whether a section should be printed. Read/write **Boolean**.|
|[Properties](report-properties-property-access.md)|Returns a reference to a control's **[Properties](properties-object-access.md)** collection object. Read-only.|
|[PrtDevMode](report-prtdevmode-property-access.md)|You can use the  **PrtDevMode** property to set or return printing device mode information specified for a form or report in the **Print** dialog box. Read/write **Variant**.|
|[PrtDevNames](report-prtdevnames-property-access.md)|You can use the  **PrtDevNames** property to set or return information about the printer selected in the **Print** dialog box for a form or report. Read/write **Variant**.|
|[PrtMip](report-prtmip-property-access.md)|You can use the  **PrtMip** property in Visual Basic to set or return the device mode information specified for a form or report in the **Print** dialog box.|
|[RecordLocks](report-recordlocks-property-access.md)|You can use the  **RecordLocks** property to determine how records are locked and what happens when two users try to edit the same record at the same time. Read/write.|
|[Recordset](report-recordset-property-access.md)|Returns or sets the ADO  **Recordset** or DAO **[Recordset](recordset-object-dao.md)** object representing the record source for the specified object. Read/write **Object**.|
|[RecordSource](report-recordsource-property-access.md)|You can use the  **RecordSource** property to specify the source of the data for a report. Read/write **String**.|
|[RecordSourceQualifier](report-recordsourcequalifier-property-access.md)|Returns or sets a  **String** indicating the SQL Server owner name of the record source for the specified report. Read/write.|
|[Report](report-report-property-access.md)|You can use the  **Report** property to refer to a report or to refer to the report associated with a subreport control. Read-only **Report**.|
|[RibbonName](report-ribbonname-property-access.md)|Gets or set the name of the customized ribbon to be displayed when the specified report is loaded. Read/write  **String**.|
|[ScaleHeight](report-scaleheight-property-access.md)|You can use the  **ScaleHeight** property to specify the number of units for the vertical measurement of the page when the **[Circle](report-circle-method-access.md)**, **[Line](report-line-method-access.md)**, **[Pset](report-pset-method-access.md)**, or **[Print](report-print-method-access.md)** method is used while a report is printed or previewed, or its output is saved to a file. Read/write **Single**.|
|[ScaleLeft](report-scaleleft-property-access.md)|You can use the  **ScaleLeft** property to specify the units for the horizontal coordinates that describe the location of the left edge of a page when the **[Circle](report-circle-method-access.md)**, **[Line](report-line-method-access.md)**, **[Pset](report-pset-method-access.md)**, or **[Print](report-print-method-access.md)** method is used while a report is previewed, printed, or its output is saved to a file. Read / write **Single**.|
|[ScaleMode](report-scalemode-property-access.md)|You can use the  **ScaleMode** property in Visual Basic to specify the unit of measurement for coordinates on a page when the **[Circle](report-circle-method-access.md)**, **[Line](report-line-method-access.md)**, **[Pset](report-pset-method-access.md)**, or **[Print](report-print-method-access.md)** method is used while a report is previewed or printed, or its output is saved to a file. Read/write **Integer**.|
|[ScaleTop](report-scaletop-property-access.md)|You can use the  **ScaleTop** property to specify the units for the vertical coordinates that describe the location of the top edge of a page when the **[Circle](report-circle-method-access.md)**, **[Line](report-line-method-access.md)**, **[Pset](report-pset-method-access.md)**, or **[Print](report-print-method-access.md)** method is used while a report is previewed, printed, or its output is saved to a file. Read / write **Single**.|
|[ScaleWidth](report-scalewidth-property-access.md)|You can use the  **ScaleWidth** property to specify the number of units for the horizontal measurement of the page when the **[Circle](report-circle-method-access.md)**, **[Line](report-line-method-access.md)**, **[Pset](report-pset-method-access.md)**, or **[Print](report-print-method-access.md)** method is used while a report is printed or previewed, or its output is saved to a file. Read/write **Single**.|
|[ScrollBars](report-scrollbars-property-access.md)|Gets or sets whether scroll bars appear on a report. Read/write  **Byte**.|
|[Section](report-section-property-access.md)|You can use the  **Section** property to identify a section of a report and provide access to the properties of that section. Read-only **Section** object.|
|[ServerFilter](report-serverfilter-property-access.md)|You can use the  **ServerFilter** property to specify a subset of records to be displayed when a server filter is applied to a report within a Microsoft Access project (.adp) or database. Read/write **String**.|
|[Shape](report-shape-property-access.md)|Returns a  **String** representing the shape command corresponding to the sorting and grouping of the specified report. Read-only.|
|[ShortcutMenuBar](report-shortcutmenubar-property-access.md)|You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.|
|[ShowPageMargins](report-showpagemargins-property-access.md)|Gets or sets whether page margins are displayed when the specified report is in Layout view. Read/write  **Boolean**.|
|[Tag](report-tag-property-access.md)|Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write  **String**.|
|[TimerInterval](report-timerinterval-property-access.md)|You can use the  **TimerInterval** property to specify the interval, in milliseconds, between **[Timer](report-timer-event-access.md)** events on a report. Read/write **Long**.|
|[Toolbar](report-toolbar-property-access.md)|Specifies a custom toolbar to display for a report. Read/write  **String**.|
|[Top](report-top-property-access.md)|You can use the  **Top** property to specify an object's location on a form or report. Read/write **Long**. .|
|[UseDefaultPrinter](report-usedefaultprinter-property-access.md)|Returns or sets a  **Boolean** indicating whether the specified report uses the default printer for the system; **True** if the form or report uses the default printer. Read/write.|
|[Visible](report-visible-property-access.md)|Returns or sets whether the object is visible. Read/write  **Boolean**.|
|[Width](report-width-property-access.md)|Gets or sets the width of the specified object in twips. Read/write  **Integer**.|
|[WindowHeight](report-windowheight-property-access.md)|Returns the height of a report in twips. Read-only  **Integer**.|
|[WindowLeft](report-windowleft-property-access.md)|Returns an  **Integer** indicating the screen position in twips of the left edge of a report relative to the left edge of the Microsoft Access window. Read-only.|
|[WindowTop](report-windowtop-property-access.md)|Returns an  **Integer** indicating the screen position in twips of the top edge of a report relative to the top of the Microsoft Access window. Read-only.|
|[WindowWidth](report-windowwidth-property-access.md)|Returns the width of a report in twips. Read-only  **Integer**.|

