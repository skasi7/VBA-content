---
title: Window Members (Excel)
ms.prod: EXCEL
ms.assetid: f11db427-24a4-041c-2fd5-03ce73ae6c16
---


# Window Members (Excel)
Represents a window.

Represents a window.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](window-activate-method-excel.md)|Brings the window to the front of the z-order. |
|[ActivateNext](window-activatenext-method-excel.md)|Activates the specified window and then sends it to the back of the window z-order.|
|[ActivatePrevious](window-activateprevious-method-excel.md)|Activates the specified window and then activates the window at the back of the window z-order.|
|[Close](window-close-method-excel.md)|Closes the object.|
|[LargeScroll](window-largescroll-method-excel.md)|Scrolls the contents of the window by pages.|
|[NewWindow](window-newwindow-method-excel.md)|Creates a new window or a copy of the specified window.|
|[PointsToScreenPixelsX](window-pointstoscreenpixelsx-method-excel.md)|Converts a horizontal measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a  **Long** value.|
|[PointsToScreenPixelsY](window-pointstoscreenpixelsy-method-excel.md)|Converts a vertical measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a  **Long** value.|
|[PrintOut](window-printout-method-excel.md)|Prints the object.|
|[PrintPreview](window-printpreview-method-excel.md)|Shows a preview of the object as it would look when printed.|
|[RangeFromPoint](window-rangefrompoint-method-excel.md)|Returns the  **[Shape](shape-object-excel.md)** or **[Range](range-object-excel.md)** object that is positioned at the specified pair of screen coordinates. If there isn't a shape located at the specified coordinates, this method returns **Nothing** .|
|[ScrollIntoView](window-scrollintoview-method-excel.md)|Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the document window or pane (depending on the value of the  _Start_ argument).|
|[ScrollWorkbookTabs](window-scrollworkbooktabs-method-excel.md)|Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.|
|[SmallScroll](window-smallscroll-method-excel.md)|Scrolls the contents of the window by rows or columns.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActiveCell](window-activecell-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.|
|[ActiveChart](window-activechart-property-excel.md)|Returns a  **[Chart](chart-object-excel.md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing** .|
|[ActivePane](window-activepane-property-excel.md)|Returns a  **[Pane](pane-object-excel.md)** object that represents the active pane in the window. Read-only.|
|[ActiveSheet](window-activesheet-property-excel.md)|Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns  **Nothing** if no sheet is active.|
|[ActiveSheetView](window-activesheetview-property-excel.md)| Returns an object that represents the view of the active sheet in the specified window. Read-only.|
|[Application](window-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoFilterDateGrouping](window-autofilterdategrouping-property-excel.md)| **True** if the auto filter for date grouping is currently displayed in the specified window. Read/write **Boolean** .|
|[Caption](window-caption-property-excel.md)|Returns or sets a  **Variant** value that represents the name that appears in the title bar of the document window.|
|[Creator](window-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DisplayFormulas](window-displayformulas-property-excel.md)| **True** if the window is displaying formulas; **False** if the window is displaying values. Read/write **Boolean** .|
|[DisplayGridlines](window-displaygridlines-property-excel.md)| **True** if gridlines are displayed. Read/write **Boolean** .|
|[DisplayHeadings](window-displayheadings-property-excel.md)| **True** if both row and column headings are displayed; **False** if no headings are displayed. Read/write **Boolean** .|
|[DisplayHorizontalScrollBar](window-displayhorizontalscrollbar-property-excel.md)| **True** if the horizontal scroll bar is displayed. Read/write **Boolean** .|
|[DisplayOutline](window-displayoutline-property-excel.md)| **True** if outline symbols are displayed. Read/write **Boolean** .|
|[DisplayRightToLeft](window-displayrighttoleft-property-excel.md)| **True** if the specified window is displayed from right to left instead of from left to right. **False** if the object is displayed from left to right. Read-only **Boolean** .|
|[DisplayRuler](window-displayruler-property-excel.md)| **True** if a ruler is displayed for the specified window. Read/write **Boolean** .|
|[DisplayVerticalScrollBar](window-displayverticalscrollbar-property-excel.md)| **True** if the vertical scroll bar is displayed. Read/write **Boolean** .|
|[DisplayWhitespace](window-displaywhitespace-property-excel.md)| **True** if whitespace is displayed. Read/write **Boolean** .|
|[DisplayWorkbookTabs](window-displayworkbooktabs-property-excel.md)| **True** if the workbook tabs are displayed. Read/write **Boolean** .|
|[DisplayZeros](window-displayzeros-property-excel.md)| **True** if zero values are displayed. Read/write **Boolean** .|
|[EnableResize](window-enableresize-property-excel.md)| **True** if the window can be resized. Read/write **Boolean** .|
|[FreezePanes](window-freezepanes-property-excel.md)| **True** if split panes are frozen. Read/write **Boolean** .|
|[GridlineColor](window-gridlinecolor-property-excel.md)|Returns or sets the gridline color as an RGB value. Read/write  **Long** .|
|[GridlineColorIndex](window-gridlinecolorindex-property-excel.md)|Returns or sets the gridline color as an index into the current color palette or as the following  **[XlColorIndex](xlcolorindex-enumeration-excel.md)** constant.|
|[Height](window-height-property-excel.md)|Returns or sets a  **Double** value that represents tThe height, in points, of the window.|
|[Hwnd](window-hwnd-property-excel.md)||
|[Index](window-index-property-excel.md)|Returns a  **Long** value that represents the index number of the object within the collection of similar objects.|
|[Left](window-left-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the client area to the left edge of the window.|
|[OnWindow](window-onwindow-property-excel.md)|Returns or sets the name of the procedure that's run whenever you activate a window. Read/write  **String** .|
|[Panes](window-panes-property-excel.md)|Returns a  **[Panes](panes-object-excel.md)** collection that represents all the panes in the specified window. Read-only.|
|[Parent](window-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[RangeSelection](window-rangeselection-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet. Read-only.|
|[ScrollColumn](window-scrollcolumn-property-excel.md)|Returns or sets the number of the leftmost column in the pane or window. Read/write  **Long** .|
|[ScrollRow](window-scrollrow-property-excel.md)|Returns or sets the number of the row that appears at the top of the pane or window. Read/write  **Long** .|
|[SelectedSheets](window-selectedsheets-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the selected sheets in the specified window. Read-only.|
|[Selection](window-selection-property-excel.md)|Returns the specified window, for a  **[Windows](windows-object-excel.md)** object.|
|[SheetViews](window-sheetviews-property-excel.md)|Returns the  **[SheetViews](sheetviews-object-excel.md)** object for the specified window. Read-only.|
|[Split](window-split-property-excel.md)| **True** if the window is split. Read/write **Boolean** .|
|[SplitColumn](window-splitcolumn-property-excel.md)|Returns or sets the column number where the window is split into panes (the number of columns to the left of the split line). Read/write  **Long** .|
|[SplitHorizontal](window-splithorizontal-property-excel.md)|Returns or sets the location of the horizontal window split, in points. Read/write  **Double** .|
|[SplitRow](window-splitrow-property-excel.md)|Returns or sets the row number where the window is split into panes (the number of rows above the split). Read/write  **Long** .|
|[SplitVertical](window-splitvertical-property-excel.md)|Returns or sets the location of the vertical window split, in points. Read/write  **Double** .|
|[TabRatio](window-tabratio-property-excel.md)|Returns or sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar (as a number between 0 (zero) and 1; the default value is 0.6). Read/write  **Double** .|
|[Top](window-top-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).|
|[Type](window-type-property-excel.md)|Returns or sets a  **[XlWindowType](xlwindowtype-enumeration-excel.md)** value that represents the window type.|
|[UsableHeight](window-usableheight-property-excel.md)|Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only  **Double** .|
|[UsableWidth](window-usablewidth-property-excel.md)|Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only  **Double** .|
|[View](window-view-property-excel.md)|Returns or sets the view showing in the window. Read/write  **[XlWindowView](xlwindowview-enumeration-excel.md)** .|
|[Visible](window-visible-property-excel.md)|Returns or sets a  **Boolean** value that determines whether the object is visible. Read/write.|
|[VisibleRange](window-visiblerange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. Read-only.|
|[Width](window-width-property-excel.md)|Returns or sets a  **Double** value that represents the width, in points, of the window.|
|[WindowNumber](window-windownumber-property-excel.md)|Returns the window number. For example, a window named "Book1.xls:2" has 2 as its window number. Most windows have the window number 1. Read-only  **Long** .|
|[WindowState](window-windowstate-property-excel.md)|Returns or sets the state of the window. Read/write  **[XlWindowState](xlwindowstate-enumeration-excel.md)** .|
|[Zoom](window-zoom-property-excel.md)|Returns or sets a  **Variant** value that represents the display size of the window, as a percentage (100 equals normal size, 200 equals double size, and so on).|

