---
title: Application Events (Excel)
ms.prod: EXCEL
ms.assetid: 726704da-6d45-4728-bd47-4ec9910626b2
---


# Application Events (Excel)
This object has the following events:

## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterCalculate](application-aftercalculate-event-excel.md)|The  **AfterCalculate** event occurs when all pending refresh activity (both synchronous and asynchronous) and all of the resultant calculation activities have been completed.|
|[NewWorkbook](application-newworkbook-event-excel.md)|Occurs when a new workbook is created.|
|[ProtectedViewWindowActivate](application-protectedviewwindowactivate-event-excel.md)|Occurs when a  **Protected View** window is activated.|
|[ProtectedViewWindowBeforeClose](application-protectedviewwindowbeforeclose-event-excel.md)|Occurs immediately before a  **Protected View** window or a workbook in a **Protected View** window closes.|
|[ProtectedViewWindowBeforeEdit](application-protectedviewwindowbeforeedit-event-excel.md)|Occurs immediately before editing is enabled on the workbook in the specified  **Protected View** window.|
|[ProtectedViewWindowDeactivate](application-protectedviewwindowdeactivate-event-excel.md)|Occurs when a  **Protected View** window is deactivated.|
|[ProtectedViewWindowOpen](application-protectedviewwindowopen-event-excel.md)|Occurs when a workbook is opened in a  **Protected View** window.|
|[ProtectedViewWindowResize](application-protectedviewwindowresize-event-excel.md)|Occurs when any  **Protected View** window is resized.|
|[SheetActivate](application-sheetactivate-event-excel.md)|Occurs when any sheet is activated.|
|[SheetBeforeDelete](application-sheetbeforedelete-event-excel.md)||
|[SheetBeforeDoubleClick](application-sheetbeforedoubleclick-event-excel.md)|Occurs when any worksheet is double-clicked, before the default double-click action.|
|[SheetBeforeRightClick](application-sheetbeforerightclick-event-excel.md)|Occurs when any worksheet is right-clicked, before the default right-click action.|
|[SheetCalculate](application-sheetcalculate-event-excel.md)|Occurs after any worksheet is recalculated or after any changed data is plotted on a chart.|
|[SheetChange](application-sheetchange-event-excel.md)|Occurs when cells in any worksheet are changed by the user or by an external link.|
|[SheetDeactivate](application-sheetdeactivate-event-excel.md)|Occurs when any sheet is deactivated.|
|[SheetFollowHyperlink](application-sheetfollowhyperlink-event-excel.md)|Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the  **[FollowHyperlink](worksheet-followhyperlink-event-excel.md)** event.|
|[SheetLensGalleryRenderComplete](application-sheetlensgalleryrendercomplete-event-excel.md)|Occurs after a callout gallery's icons (dynamic &; static) have finished rendering.|
|[SheetPivotTableAfterValueChange](application-sheetpivottableaftervaluechange-event-excel.md)|Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).|
|[SheetPivotTableBeforeAllocateChanges](application-sheetpivottablebeforeallocatechanges-event-excel.md)|Occurs before changes are applied to a PivotTable.|
|[SheetPivotTableBeforeCommitChanges](application-sheetpivottablebeforecommitchanges-event-excel.md)|Occurs before changes are committed against the OLAP data source for a PivotTable.|
|[SheetPivotTableBeforeDiscardChanges](application-sheetpivottablebeforediscardchanges-event-excel.md)|Occurs before changes to a PivotTable are discarded.|
|[SheetPivotTableUpdate](application-sheetpivottableupdate-event-excel.md)|Occurs after the sheet of the PivotTable report has been updated.|
|[SheetSelectionChange](application-sheetselectionchange-event-excel.md)|Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).|
|[SheetTableUpdate](application-sheettableupdate-event-excel.md)|Occurs when a table on a worksheet is updated.|
|[WindowActivate](application-windowactivate-event-excel.md)|Occurs when any workbook window is activated.|
|[WindowDeactivate](application-windowdeactivate-event-excel.md)|Occurs when any workbook window is deactivated.|
|[WindowResize](application-windowresize-event-excel.md)|Occurs when any workbook window is resized.|
|[WorkbookActivate](application-workbookactivate-event-excel.md)|Occurs when any workbook is activated.|
|[WorkbookAddinInstall](application-workbookaddininstall-event-excel.md)|Occurs when a workbook is installed as an add-in.|
|[WorkbookAddinUninstall](application-workbookaddinuninstall-event-excel.md)|Occurs when any add-in workbook is uninstalled.|
|[WorkbookAfterSave](application-workbookaftersave-event-excel.md)|Occurs after the workbook is saved.|
|[WorkbookAfterXmlExport](application-workbookafterxmlexport-event-excel.md)|Occurs after Microsoft Excel saves or exports XML data from the specified workbook.|
|[WorkbookAfterXmlImport](application-workbookafterxmlimport-event-excel.md)|Occurs after an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.|
|[WorkbookBeforeClose](application-workbookbeforeclose-event-excel.md)|Occurs immediately before any open workbook closes.|
|[WorkbookBeforePrint](application-workbookbeforeprint-event-excel.md)|Occurs before any open workbook is printed.|
|[WorkbookBeforeSave](application-workbookbeforesave-event-excel.md)|Occurs before any open workbook is saved.|
|[WorkbookBeforeXmlExport](application-workbookbeforexmlexport-event-excel.md)|Occurs before Microsoft Excel saves or exports XML data from the specified workbook.|
|[WorkbookBeforeXmlImport](application-workbookbeforexmlimport-event-excel.md)|Occurs before an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.|
|[WorkbookDeactivate](application-workbookdeactivate-event-excel.md)|Occurs when any open workbook is deactivated.|
|[WorkbookModelChange](application-workbookmodelchange-event-excel.md)|Occurs when the data model is updated.|
|[WorkbookNewChart](application-workbooknewchart-event-excel.md)|Occurs when a new chart is created in any open workbook.|
|[WorkbookNewSheet](application-workbooknewsheet-event-excel.md)|Occurs when a new sheet is created in any open workbook.|
|[WorkbookOpen](application-workbookopen-event-excel.md)|Occurs when a workbook is opened.|
|[WorkbookPivotTableCloseConnection](application-workbookpivottablecloseconnection-event-excel.md)|Occurs after a PivotTable report connection has been closed.|
|[WorkbookPivotTableOpenConnection](application-workbookpivottableopenconnection-event-excel.md)|Occurs after a PivotTable report connection has been opened.|
|[WorkbookRowsetComplete](application-workbookrowsetcomplete-event-excel.md)|The  **WorkbookRowsetComplete** event occurs when the user either drills through the recordset or invokes the rowset action on an OLAP PivotTable.|
|[WorkbookSync](application-workbooksync-event-excel.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|

