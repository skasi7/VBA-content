---
title: Application Methods (Excel)
ms.prod: EXCEL
ms.assetid: b82ec21c-30e6-4d37-8cf8-a69b27ab9771
---


# Application Methods (Excel)

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ActivateMicrosoftApp](application-activatemicrosoftapp-method-excel.md)|Activates a Microsoft application. If the application is already running, this method activates the running application. If the application isn't running, this method starts a new instance of the application.|
|[AddCustomList](application-addcustomlist-method-excel.md)|Adds a custom list for custom autofill and/or custom sort.|
|[Calculate](application-calculate-method-excel.md)|Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.|
|[CalculateFull](application-calculatefull-method-excel.md)|Forces a full calculation of the data in all open workbooks.|
|[CalculateFullRebuild](application-calculatefullrebuild-method-excel.md)|For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.|
|[CalculateUntilAsyncQueriesDone](application-calculateuntilasyncqueriesdone-method-excel.md)|Runs all pending queries to OLEDB and OLAP data sources.|
|[CentimetersToPoints](application-centimeterstopoints-method-excel.md)|Converts a measurement from centimeters to points (one point equals 0.035 centimeters).|
|[CheckAbort](application-checkabort-method-excel.md)|Stops recalculation in a Microsoft Excel application.|
|[CheckSpelling](application-checkspelling-method-excel.md)|Checks the spelling of a single word.|
|[ConvertFormula](application-convertformula-method-excel.md)|Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both.  **Variant** .|
|[DDEExecute](application-ddeexecute-method-excel.md)|Runs a command or performs some other action or actions in another application by way of the specified DDE channel.|
|[DDEInitiate](application-ddeinitiate-method-excel.md)|Opens a DDE channel to an application.|
|[DDEPoke](application-ddepoke-method-excel.md)|Sends data to an application.|
|[DDERequest](application-dderequest-method-excel.md)|Requests information from the specified application. This method always returns an array.|
|[DDETerminate](application-ddeterminate-method-excel.md)|Closes a channel to another application.|
|[DeleteCustomList](application-deletecustomlist-method-excel.md)|Deletes a custom list.|
|[DisplayXMLSourcePane](application-displayxmlsourcepane-method-excel.md)|Opens the  **XML Source** task pane and displays the XML map specified by the _XmlMap_ argument.|
|[DoubleClick](application-doubleclick-method-excel.md)|Equivalent to double-clicking the active cell.|
|[Evaluate](application-evaluate-method-excel.md)|Converts a Microsoft Excel name to an object or a value.|
|[ExecuteExcel4Macro](application-executeexcel4macro-method-excel.md)|Runs a Microsoft Excel 4.0 macro function and then returns the result of the function. The return type depends on the function.|
|[FindFile](application-findfile-method-excel.md)|Displays the  **Open** dialog box.|
|[GetCustomListContents](application-getcustomlistcontents-method-excel.md)|Returns a custom list (an array of strings).|
|[GetCustomListNum](application-getcustomlistnum-method-excel.md)|Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.|
|[GetOpenFilename](application-getopenfilename-method-excel.md)|Displays the standard  **Open** dialog box and gets a file name from the user without actually opening any files.|
|[GetPhonetic](application-getphonetic-method-excel.md)|Returns the Japanese phonetic text of the specified text string. This method is available to you only if you have selected or installed Japanese language support for Microsoft Office.|
|[GetSaveAsFilename](application-getsaveasfilename-method-excel.md)|Displays the standard  **Save As** dialog box and gets a file name from the user without actually saving any files.|
|[Goto](application-goto-method-excel.md)|Selects any range or Visual Basic procedure in any workbook, and activates that workbook if it's not already active.|
|[Help](application-help-method-excel.md)|Displays a Help topic.|
|[InchesToPoints](application-inchestopoints-method-excel.md)|Converts a measurement from inches to points.|
|[InputBox](application-inputbox-method-excel.md)|Displays a dialog box for user input. Returns the information entered in the dialog box.|
|[Intersect](application-intersect-method-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the rectangular intersection of two or more ranges.|
|[MacroOptions](application-macrooptions-method-excel.md)|Corresponds to options in the  **Macro Options** dialog box. You can also use this method to display a user defined function (UDF) in a built-in or new category within the **Insert Function** dialog box.|
|[MailLogoff](application-maillogoff-method-excel.md)|Closes a MAPI mail session established by Microsoft Excel.|
|[MailLogon](application-maillogon-method-excel.md)|Logs in to MAPI Mail or Microsoft Exchange and establishes a mail session. If Microsoft Mail isn't already running, you must use this method to establish a mail session before mail or document routing functions can be used.|
|[NextLetter](application-nextletter-method-excel.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
|[OnKey](application-onkey-method-excel.md)|Runs a specified procedure when a particular key or key combination is pressed.|
|[OnRepeat](application-onrepeat-method-excel.md)|Sets the  **Repeat** item and the name of the procedure that will run if you choose the **Repeat** command after running the procedure that sets this property.|
|[OnTime](application-ontime-method-excel.md)|Schedules a procedure to be run at a specified time in the future (either at a specific time of day or after a specific amount of time has passed).|
|[OnUndo](application-onundo-method-excel.md)|Sets the text of the  **Undo** command and the name of the procedure that's run if you choose the **Undo** command after running the procedure that sets this property.|
|[Quit](application-quit-method-excel.md)|Quits Microsoft Excel.|
|[RecordMacro](application-recordmacro-method-excel.md)|Records code if the macro recorder is on.|
|[RegisterXLL](application-registerxll-method-excel.md)|Loads an XLL code resource and automatically registers the functions and commands contained in the resource.|
|[Repeat](application-repeat-method-excel.md)|Repeats the last user-interface action.|
|[Run](application-run-method-excel.md)|Runs a macro or calls a function. This can be used to run a macro written in Visual Basic or the Microsoft Excel macro language, or to run a function in a DLL or XLL.|
|[SendKeys](application-sendkeys-method-excel.md)|Sends keystrokes to the active application.|
|[SharePointVersion](application-sharepointversion-method-excel.md)|Returns the version number of SharePoint Foundation instances running at site for the specified URL.|
|[Undo](application-undo-method-excel.md)|Cancels the last user-interface action.|
|[Union](application-union-method-excel.md)|Returns the union of two or more ranges.|
|[Volatile](application-volatile-method-excel.md)|Marks a user-defined function as volatile. A volatile function must be recalculated whenever calculation occurs in any cells on the worksheet. A nonvolatile function is recalculated only when the input variables change. This method has no effect if it's not inside a user-defined function used to calculate a worksheet cell.|
|[Wait](application-wait-method-excel.md)|Pauses a running macro until a specified time. Returns  **True** if the specified time has arrived.|

