---
title: DoCmd Object (Access)
keywords: vbaac10.chm4241
f1_keywords:
- vbaac10.chm4241
ms.prod: ACCESS
api_name:
- Access.DoCmd
ms.assetid: 3ce44cca-9979-0a1e-9787-079a52ce528f
---


# DoCmd Object (Access)

You can use the methods of the  **DoCmd** object to run Microsoft Office Access actions from Visual Basic. An action performs tasks such as closing windows, opening forms, and setting the value of controls.


## Remarks

For example, you can use the  **OpenForm** method of the **DoCmd** object to open a form, or use the **Hourglass** method to change the mouse pointer to an hourglass icon.

Most of the methods of the  **DoCmd** object have arguments — some are required, while others are optional. If you omit optional arguments, the arguments assume the default values for the particular method. For example, the **OpenForm** method uses seven arguments, but only the first argument, _FormName_, is required. The following example shows how you can open the Employees form in the current database. Only employees with the title Sales Representative are included.




```
DoCmd.OpenForm "Employees", , ,"[Title] = 'Sales Representative'"
```

The  **DoCmd** object doesn't support methods corresponding to the following actions:
    
- MsgBox. Use the  **MsgBox** function.
    
- RunApp. Use the  **Shell** function to run another application.
    
- RunCode. Run the function directly in Visual Basic.
    
- SendKeys. Use the  **SendKeys** statement.
    
- SetValue. Set the value directly in Visual Basic.
    
- StopAllMacros.
    
- StopMacro.
    

## Example

The following example opens a form in Form view and moves to a new record.


```
Sub ShowNewRecord() 
 DoCmd.OpenForm "Employees", acNormal 
 DoCmd.GoToRecord , , acNewRec 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddMenu](http://msdn.microsoft.com/library/docmd-addmenu-method-access%28Office.15%29.aspx)|
|[ApplyFilter](http://msdn.microsoft.com/library/docmd-applyfilter-method-access%28Office.15%29.aspx)|
|[Beep](http://msdn.microsoft.com/library/docmd-beep-method-access%28Office.15%29.aspx)|
|[BrowseTo](http://msdn.microsoft.com/library/docmd-browseto-method-access%28Office.15%29.aspx)|
|[CancelEvent](http://msdn.microsoft.com/library/docmd-cancelevent-method-access%28Office.15%29.aspx)|
|[ClearMacroError](http://msdn.microsoft.com/library/docmd-clearmacroerror-method-access%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/docmd-close-method-access%28Office.15%29.aspx)|
|[CloseDatabase](http://msdn.microsoft.com/library/docmd-closedatabase-method-access%28Office.15%29.aspx)|
|[CopyDatabaseFile](http://msdn.microsoft.com/library/docmd-copydatabasefile-method-access%28Office.15%29.aspx)|
|[CopyObject](http://msdn.microsoft.com/library/docmd-copyobject-method-access%28Office.15%29.aspx)|
|[DeleteObject](http://msdn.microsoft.com/library/docmd-deleteobject-method-access%28Office.15%29.aspx)|
|[DoMenuItem](http://msdn.microsoft.com/library/docmd-domenuitem-method-access%28Office.15%29.aspx)|
|[Echo](http://msdn.microsoft.com/library/docmd-echo-method-access%28Office.15%29.aspx)|
|[FindNext](http://msdn.microsoft.com/library/docmd-findnext-method-access%28Office.15%29.aspx)|
|[FindRecord](http://msdn.microsoft.com/library/docmd-findrecord-method-access%28Office.15%29.aspx)|
|[GoToControl](http://msdn.microsoft.com/library/docmd-gotocontrol-method-access%28Office.15%29.aspx)|
|[GoToPage](http://msdn.microsoft.com/library/docmd-gotopage-method-access%28Office.15%29.aspx)|
|[GoToRecord](http://msdn.microsoft.com/library/docmd-gotorecord-method-access%28Office.15%29.aspx)|
|[Hourglass](http://msdn.microsoft.com/library/docmd-hourglass-method-access%28Office.15%29.aspx)|
|[LockNavigationPane](http://msdn.microsoft.com/library/docmd-locknavigationpane-method-access%28Office.15%29.aspx)|
|[Maximize](http://msdn.microsoft.com/library/docmd-maximize-method-access%28Office.15%29.aspx)|
|[Minimize](http://msdn.microsoft.com/library/docmd-minimize-method-access%28Office.15%29.aspx)|
|[MoveSize](http://msdn.microsoft.com/library/docmd-movesize-method-access%28Office.15%29.aspx)|
|[NavigateTo](http://msdn.microsoft.com/library/docmd-navigateto-method-access%28Office.15%29.aspx)|
|[OpenDataAccessPage](http://msdn.microsoft.com/library/docmd-opendataaccesspage-method-access%28Office.15%29.aspx)|
|[OpenDiagram](http://msdn.microsoft.com/library/docmd-opendiagram-method-access%28Office.15%29.aspx)|
|[OpenForm](http://msdn.microsoft.com/library/docmd-openform-method-access%28Office.15%29.aspx)|
|[OpenFunction](http://msdn.microsoft.com/library/docmd-openfunction-method-access%28Office.15%29.aspx)|
|[OpenModule](http://msdn.microsoft.com/library/docmd-openmodule-method-access%28Office.15%29.aspx)|
|[OpenQuery](http://msdn.microsoft.com/library/docmd-openquery-method-access%28Office.15%29.aspx)|
|[OpenReport](http://msdn.microsoft.com/library/docmd-openreport-method-access%28Office.15%29.aspx)|
|[OpenStoredProcedure](http://msdn.microsoft.com/library/docmd-openstoredprocedure-method-access%28Office.15%29.aspx)|
|[OpenTable](http://msdn.microsoft.com/library/docmd-opentable-method-access%28Office.15%29.aspx)|
|[OpenView](http://msdn.microsoft.com/library/docmd-openview-method-access%28Office.15%29.aspx)|
|[OutputTo](http://msdn.microsoft.com/library/docmd-outputto-method-access%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/docmd-printout-method-access%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/docmd-quit-method-access%28Office.15%29.aspx)|
|[RefreshRecord](http://msdn.microsoft.com/library/docmd-refreshrecord-method-access%28Office.15%29.aspx)|
|[Rename](http://msdn.microsoft.com/library/docmd-rename-method-access%28Office.15%29.aspx)|
|[RepaintObject](http://msdn.microsoft.com/library/docmd-repaintobject-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/docmd-requery-method-access%28Office.15%29.aspx)|
|[Restore](http://msdn.microsoft.com/library/docmd-restore-method-access%28Office.15%29.aspx)|
|[RunCommand](http://msdn.microsoft.com/library/docmd-runcommand-method-access%28Office.15%29.aspx)|
|[RunDataMacro](http://msdn.microsoft.com/library/docmd-rundatamacro-method-access%28Office.15%29.aspx)|
|[RunMacro](http://msdn.microsoft.com/library/docmd-runmacro-method-access%28Office.15%29.aspx)|
|[RunSavedImportExport](http://msdn.microsoft.com/library/docmd-runsavedimportexport-method-access%28Office.15%29.aspx)|
|[RunSQL](http://msdn.microsoft.com/library/docmd-runsql-method-access%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/docmd-save-method-access%28Office.15%29.aspx)|
|[SearchForRecord](http://msdn.microsoft.com/library/docmd-searchforrecord-method-access%28Office.15%29.aspx)|
|[SelectObject](http://msdn.microsoft.com/library/docmd-selectobject-method-access%28Office.15%29.aspx)|
|[SendObject](http://msdn.microsoft.com/library/docmd-sendobject-method-access%28Office.15%29.aspx)|
|[SetDisplayedCategories](http://msdn.microsoft.com/library/docmd-setdisplayedcategories-method-access%28Office.15%29.aspx)|
|[SetFilter](http://msdn.microsoft.com/library/docmd-setfilter-method-access%28Office.15%29.aspx)|
|[SetMenuItem](http://msdn.microsoft.com/library/docmd-setmenuitem-method-access%28Office.15%29.aspx)|
|[SetOrderBy](http://msdn.microsoft.com/library/docmd-setorderby-method-access%28Office.15%29.aspx)|
|[SetParameter](http://msdn.microsoft.com/library/docmd-setparameter-method-access%28Office.15%29.aspx)|
|[SetProperty](http://msdn.microsoft.com/library/docmd-setproperty-method-access%28Office.15%29.aspx)|
|[SetWarnings](http://msdn.microsoft.com/library/docmd-setwarnings-method-access%28Office.15%29.aspx)|
|[ShowAllRecords](http://msdn.microsoft.com/library/docmd-showallrecords-method-access%28Office.15%29.aspx)|
|[ShowToolbar](http://msdn.microsoft.com/library/docmd-showtoolbar-method-access%28Office.15%29.aspx)|
|[SingleStep](http://msdn.microsoft.com/library/docmd-singlestep-method-access%28Office.15%29.aspx)|
|[TransferDatabase](http://msdn.microsoft.com/library/docmd-transferdatabase-method-access%28Office.15%29.aspx)|
|[TransferSharePointList](http://msdn.microsoft.com/library/docmd-transfersharepointlist-method-access%28Office.15%29.aspx)|
|[TransferSpreadsheet](http://msdn.microsoft.com/library/docmd-transferspreadsheet-method-access%28Office.15%29.aspx)|
|[TransferSQLDatabase](http://msdn.microsoft.com/library/docmd-transfersqldatabase-method-access%28Office.15%29.aspx)|
|[TransferText](http://msdn.microsoft.com/library/docmd-transfertext-method-access%28Office.15%29.aspx)|

## See also


<<<<<<< HEAD
#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
=======
[Access Object Model Reference](object-model-access-vba-reference.md)
>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

