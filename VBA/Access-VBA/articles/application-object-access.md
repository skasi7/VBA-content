---
title: Application Object (Access)
keywords: vbaac10.chm12627
f1_keywords:
- vbaac10.chm12627
ms.prod: ACCESS
api_name:
- Access.Application
ms.assetid: aefb0713-97e6-e2c7-e530-8fd2e1316a55
---


# Application Object (Access)

The  **Application** object refers to the active Microsoft Access application.


## Remarks

The  **Application** object contains all Access objects and collections.

You can use the  **Application** object to apply methods or property settings to the entire Access application. For example, you can use the **[SetOption](http://msdn.microsoft.com/library/application-setoption-method-access%28Office.15%29.aspx)** method of the **Application** object to set database options from Visual Basic. The following example shows how you can set the **Display Status Bar** check box on the **Current Database** tab of the **Access Options** dialog box.




```
Application.SetOption "Show Status Bar", True
```

Access is a COM component that supports Automation, formerly called OLE Automation. You can manipulate Access objects from another application that also supports Automation. To do this, you use the  **Application** object.

For example, Microsoft Visual Basic is a COM component. You can open anAccess database from Visual Basic and work with its objects. From Visual Basic, first create a reference to the Access object library. Then create a new instance of the  **Application** class and point an object variable to it, as in the following example:




```
Dim appAccess As New Access.Application
```

From applications that do not support the  **New** keyword, you can create a new instance of the **Application** class by using the **CreateObject** function:




```
Dim appAccess As Object 
Set appAccess = CreateObject("Access.Application")
```

After you create a new instance of the  **Application** class, you can open a database or create a new database, by using either the **[OpenCurrentDatabase](http://msdn.microsoft.com/library/application-opencurrentdatabase-method-access%28Office.15%29.aspx)** method or the **[NewCurrentDatabase](http://msdn.microsoft.com/library/application-newcurrentdatabase-method-access%28Office.15%29.aspx)** method. You can then set the properties of the **Application** object and call its methods. When you return a reference to the **CommandBars** object by using the **CommandBars** property of the **Application** object, you can access all Microsoft Office command bar objects and collections by using this reference.

You can also manipulate other Access objects through the  **Application** object. For example, by using the **[OpenForm](http://msdn.microsoft.com/library/docmd-openform-method-access%28Office.15%29.aspx)** method of the Access **[DoCmd](docmd-object-access.md)** object, you can open an Access form from Microsoft Office Excel:




```
appAccess.DoCmd.OpenForm "Orders"
```

For more information about creating a reference and controlling objects by using Automation, see the documentation for the application that is acting as the COM component.


## Methods



|**Name**|
|:-----|
|[AccessError](http://msdn.microsoft.com/library/application-accesserror-method-access%28Office.15%29.aspx)|
|[AddToFavorites](http://msdn.microsoft.com/library/application-addtofavorites-method-access%28Office.15%29.aspx)|
|[BuildCriteria](http://msdn.microsoft.com/library/application-buildcriteria-method-access%28Office.15%29.aspx)|
|[CloseCurrentDatabase](http://msdn.microsoft.com/library/application-closecurrentdatabase-method-access%28Office.15%29.aspx)|
|[CodeDb](http://msdn.microsoft.com/library/application-codedb-method-access%28Office.15%29.aspx)|
|[ColumnHistory](http://msdn.microsoft.com/library/application-columnhistory-method-access%28Office.15%29.aspx)|
|[CompactRepair](http://msdn.microsoft.com/library/application-compactrepair-method-access%28Office.15%29.aspx)|
|[ConvertAccessProject](http://msdn.microsoft.com/library/application-convertaccessproject-method-access%28Office.15%29.aspx)|
|[CreateAccessProject](http://msdn.microsoft.com/library/application-createaccessproject-method-access%28Office.15%29.aspx)|
|[CreateAdditionalData](http://msdn.microsoft.com/library/application-createadditionaldata-method-access%28Office.15%29.aspx)|
|[CreateControl](http://msdn.microsoft.com/library/application-createcontrol-method-access%28Office.15%29.aspx)|
|[CreateForm](http://msdn.microsoft.com/library/application-createform-method-access%28Office.15%29.aspx)|
|[CreateGroupLevel](http://msdn.microsoft.com/library/application-creategrouplevel-method-access%28Office.15%29.aspx)|
|[CreateReport](http://msdn.microsoft.com/library/application-createreport-method-access%28Office.15%29.aspx)|
|[CreateReportControl](http://msdn.microsoft.com/library/application-createreportcontrol-method-access%28Office.15%29.aspx)|
|[CurrentDb](http://msdn.microsoft.com/library/application-currentdb-method-access%28Office.15%29.aspx)|
|[CurrentUser](http://msdn.microsoft.com/library/application-currentuser-method-access%28Office.15%29.aspx)|
|[CurrentWebUser](http://msdn.microsoft.com/library/application-currentwebuser-method-access%28Office.15%29.aspx)|
|[CurrentWebUserGroups](http://msdn.microsoft.com/library/application-currentwebusergroups-method-access%28Office.15%29.aspx)|
|[DAvg](http://msdn.microsoft.com/library/application-davg-method-access%28Office.15%29.aspx)|
|[DCount](http://msdn.microsoft.com/library/application-dcount-method-access%28Office.15%29.aspx)|
|[DDEExecute](http://msdn.microsoft.com/library/application-ddeexecute-method-access%28Office.15%29.aspx)|
|[DDEInitiate](http://msdn.microsoft.com/library/application-ddeinitiate-method-access%28Office.15%29.aspx)|
|[DDEPoke](http://msdn.microsoft.com/library/application-ddepoke-method-access%28Office.15%29.aspx)|
|[DDERequest](http://msdn.microsoft.com/library/application-dderequest-method-access%28Office.15%29.aspx)|
|[DDETerminate](http://msdn.microsoft.com/library/application-ddeterminate-method-access%28Office.15%29.aspx)|
|[DDETerminateAll](http://msdn.microsoft.com/library/application-ddeterminateall-method-access%28Office.15%29.aspx)|
|[DefaultWorkspaceClone](http://msdn.microsoft.com/library/application-defaultworkspaceclone-method-access%28Office.15%29.aspx)|
|[DeleteControl](http://msdn.microsoft.com/library/application-deletecontrol-method-access%28Office.15%29.aspx)|
|[DeleteReportControl](http://msdn.microsoft.com/library/application-deletereportcontrol-method-access%28Office.15%29.aspx)|
|[DFirst](http://msdn.microsoft.com/library/application-dfirst-method-access%28Office.15%29.aspx)|
|[DirtyObject](http://msdn.microsoft.com/library/application-dirtyobject-method-access%28Office.15%29.aspx)|
|[DLast](http://msdn.microsoft.com/library/application-dlast-method-access%28Office.15%29.aspx)|
|[DLookup](http://msdn.microsoft.com/library/application-dlookup-method-access%28Office.15%29.aspx)|
|[DMax](http://msdn.microsoft.com/library/application-dmax-method-access%28Office.15%29.aspx)|
|[DMin](http://msdn.microsoft.com/library/application-dmin-method-access%28Office.15%29.aspx)|
|[DStDev](http://msdn.microsoft.com/library/application-dstdev-method-access%28Office.15%29.aspx)|
|[DStDevP](http://msdn.microsoft.com/library/application-dstdevp-method-access%28Office.15%29.aspx)|
|[DSum](http://msdn.microsoft.com/library/application-dsum-method-access%28Office.15%29.aspx)|
|[DVar](http://msdn.microsoft.com/library/application-dvar-method-access%28Office.15%29.aspx)|
|[DVarP](http://msdn.microsoft.com/library/application-dvarp-method-access%28Office.15%29.aspx)|
|[Echo](http://msdn.microsoft.com/library/application-echo-method-access%28Office.15%29.aspx)|
|[EuroConvert](http://msdn.microsoft.com/library/application-euroconvert-method-access%28Office.15%29.aspx)|
|[Eval](http://msdn.microsoft.com/library/application-eval-method-access%28Office.15%29.aspx)|
|[ExportNavigationPane](http://msdn.microsoft.com/library/application-exportnavigationpane-method-access%28Office.15%29.aspx)|
|[ExportXML](http://msdn.microsoft.com/library/application-exportxml-method-access%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/application-followhyperlink-method-access%28Office.15%29.aspx)|
|[GetHiddenAttribute](http://msdn.microsoft.com/library/application-gethiddenattribute-method-access%28Office.15%29.aspx)|
|[GetOption](http://msdn.microsoft.com/library/application-getoption-method-access%28Office.15%29.aspx)|
|[GUIDFromString](http://msdn.microsoft.com/library/application-guidfromstring-method-access%28Office.15%29.aspx)|
|[HtmlEncode](http://msdn.microsoft.com/library/application-htmlencode-method-access%28Office.15%29.aspx)|
|[hWndAccessApp](http://msdn.microsoft.com/library/application-hwndaccessapp-method-access%28Office.15%29.aspx)|
|[HyperlinkPart](http://msdn.microsoft.com/library/application-hyperlinkpart-method-access%28Office.15%29.aspx)|
|[ImportNavigationPane](http://msdn.microsoft.com/library/application-importnavigationpane-method-access%28Office.15%29.aspx)|
|[ImportXML](http://msdn.microsoft.com/library/application-importxml-method-access%28Office.15%29.aspx)|
|[InstantiateTemplate](http://msdn.microsoft.com/library/application-instantiatetemplate-method-access%28Office.15%29.aspx)|
|[IsCurrentWebUserInGroup](http://msdn.microsoft.com/library/application-iscurrentwebuseringroup-method-access%28Office.15%29.aspx)|
|[LoadCustomUI](http://msdn.microsoft.com/library/application-loadcustomui-method-access%28Office.15%29.aspx)|
|[LoadFromAXL](http://msdn.microsoft.com/library/application-loadfromaxl-method-access%28Office.15%29.aspx)|
|[LoadPicture](http://msdn.microsoft.com/library/application-loadpicture-method-access%28Office.15%29.aspx)|
|[NewAccessProject](http://msdn.microsoft.com/library/application-newaccessproject-method-access%28Office.15%29.aspx)|
|[NewCurrentDatabase](http://msdn.microsoft.com/library/application-newcurrentdatabase-method-access%28Office.15%29.aspx)|
|[Nz](http://msdn.microsoft.com/library/application-nz-method-access%28Office.15%29.aspx)|
|[OpenAccessProject](http://msdn.microsoft.com/library/application-openaccessproject-method-access%28Office.15%29.aspx)|
|[OpenCurrentDatabase](http://msdn.microsoft.com/library/application-opencurrentdatabase-method-access%28Office.15%29.aspx)|
|[PlainText](http://msdn.microsoft.com/library/application-plaintext-method-access%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application-quit-method-access%28Office.15%29.aspx)|
|[RefreshDatabaseWindow](http://msdn.microsoft.com/library/application-refreshdatabasewindow-method-access%28Office.15%29.aspx)|
|[RefreshTitleBar](http://msdn.microsoft.com/library/application-refreshtitlebar-method-access%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/application-run-method-access%28Office.15%29.aspx)|
|[RunCommand](http://msdn.microsoft.com/library/application-runcommand-method-access%28Office.15%29.aspx)|
|[SaveAsAXL](http://msdn.microsoft.com/library/application-saveasaxl-method-access%28Office.15%29.aspx)|
|[SaveAsTemplate](http://msdn.microsoft.com/library/application-saveastemplate-method-access%28Office.15%29.aspx)|
|[SetDefaultWorkgroupFile](http://msdn.microsoft.com/library/application-setdefaultworkgroupfile-method-access%28Office.15%29.aspx)|
|[SetHiddenAttribute](http://msdn.microsoft.com/library/application-sethiddenattribute-method-access%28Office.15%29.aspx)|
|[SetOption](http://msdn.microsoft.com/library/application-setoption-method-access%28Office.15%29.aspx)|
|[StringFromGUID](http://msdn.microsoft.com/library/application-stringfromguid-method-access%28Office.15%29.aspx)|
|[SysCmd](http://msdn.microsoft.com/library/application-syscmd-method-access%28Office.15%29.aspx)|
|[TransformXML](http://msdn.microsoft.com/library/application-transformxml-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/application-application-property-access%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/application-assistance-property-access%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/application-autocorrect-property-access%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/application-automationsecurity-property-access%28Office.15%29.aspx)|
|[BrokenReference](http://msdn.microsoft.com/library/application-brokenreference-property-access%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/application-build-property-access%28Office.15%29.aspx)|
|[CodeContextObject](http://msdn.microsoft.com/library/application-codecontextobject-property-access%28Office.15%29.aspx)|
|[CodeData](http://msdn.microsoft.com/library/application-codedata-property-access%28Office.15%29.aspx)|
|[CodeProject](http://msdn.microsoft.com/library/application-codeproject-property-access%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/application-comaddins-property-access%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/application-commandbars-property-access%28Office.15%29.aspx)|
|[CurrentData](http://msdn.microsoft.com/library/application-currentdata-property-access%28Office.15%29.aspx)|
|[CurrentObjectName](http://msdn.microsoft.com/library/application-currentobjectname-property-access%28Office.15%29.aspx)|
|[CurrentObjectType](http://msdn.microsoft.com/library/application-currentobjecttype-property-access%28Office.15%29.aspx)|
|[CurrentProject](http://msdn.microsoft.com/library/application-currentproject-property-access%28Office.15%29.aspx)|
|[DBEngine](http://msdn.microsoft.com/library/application-dbengine-property-access%28Office.15%29.aspx)|
|[DoCmd](http://msdn.microsoft.com/library/application-docmd-property-access%28Office.15%29.aspx)|
|[FeatureInstall](http://msdn.microsoft.com/library/application-featureinstall-property-access%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/application-filedialog-property-access%28Office.15%29.aspx)|
|[Forms](http://msdn.microsoft.com/library/application-forms-property-access%28Office.15%29.aspx)|
|[IsCompiled](http://msdn.microsoft.com/library/application-iscompiled-property-access%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/application-languagesettings-property-access%28Office.15%29.aspx)|
|[MacroError](http://msdn.microsoft.com/library/application-macroerror-property-access%28Office.15%29.aspx)|
|[MenuBar](http://msdn.microsoft.com/library/application-menubar-property-access%28Office.15%29.aspx)|
|[Modules](http://msdn.microsoft.com/library/application-modules-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/application-name-property-access%28Office.15%29.aspx)|
|[NewFileTaskPane](http://msdn.microsoft.com/library/application-newfiletaskpane-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/application-parent-property-access%28Office.15%29.aspx)|
|[Printer](http://msdn.microsoft.com/library/application-printer-property-access%28Office.15%29.aspx)|
|[Printers](http://msdn.microsoft.com/library/application-printers-property-access%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/application-productcode-property-access%28Office.15%29.aspx)|
|[References](http://msdn.microsoft.com/library/application-references-property-access%28Office.15%29.aspx)|
|[Reports](http://msdn.microsoft.com/library/application-reports-property-access%28Office.15%29.aspx)|
|[ReturnVars](http://msdn.microsoft.com/library/application-returnvars-property-access%28Office.15%29.aspx)|
|[Screen](http://msdn.microsoft.com/library/application-screen-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/application-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[TempVars](http://msdn.microsoft.com/library/application-tempvars-property-access%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/application-usercontrol-property-access%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/application-vbe-property-access%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/application-version-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/application-visible-property-access%28Office.15%29.aspx)|
|[WebServices](http://msdn.microsoft.com/library/application-webservices-property-access%28Office.15%29.aspx)|

## See also


<<<<<<< HEAD
#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
=======
[Access Object Model Reference](object-model-access-vba-reference.md)
>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

