---
title: Workbook Properties (Excel)
ms.prod: EXCEL
ms.assetid: 6bbb7ce8-a8dd-459f-b440-690a9ffe30e8
---


# Workbook Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AccuracyVersion](workbook-accuracyversion-property-excel.md)|Specifies whether certain worksheet functions use the latest accuracy algorithms to calculate their results. Read/write|
|[ActiveChart](workbook-activechart-property-excel.md)|Returns a  **[Chart](chart-object-excel.md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing** .|
|[ActiveSheet](workbook-activesheet-property-excel.md)|Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns  **Nothing** if no sheet is active.|
|[ActiveSlicer](workbook-activeslicer-property-excel.md)|Returns an object that represents the active slicer in the active workbook or in the specified workbook. Returns  **Nothing** if no slicer is active. Read-only.|
|[Application](workbook-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoUpdateFrequency](workbook-autoupdatefrequency-property-excel.md)|Returns or sets the number of minutes between automatic updates to the shared workbook. Read/write  **Long** .|
|[AutoUpdateSaveChanges](workbook-autoupdatesavechanges-property-excel.md)| **True** if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. **False** if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is **True** . Read/write **Boolean** .|
|[BuiltinDocumentProperties](workbook-builtindocumentproperties-property-excel.md)|Returns a  **[DocumentProperties](documentproperties-object-office.md)** collection that represents all the built-in document properties for the specified workbook. Read-only.|
|[CalculationVersion](workbook-calculationversion-property-excel.md)|Returns the information about the version of Excel that the workbook was last fully recalculated by. Read-only  **Long** .|
|[CaseSensitive](workbook-casesensitive-property-excel.md)| **True** if the workbook distinguishes between upper and lower case when comparing content. Read-only **Boolean**|
|[ChangeHistoryDuration](workbook-changehistoryduration-property-excel.md)|Returns or sets the number of days shown in the shared workbook's change history. Read/write  **Long** .|
|[ChartDataPointTrack](workbook-chartdatapointtrack-property-excel.md)| **True** will cause all charts in the current document to track the actual data point to which it's attached. **False** will revert back to tracking the index of the data point. **Boolean** Read/Write|
|[Charts](workbook-charts-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the chart sheets in the specified workbook.|
|[CheckCompatibility](workbook-checkcompatibility-property-excel.md)|Controls whether or not the compatibility checker is run automatically when the workbook is saved. Read/write  **Boolean** .|
|[CodeName](workbook-codename-property-excel.md)|Returns the code name for the object. Read-only  **String** .|
|[Colors](workbook-colors-property-excel.md)|Returns or sets colors in the palette for the workbook. The palette has 56 entries, each represented by an RGB value. Read/write  **Variant** .|
|[CommandBars](workbook-commandbars-property-excel.md)|Returns a  **[CommandBars](commandbars-object-office.md)** object that represents the Microsoft Excel command bars. Read-only.|
|[ConflictResolution](workbook-conflictresolution-property-excel.md)|Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated. Read/write  **[XlSaveConflictResolution](xlsaveconflictresolution-enumeration-excel.md)** .|
|[Connections](workbook-connections-property-excel.md)|The  **Connections** property establishes a connection between the workbook and an ODBC or an OLEDB data source and refreshes the data without prompting the user. Read-only.|
|[ConnectionsDisabled](workbook-connectionsdisabled-property-excel.md)|Disables the external connections or links in the workbook. Read-only|
|[Container](workbook-container-property-excel.md)|Returns the object that represents the container application for the specified OLE object. Read-only  **Object** .|
|[ContentTypeProperties](workbook-contenttypeproperties-property-excel.md)|Returns a  **[MetaProperties](metaproperties-object-office.md)** collection that describes the metadata stored in the workbook. Read-only.|
|[CreateBackup](workbook-createbackup-property-excel.md)| **True** if a backup file is created when this file is saved. Read-only **Boolean** .|
|[Creator](workbook-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[CustomDocumentProperties](workbook-customdocumentproperties-property-excel.md)|Returns or sets a  **[DocumentProperties](documentproperties-object-office.md)** collection that represents all the custom document properties for the specified workbook.|
|[CustomViews](workbook-customviews-property-excel.md)|Returns a  **[CustomViews](customviews-object-excel.md)** collection that represents all the custom views for the workbook.|
|[CustomXMLParts](workbook-customxmlparts-property-excel.md)|Returns a  **[CustomXMLParts](customxmlparts-object-office.md)** collection that represents the custom XML in the XML data store. Read-only.|
|[Date1904](workbook-date1904-property-excel.md)| **True** if the workbook uses the 1904 date system. Read/write **Boolean** .|
|[DefaultPivotTableStyle](workbook-defaultpivottablestyle-property-excel.md)|Specifies the table style from the  **TableStyles** collection that is used as the default style for PivotTables. Read/write.|
|[DefaultSlicerStyle](workbook-defaultslicerstyle-property-excel.md)|Specifies the style from the  **[TableStyles](tablestyles-object-excel.md)** object that is used as the default style for slicers. Read/write.|
|[DefaultTableStyle](workbook-defaulttablestyle-property-excel.md)|Specifies the table style from the  **TableStyles** collection that is used as the default TableStyle. Read/write **Variant** .|
|[DefaultTimelineStyle](workbook-defaulttimelinestyle-property-excel.md)|The name of the default slicer style of the workbook.  **Variant**. Read/Write|
|[DisplayDrawingObjects](workbook-displaydrawingobjects-property-excel.md)|Returns or sets how shapes are displayed. Read/write  **Long** .|
|[DisplayInkComments](workbook-displayinkcomments-property-excel.md)|A  **Boolean** value that determines whether ink comments are displayed in the workbook. Read/write **Boolean** .|
|[DocumentInspectors](workbook-documentinspectors-property-excel.md)|Returns a  **[DocumentInspectors](documentinspectors-object-office.md)** collection that represents the Document Inspector modules for the specified workbook. Read-only.|
|[DocumentLibraryVersions](workbook-documentlibraryversions-property-excel.md)|Returns a  **[DocumentLibraryVersions](documentlibraryversions-object-office.md)** collection that represents the collection of versions of a shared workbook that has versioning enabled and that is stored in a document library on a server.|
|[DoNotPromptForConvert](workbook-donotpromptforconvert-property-excel.md)|Returns or sets if the user should be prompted to convert the workbook if the workbook contains features that are not supported by versions of Excel earlier than Excel 2007. Read/write  **Boolean** .|
|[EnableAutoRecover](workbook-enableautorecover-property-excel.md)|Saves changed files, of all formats, on a timed interval. Read/write  **Boolean** .|
|[EncryptionProvider](workbook-encryptionprovider-property-excel.md)|Returns a  **String** specifying the name of the algorithm encryption provider that Microsoft Office Excel 2007 uses when encrypting documents. Read/write.|
|[EnvelopeVisible](workbook-envelopevisible-property-excel.md)| **True** if the e-mail composition header and the envelope toolbar are both visible. Read/write **Boolean** .|
|[Excel4IntlMacroSheets](workbook-excel4intlmacrosheets-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.|
|[Excel4MacroSheets](workbook-excel4macrosheets-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.|
|[Excel8CompatibilityMode](workbook-excel8compatibilitymode-property-excel.md)|The  **Excel8CompatibilityMode** property provides developers with a way to check if the workbook is in compatibility mode. Read-only **Boolean** .|
|[FileFormat](workbook-fileformat-property-excel.md)|Returns the file format and/or type of the workbook. Read-only  **[XlFileFormat](xlfileformat-enumeration-excel.md)** .|
|[Final](workbook-final-property-excel.md)|Returns or sets a  **Boolean** that indicates whether a workbook is final. Read/write **Boolean** .|
|[ForceFullCalculation](workbook-forcefullcalculation-property-excel.md)|Returns or sets the specified workbook to forced calculation mode. Read/write.|
|[FullName](workbook-fullname-property-excel.md)|Returns the name of the object, including its path on disk, as a string. Read-only  **String** .|
|[FullNameURLEncoded](workbook-fullnameurlencoded-property-excel.md)|Returns a  **String** indicating the name of the object, including its path on disk, as a string. Read-only.|
|[HasPassword](workbook-haspassword-property-excel.md)| **True** if the workbook has a protection password. Read-only **Boolean** .|
|[HasVBProject](workbook-hasvbproject-property-excel.md)|Returns a  **Boolean** that represents whether a workbook has an attached Microsoft Visual Basic for Applications project. Read-only **Boolean** .|
|[HighlightChangesOnScreen](workbook-highlightchangesonscreen-property-excel.md)| **True** if changes to the shared workbook are highlighted on-screen. Read/write **Boolean** .|
|[IconSets](workbook-iconsets-property-excel.md)|This property is used to filter data in a workbook based on a cell icon from the  **IconSet** collection. Read-only.|
|[InactiveListBorderVisible](workbook-inactivelistbordervisible-property-excel.md)|A  **Boolean** value that specifies whether list borders are visible when a list is not active. Returns **True** if the border is visible. Read/write **Boolean** .|
|[IsAddin](workbook-isaddin-property-excel.md)| **True** if the workbook is running as an add-in. Read/write **Boolean** .|
|[IsInplace](workbook-isinplace-property-excel.md)| **True** if the specified workbook is being edited in place. **False** if the workbook has been opened in Microsoft Excel for editing. Read-only **Boolean** .|
|[KeepChangeHistory](workbook-keepchangehistory-property-excel.md)| **True** if change tracking is enabled for the shared workbook. Read/write **Boolean** .|
|[ListChangesOnNewSheet](workbook-listchangesonnewsheet-property-excel.md)| **True** if changes to the shared workbook are shown on a separate worksheet. Read/write **Boolean** .|
|[Mailer](workbook-mailer-property-excel.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
|[Model](workbook-model-property-excel.md)|Returns the top level  **Model** object which is the one Data Model for the workbook. Read-only|
|[MultiUserEditing](workbook-multiuserediting-property-excel.md)| **True** if the workbook is open as a shared list. Read-only **Boolean** .|
|[Name](workbook-name-property-excel.md)|Returns a  **String** value that represents the name of the object.|
|[Names](workbook-names-property-excel.md)|Returns a  **[Names](names-object-excel.md)** collection that represents all the names in the specified workbook (including all worksheet-specific names). Read-only **Names** object.|
|[Parent](workbook-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Password](workbook-password-property-excel.md)|Returns or sets the password that must be supplied to open the specified workbook. Read/write  **String** .|
|[PasswordEncryptionAlgorithm](workbook-passwordencryptionalgorithm-property-excel.md)|Returns a  **String** indicating the algorithm Microsoft Excel uses to encrypt passwords for the specified workbook. Read-only.|
|[PasswordEncryptionFileProperties](workbook-passwordencryptionfileproperties-property-excel.md)| **True** if Microsoft Excel encrypts file properties for the specified password-protected workbook. Read-only **Boolean** .|
|[PasswordEncryptionKeyLength](workbook-passwordencryptionkeylength-property-excel.md)|Returns a  **Long** indicating the key length of the algorithm Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.|
|[PasswordEncryptionProvider](workbook-passwordencryptionprovider-property-excel.md)|Returns a  **String** specifying the name of the algorithm encryption provider that Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.|
|[Path](workbook-path-property-excel.md)|Returns a  **String** that represents the complete path to the workbook/file that this workbook object respresents.|
|[Permission](workbook-permission-property-excel.md)|Returns a  **Permission** object that represents the permission settings in the specified workbook.|
|[PersonalViewListSettings](workbook-personalviewlistsettings-property-excel.md)| **True** if filter and sort settings for lists are included in the user's personal view of the shared workbook. Read/write **Boolean** .|
|[PersonalViewPrintSettings](workbook-personalviewprintsettings-property-excel.md)| **True** if print settings are included in the user's personal view of the shared workbook. Read-write **Boolean** .|
|[PivotTables](workbook-pivottables-property-excel.md)|Returns an object that represents a collection of all the PivotTable reports on a worksheet. Read-only.|
|[PrecisionAsDisplayed](workbook-precisionasdisplayed-property-excel.md)| **True** if calculations in this workbook will be done using only the precision of the numbers as they're displayed. Read/write **Boolean** .|
|[ProtectStructure](workbook-protectstructure-property-excel.md)| **True** if the order of the sheets in the workbook is protected. Read-only **Boolean** .|
|[ProtectWindows](workbook-protectwindows-property-excel.md)| **True** if the windows of the workbook are protected. Read-only **Boolean** .|
|[PublishObjects](workbook-publishobjects-property-excel.md)|Returns the  **[PublishObjects](publishobjects-object-excel.md)** collection. Read-only.|
|[ReadOnly](workbook-readonly-property-excel.md)| Returns **True** if the object has been opened as read-only. Read-only **Boolean** .|
|[ReadOnlyRecommended](workbook-readonlyrecommended-property-excel.md)| **True** if the workbook was saved as read-only recommended. Read-only **Boolean** .|
|[RemovePersonalInformation](workbook-removepersonalinformation-property-excel.md)| **True** if personal information can be removed from the specified workbook. The default value is **False** . Read/write **Boolean** .|
|[Research](workbook-research-property-excel.md)|Returns a  **Research** object that represents the research service for a workbook. Read-only.|
|[RevisionNumber](workbook-revisionnumber-property-excel.md)|Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns 0 (zero). Read-only  **Long** .|
|[Saved](workbook-saved-property-excel.md)| **True** if no changes have been made to the specified workbook since it was last saved. Read/write **Boolean** .|
|[SaveLinkValues](workbook-savelinkvalues-property-excel.md)| **True** if Microsoft Excel saves external link values with the workbook. Read/write **Boolean** .|
|[ServerPolicy](workbook-serverpolicy-property-excel.md)|Returns a  **ServerPolicy** object that represents a policy specified for a workbook stored on a server running SharePoint Server 2007 or later. Read-only.|
|[ServerViewableItems](workbook-serverviewableitems-property-excel.md)|Allows a developer to interact with the list of published objects in the workbook that are shown on the server. Read-only.|
|[SharedWorkspace](workbook-sharedworkspace-property-excel.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|
|[Sheets](workbook-sheets-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the sheets in the specified workbook. Read-only **Sheets** object.|
|[ShowConflictHistory](workbook-showconflicthistory-property-excel.md)| **True** if the Conflict History worksheet is visible in the workbook that's open as a shared list. Read/write **Boolean** .|
|[ShowPivotChartActiveFields](workbook-showpivotchartactivefields-property-excel.md)|This property controls the visibility of the PivotChart Filter Pane. Read/write  **Boolean** .|
|[ShowPivotTableFieldList](workbook-showpivottablefieldlist-property-excel.md)| **True** (default) if the PivotTable field list can be shown. Read/write **Boolean** .|
|[Signatures](workbook-signatures-property-excel.md)|Returns the digital signatures for a workbook. Read-only.|
|[SlicerCaches](workbook-slicercaches-property-excel.md)|Returns the  **[SlicerCaches](slicercaches-object-excel.md)** object associated with the workbook. Read-only.|
|[SmartDocument](workbook-smartdocument-property-excel.md)|Returns a  **SmartDocument** object that represents the settings for a smart document solution. Read-only.|
|[Styles](workbook-styles-property-excel.md)|Returns a  **[Styles](styles-object-excel.md)** collection that represents all the styles in the specified workbook. Read-only.|
|[Sync](workbook-sync-property-excel.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|
|[TableStyles](workbook-tablestyles-property-excel.md)|Returns a  **TableStyles** collection object for the current workbook that refers to the styles used in the current workbook. Read-only.|
|[TemplateRemoveExtData](workbook-templateremoveextdata-property-excel.md)| **True** if external data references are removed when the workbook is saved as a template. Read/write **Boolean** .|
|[Theme](workbook-theme-property-excel.md)|Returns the theme applied to the current workbook. Read-only.|
|[UpdateLinks](workbook-updatelinks-property-excel.md)|Returns or sets an  **[XlUpdateLink](xlupdatelinks-enumeration-excel.md)** constant indicating a workbook's setting for updating embedded OLE links. Read/write.|
|[UpdateRemoteReferences](workbook-updateremotereferences-property-excel.md)| **True** if Microsoft Excel updates remote references in the workbook. Read/write **Boolean** .|
|[UserStatus](workbook-userstatus-property-excel.md)|Returns a 1-based, two-dimensional array that provides information about each user who has the workbook open as a shared list. Read-only  **Variant** .|
|[UseWholeCellCriteria](workbook-usewholecellcriteria-property-excel.md)| **True** if the workbook uses search patterns that match the entire content of a cell. Read-only **Boolean** .|
|[UseWildcards](workbook-usewildcards-property-excel.md)| **True** if the workbook enables wildcards for character string comparisons and searching. Read-only **Boolean**|
|[VBASigned](workbook-vbasigned-property-excel.md)| **True** if the Visual Basic for Applications project for the specified workbook has been digitally signed. Read-only **Boolean** .|
|[VBProject](workbook-vbproject-property-excel.md)|Returns a  **VBProject** object that represents the Visual Basic project in the specified workbook. Read-only.|
|[WebOptions](workbook-weboptions-property-excel.md)|Returns the  **[WebOptions](weboptions-object-excel.md)** collection, which contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page. Read-only.|
|[Windows](workbook-windows-property-excel.md)|Returns a  **[Windows](windows-object-excel.md)** collection that represents all the windows in the specified workbook. Read-only **Windows** object.|
|[Worksheets](workbook-worksheets-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the worksheets in the specified workbook. Read-only **Sheets** object.|
|[WritePassword](workbook-writepassword-property-excel.md)|Returns or sets a  **String** for the write password of a workbook. Read/write.|
|[WriteReserved](workbook-writereserved-property-excel.md)| **True** if the workbook is write-reserved. Read-only **Boolean** .|
|[WriteReservedBy](workbook-writereservedby-property-excel.md)|Returns the name of the user who currently has write permission for the workbook. Read-only  **String** .|
|[XmlMaps](workbook-xmlmaps-property-excel.md)| Returns an **XmlMaps** collection that represents the schema maps that have been added to the specified workbook. Read-only.|
|[XmlNamespaces](workbook-xmlnamespaces-property-excel.md)|Returns an  **[XmlNamespaces](xmlnamespaces-object-excel.md)** collection that represents the XML namespaces contained in the specified workbook. Read-only.|
|[Queries](workbook-queries-property-excel.md)||

