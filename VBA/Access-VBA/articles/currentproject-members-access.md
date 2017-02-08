---
title: CurrentProject Members (Access)
ms.prod: ACCESS
ms.assetid: adb319f1-487a-d7d1-5755-d57c31c776b8
---


# CurrentProject Members (Access)


The  **CurrentProject** object refers to the project for the current Microsoft Access project (.adp) or Access database.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddSharedImage](currentproject-addsharedimage-method-access.md)|Imports the the specified image into the database and adds it to the  **[SharedResources](sharedresources-object-access.md)** collection.|
|[CloseConnection](currentproject-closeconnection-method-access.md)|You can use the  **CloseConnection** method to close the current connection between the **CurrentProject** object in a Microsoft Access project (.adp) or Access database and the database specified in the project's base connection string.|
|[OpenConnection](currentproject-openconnection-method-access.md)|You can use the  **OpenConnection** method to open an ADO connection to an existing Microsoft Access project (.adp) or Access database as the current Access project or database in the Microsoft Access window.|
|[UpdateDependencyInfo](currentproject-updatedependencyinfo-method-access.md)|Updates the dependency information for the database.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AccessConnection](currentproject-accessconnection-property-access.md)|You can use the  **AccessConnection** property to return a reference to the current Microsoft ActiveX Data Objects (ADO) **Connection** object and its related properties. Read-only **Connection**.|
|[AllForms](currentproject-allforms-property-access.md)|You can use the  **AllForms** property to reference the **[AllForms](allforms-object-access.md)** collection and its related properties. Read-only **AllForms** object.|
|[AllMacros](currentproject-allmacros-property-access.md)|You can use the  **AllMacros** property to reference the **[AllMacros](allmacros-object-access.md)** collection and its related properties. Read-only **AllMacros** object.|
|[AllModules](currentproject-allmodules-property-access.md)|You can use the  **AllModules** property to reference the **[AllModules](allmodules-object-access.md)** collection and its related properties. Read-only **AllModules** object.|
|[AllReports](currentproject-allreports-property-access.md)|You can use the  **AllReports** property to reference the **[AllReports](allreports-object-access.md)** collection and its related properties. Read-only **AllReports** object.|
|[Application](currentproject-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.|
|[BaseConnectionString](currentproject-baseconnectionstring-property-access.md)|You can use the  **BaseConnectionString** property to return the base connection string for the specified object. Read-only **String**.|
|[Connection](currentproject-connection-property-access.md)|You can use the  **Connection** property to return a reference to the current ActiveX Data Objects (ADO) **Connection** object and its related properties. Read-only **Connection**.|
|[FileFormat](currentproject-fileformat-property-access.md)|Returns an  **[AcFileFormat](acfileformat-enumeration-access.md)** constant indicating the Microsoft Access version format of the specified project. Read-only.|
|[FullName](currentproject-fullname-property-access.md)|Sets or returns the full path (including file name) of a specific object. Read-only  **String**.|
|[ImportExportSpecifications](currentproject-importexportspecifications-property-access.md)|Returns a  **[ImportExportSpecifications](importexportspecifications-object-access.md)** collection that represents the collection of saved import or export operations for the specified object. Read-only.|
|[IsConnected](currentproject-isconnected-property-access.md)|You can use the  **IsConnected** property to determine if the **[CurrentProject](currentproject-object-access.md)** object is currently connected. Read-only **Boolean**.|
|[IsTrusted](currentproject-istrusted-property-access.md)|Gets whether or not macros and Visual Basic for Applications (VBA) code have been enabled in the current project. Read-only  **Boolean**.|
|[IsWeb](currentproject-isweb-property-access.md)|Gets whether the database is a Web database. Read-only  **Boolean**.|
|[Name](currentproject-name-property-access.md)|You can use the  **Name** property to determine the string expression that identifies the name of an object. Read-only **String**.|
|[Parent](currentproject-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Path](currentproject-path-property-access.md)|You can use the  **Path** property to determine the location where data is stored for a Microsoft Access project (.adp) or Microsoft Access database. Read-only **String**.|
|[ProjectType](currentproject-projecttype-property-access.md)|You can use the  **ProjectType** property to determine the type of project that is currently open. Read-only **[AcProjectType](acprojecttype-enumeration-access.md)**.|
|[Properties](currentproject-properties-property-access.md)|Returns a reference to a  **[CurrentProject](currentproject-object-access.md)** object's **[AccessObjectProperties](accessobjectproperties-object-access.md)** collection. Read-only.|
|[RemovePersonalInformation](currentproject-removepersonalinformation-property-access.md)|Returns or sets a  **Boolean** indicating whether personal information about the user is stored in the specified project. **True** if personal information is removed. Read-write.|
|[Resources](currentproject-resources-property-access.md)|Gets the  **[SharedResources](sharedresources-object-access.md)** collection for the specified object. Read-only **SharedResources**.|
|[WebSite](currentproject-website-property-access.md)|Gets the Uniform Resource Locator (URL) of the Web site to which the database has been published. Read-only  **String**.|
|[IsSQLBackend](currentproject-issqlbackend-property-access.md)|Returns  **true** if the current project was created in Access 2013 and onwards and **false** if the current project was created prior to Access 2013 . Read-only **Boolean** Introduced in Office 2016.|

