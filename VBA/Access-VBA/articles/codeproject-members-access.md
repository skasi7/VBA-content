---
title: CodeProject Members (Access)
ms.prod: ACCESS
ms.assetid: cd3b6b70-8312-2f2f-0f4d-7679d8bea9f5
---


# CodeProject Members (Access)


The  **CodeProject** object refers to the project for the code database of a Microsoft Access project (.adp) or Access database.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddSharedImage](codeproject-addsharedimage-method-access.md)|Imports the specified image into the database and adds it to the  **[SharedResources](sharedresources-object-access.md)** collection.|
|[CloseConnection](codeproject-closeconnection-method-access.md)|You can use the  **CloseConnection** method to close the current connection between the **CodeProject** object in a Microsoft Access project (.adp) or Access database and the database specified in the project's base connection string.|
|[OpenConnection](codeproject-openconnection-method-access.md)|You can use the  **OpenConnection** method to open an ADO connection to an existing Microsoft Access project (.adp) or Access database as the current Access project or database in the Microsoft Access window.|
|[UpdateDependencyInfo](codeproject-updatedependencyinfo-method-access.md)|Updates the dependency information for the database.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AccessConnection](codeproject-accessconnection-property-access.md)|You can use the  **AccessConnection** property to return a reference to the current Microsoft ActiveX Data Objects (ADO) **Connection** object and its related properties. Read-only **Connection**.|
|[AllForms](codeproject-allforms-property-access.md)|You can use the  **AllForms** property to reference the **[AllForms](allforms-object-access.md)** collection and its related properties. Read-only **AllForms** object.|
|[AllMacros](codeproject-allmacros-property-access.md)|You can use the  **AllMacros** property to reference the **[AllMacros](allmacros-object-access.md)** collection and its related properties. Read-only **AllMacros** object.|
|[AllModules](codeproject-allmodules-property-access.md)|You can use the  **AllModules** property to reference the **[AllModules](allmodules-object-access.md)** collection and its related properties. Read-only **AllModules** object.|
|[AllReports](codeproject-allreports-property-access.md)|You can use the  **AllReports** property to reference the **[AllReports](allreports-object-access.md)** collection and its related properties. Read-only **AllReports** object.|
|[Application](codeproject-application-property-access.md)|You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.|
|[BaseConnectionString](codeproject-baseconnectionstring-property-access.md)|You can use the  **BaseConnectionString** property to return the base connection string for the specified object. Read-only **String**.|
|[Connection](codeproject-connection-property-access.md)|You can use the  **Connection** property to return a reference to the current ActiveX Data Objects (ADO) **Connection** object and its related properties. Read-only **Connection**.|
|[FileFormat](codeproject-fileformat-property-access.md)|Returns an  **[AcFileFormat](acfileformat-enumeration-access.md)** constant indicating the Microsoft Access version format of the specified project. Read-only.|
|[FullName](codeproject-fullname-property-access.md)|Sets or returns the full path (including file name) of a specific object. Read-only  **String**.|
|[ImportExportSpecifications](codeproject-importexportspecifications-property-access.md)|Returns a  **[ImportExportSpecifications](importexportspecifications-object-access.md)** collection that represents the collection of saved import or export operations for the specified object. Read-only.|
|[IsConnected](codeproject-isconnected-property-access.md)|You can use the  **IsConnected** property to determine if the **[CodeProject](codeproject-object-access.md)** object is currently connected. Read-only **Boolean**.|
|[IsTrusted](codeproject-istrusted-property-access.md)|Gets whether or not macros and Visual Basic for Applications (VBA) code have been enabled in the current project. Read-only  **Boolean**.|
|[IsWeb](codeproject-isweb-property-access.md)|Gets whether the database is a Web database. Read-only  **Boolean**.|
|[Name](codeproject-name-property-access.md)|You can use the  **Name** property to determine the string expression that identifies the name of an object. Read-only **String**.|
|[Parent](codeproject-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Path](codeproject-path-property-access.md)|You can use the  **Path** property to determine the location where data is stored for a Microsoft Access project (.adp) or Microsoft Access database. Read-only **String**.|
|[ProjectType](codeproject-projecttype-property-access.md)|You can use the  **ProjectType** property to determine the type of project that is currently open. Read-only **[AcProjectType](acprojecttype-enumeration-access.md)**.|
|[Properties](codeproject-properties-property-access.md)|Returns a reference to a  **[CodeProject](codeproject-object-access.md)** object's **[AccessObjectProperties](accessobjectproperties-object-access.md)** collection. Read-only.|
|[RemovePersonalInformation](codeproject-removepersonalinformation-property-access.md)|Returns or sets a  **Boolean** indicating whether personal information about the user is stored in the specified project. **True** if personal information is removed. Read-write.|
|[Resources](codeproject-resources-property-access.md)|Gets the  **[SharedResources](sharedresources-object-access.md)** collection for the specified object. Read-only **SharedResources**.|
|[WebSite](codeproject-website-property-access.md)|Gets the Uniform Resource Locator (URL) of the Web site to which the database has been published. Read-only  **String**.|
|[IsSQLBackend](codeproject-issqlbackend-property-access.md)|Returns  **true** if the code project was created in Access 2013 and onwards and **false** if the code project was created prior to Access 2013 . Read-only **Boolean** Introduced in Office 2016.|

