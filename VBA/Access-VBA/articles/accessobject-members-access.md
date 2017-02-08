---
title: AccessObject Members (Access)
ms.prod: ACCESS
ms.assetid: 78aaacb1-c0d3-d809-088d-d543ecd71de3
---


# AccessObject Members (Access)


An  **AccessObject** object refers to a particular Access object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[GetDependencyInfo](accessobject-getdependencyinfo-method-access.md)| Returns a **[DependencyInfo](dependencyinfo-object-access.md)** object that represents the database objects that are dependent upon the specified object.|
|[IsDependentUpon](accessobject-isdependentupon-method-access.md)|Returns a  **Boolean** value that indicates whether the specified object is dependent upon the database object specified in the _ObjectName_ argument.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[CurrentView](accessobject-currentview-property-access.md)|Returns the current view for the specified Access object. Read-only  **[AcCurrentView](accurrentview-enumeration-access.md)**.|
|[DateCreated](accessobject-datecreated-property-access.md)|Returns a  **Date** indicating the date and time when the design of the specified object was last modified. Read-only.|
|[DateModified](accessobject-datemodified-property-access.md)|Returns a  **Date** indicating the date and time when the design of the specified object was last modified. Read-only.|
|[FullName](accessobject-fullname-property-access.md)|Sets or returns the full path (including file name) of a specific object. Read/write  **String**.|
|[IsLoaded](accessobject-isloaded-property-access.md)|You can use the  **IsLoaded** property to determine if an **[AccessObject](accessobject-object-access.md)** is currently loaded. Read-only **Boolean**.|
|[IsWeb](accessobject-isweb-property-access.md)|Gets whether the specified object is a Web object. Read-only  **Boolean**.|
|[Name](accessobject-name-property-access.md)|You can use the  **Name** property to determine the string expression that identifies the name of an object. Read-only **String**.|
|[Parent](accessobject-parent-property-access.md)|Returns the parent object for the specified object. Read-only.|
|[Properties](accessobject-properties-property-access.md)|Returns a reference to a  **[AccessObject](accessobject-object-access.md)** object's **[AccessObjectProperties](accessobjectproperties-object-access.md)** collection. Read-only.|
|[Type](accessobject-type-property-access.md)|Returns the value of an  **[AccessObject](accessobject-object-access.md)** object type. Read-only **[AcObjectType](acobjecttype-enumeration-access.md)**.|

