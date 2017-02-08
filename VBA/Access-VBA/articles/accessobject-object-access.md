---
title: AccessObject Object (Access)
keywords: vbaac10.chm12743
f1_keywords:
- vbaac10.chm12743
ms.prod: ACCESS
api_name:
- Access.AccessObject
ms.assetid: 8a770b33-5bff-120a-6707-ca214ee5ced3
---


# AccessObject Object (Access)

An  **AccessObject** object refers to a particular Access object.


## Remarks

An  **AccessObject** object includes information about one instance of an object. The following table list the types of objects each **AccessObject** describes, the name of its collection, and what type of information **AccessObject** contains.



|**AccessObject**|**Collection**|**Contains information about**|
|:-----|:-----|:-----|
|**Database diagram**|**AllDatabaseDiagrams**|Saved database diagrams|
|**Form**|**AllForms**|Saved forms|
|**Function**|**AllFunctions**|Saved functions|
|**Macro**|**AllMacros**|Saved macros|
|**Module**|**AllModules**|Saved modules|
|**Query**|**AllQueries**|Saved queries|
|**Report**|**AllReports**|Saved reports|
|**Stored procedure**|**AllStoredProcedures**|Saved stored procedures|
|**Table**|**AllTables**|Saved tables|
|**View**|**AllViews**|Saved views|
Because an  **AccessObject** object corresponds to an existing object, you can't create new **AccessObject** objects or delete existing ones. To refer to an **AccessObject** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms:


||
|:-----|
|**AllForms** (0)|
|**AllForms** (" _name_ ")|
|**AllForms** ![ _name_ ]|

## Methods



|**Name**|
|:-----|
|[GetDependencyInfo](http://msdn.microsoft.com/library/accessobject-getdependencyinfo-method-access%28Office.15%29.aspx)|
|[IsDependentUpon](http://msdn.microsoft.com/library/accessobject-isdependentupon-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[CurrentView](http://msdn.microsoft.com/library/accessobject-currentview-property-access%28Office.15%29.aspx)|
|[DateCreated](http://msdn.microsoft.com/library/accessobject-datecreated-property-access%28Office.15%29.aspx)|
|[DateModified](http://msdn.microsoft.com/library/accessobject-datemodified-property-access%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/accessobject-fullname-property-access%28Office.15%29.aspx)|
|[IsLoaded](http://msdn.microsoft.com/library/accessobject-isloaded-property-access%28Office.15%29.aspx)|
|[IsWeb](http://msdn.microsoft.com/library/accessobject-isweb-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/accessobject-name-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/accessobject-parent-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/accessobject-properties-property-access%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/accessobject-type-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[AccessObject Object Members](http://msdn.microsoft.com/library/accessobject-members-access%28Office.15%29.aspx)
