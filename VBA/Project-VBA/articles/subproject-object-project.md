---
title: Subproject Object (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.Subproject
ms.assetid: 1a3b0d18-6464-a4f2-479f-710e19faffa8
---


# Subproject Object (Project)



Represents a subproject. The  **Subproject** object is a member of the **[Subprojects](subprojects-object-project.md)** collection.
 **Using the Subproject Object**
Use  **Subprojects** ( _Index_ ), where _Index_ is the subproject index or project summary task name, to return a single **Subproject** object. The following example prevents changes made to the specified subproject in a master project from being automatically made to the source project.
 **Using the Subprojects Collection**
Use the  **[Subprojects](http://msdn.microsoft.com/library/project-subprojects-property-project%28Office.15%29.aspx)** property to return a **Subprojects** collection. The following example cautions the user if any of the subprojects in the active project are not on the hard disk.

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/subproject-application-property-project%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/subproject-index-property-project%28Office.15%29.aspx)|
|[InsertedProjectSummary](http://msdn.microsoft.com/library/subproject-insertedprojectsummary-property-project%28Office.15%29.aspx)|
|[IsLoaded](http://msdn.microsoft.com/library/subproject-isloaded-property-project%28Office.15%29.aspx)|
|[LinkToSource](http://msdn.microsoft.com/library/subproject-linktosource-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/subproject-parent-property-project%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/subproject-path-property-project%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/subproject-readonly-property-project%28Office.15%29.aspx)|
|[SourceProject](http://msdn.microsoft.com/library/subproject-sourceproject-property-project%28Office.15%29.aspx)|

