---
title: TaskDependency Object (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.TaskDependency
ms.assetid: 05d759fb-0203-761e-10f3-65b07d233f4d
---


# TaskDependency Object (Project)



Represents the link type and link lag information between two tasks. The  **TaskDependency** object is a member of the **[TaskDependencies](taskdependencies-object-project.md)** collection.
 **Using the TaskDependency Object**
Use  **TaskDependencies** ( _Index_ ), where _Index_ is the dependency index, to return a single **TaskDependency** object. The following example adds 1.5 days of lag to the link between the specified task and the predecessor specified in its first task dependency.
 **Using the TaskDependencies Collection**
Use the  **[TaskDependencies](http://msdn.microsoft.com/library/task-taskdependencies-property-project%28Office.15%29.aspx)** property to return a **TaskDependencies** collection. The following example examines each predecessor for the specified task and displays a message for each that has a priority of "High" or better.
Use the  **[Add](http://msdn.microsoft.com/library/taskdependencies-add-method-project%28Office.15%29.aspx)** method to add a **TaskDependency** object to the **TaskDependencies** collection. The following example links "Preliminary Research &amp; Approval" as a predecessor to "Draft Initial Business Case" in a finish-to-start relationship.

## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/taskdependency-delete-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/taskdependency-application-property-project%28Office.15%29.aspx)|
|[From](http://msdn.microsoft.com/library/taskdependency-from-property-project%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/taskdependency-index-property-project%28Office.15%29.aspx)|
|[Lag](http://msdn.microsoft.com/library/taskdependency-lag-property-project%28Office.15%29.aspx)|
|[LagType](http://msdn.microsoft.com/library/taskdependency-lagtype-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/taskdependency-parent-property-project%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/taskdependency-path-property-project%28Office.15%29.aspx)|
|[To](http://msdn.microsoft.com/library/taskdependency-to-property-project%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/taskdependency-type-property-project%28Office.15%29.aspx)|

