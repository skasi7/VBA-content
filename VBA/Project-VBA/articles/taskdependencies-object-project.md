---
title: TaskDependencies Object (Project)
ms.prod: PROJECTSERVER
ms.assetid: 60bda111-998f-1cc2-0b18-b419041767f5
---


# TaskDependencies Object (Project)

Contains a collection of  **[TaskDependency](taskdependency-object-project.md)** objects.


## Example

 **Using the TaskDependency Object**

Use  **TaskDependencies** ( _Index_ ), where _Index_ is the dependency index, to return a single **TaskDependency** object. The following example adds 1.5 days of lag to the link between the specified task and the predecessor specified in its first task dependency.




```
ActiveProject.Tasks("Draft Initial Business Case").TaskDependencies(1).Lag = "1.5d"
```

 **Using the TaskDependencies Collection**

Use the  **[TaskDependencies](http://msdn.microsoft.com/library/task-taskdependencies-property-project%28Office.15%29.aspx)** property to return a **TaskDependencies** collection. The following example examines each predecessor for the specified task and displays a message for each that has a priority of "High" or better.




```
Dim TaskDep As TaskDependency 

 

For Each TaskDep In ActiveProject.Tasks("Write Requirements Brief").TaskDependencies 

 If TaskDep.From.Priority > 500 Then 

 MsgBox "Task #" &amp; TaskDep.From.ID &amp; " (" &amp; TaskDep.From.Name &amp; ") " &amp; _ 

 "has a priority higher than medium." 

 End If 

Next TaskDep
```

Use the  **[Add](http://msdn.microsoft.com/library/taskdependencies-add-method-project%28Office.15%29.aspx)** method to add a **TaskDependency** object to the **TaskDependencies** collection. The following example links "Preliminary Research &amp; Approval" as a predecessor to "Draft Initial Business Case" in a finish-to-start relationship.




```
ActiveProject.Tasks("Draft Initial Business Case").TaskDependencies.Add ActiveProject.Tasks("Preliminary Research &amp; Approval"), pjFinishToStart
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/taskdependencies-add-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/taskdependencies-application-property-project%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/taskdependencies-count-property-project%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/taskdependencies-item-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/taskdependencies-parent-property-project%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/project-object-model%28Office.15%29.aspx)
