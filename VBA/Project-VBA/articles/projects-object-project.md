---
title: Projects Object (Project)
keywords: vbapj.chm131311
f1_keywords:
- vbapj.chm131311
ms.prod: PROJECTSERVER
ms.assetid: 5a254428-f50d-e74f-dd31-5cdb260a4364
---


# Projects Object (Project)

Contains a collection of **[Project](project-object-project.md)** objects.


## Example

 **Using the Project Object**

Use  **Projects** (Index), where Index is the project index number or project name, to return a single **Project** object. The following example switches among all the open projects, memorizes the full name of each, and then displays the results.




```
Dim Temp As Long, Names As String 

 

For Temp = 1 To Projects.Count 

 Projects(Temp).Activate 

 Names = Names &amp; Projects(Temp).FullName &amp; vbCrLf 

Next Temp 

 

MsgBox Names
```

 **Using the Projects Collection**

Use the  **[Projects](http://msdn.microsoft.com/library/application-projects-property-project%28Office.15%29.aspx)** property to return a **Projects** collection. The following example counts the number of open projects.




```
Application.Projects.Count
```

Because the  **Projects** collection is a top-level object, the following example is functionally identical to the preceding one.




```
Projects.Count
```

Use the  **[Add](http://msdn.microsoft.com/library/projects-add-method-project%28Office.15%29.aspx)** method to add a **Project** object to the **Projects** collection. The following example creates a new project without prompting for project information.




```
Projects.Add False
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/projects-add-method-project%28Office.15%29.aspx)|
|[CanCheckOut](http://msdn.microsoft.com/library/projects-cancheckout-method-project%28Office.15%29.aspx)|
|[CheckOut](http://msdn.microsoft.com/library/projects-checkout-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/projects-application-property-project%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/projects-count-property-project%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/projects-item-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/projects-parent-property-project%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/project-object-model%28Office.15%29.aspx)
