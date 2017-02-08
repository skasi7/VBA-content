---
title: Subprojects Object (Project)
ms.prod: PROJECTSERVER
ms.assetid: 15688529-6d9c-6429-0d22-a5a16c033dcc
---


# Subprojects Object (Project)

Contains a collection of  **[Subproject](subproject-object-project.md)** objects


## Example

 **Using the Subprojects Collection Object**

Use  **Subprojects** ( _Index_ ), where _Index_ is the subproject index or project summary task name, to return a single **Subproject** object. The following example prevents changes made to the specified subproject in a master project from being automatically made to the source project.




```
ActiveProject.Subprojects("Arcadia Bay Online Catalog Plan").LinkToSource = False
```

 **Getting the Subprojects Collection object**

Use the  **[Subprojects](http://msdn.microsoft.com/library/project-subprojects-property-project%28Office.15%29.aspx)** property to return a **Subprojects** collection. The following example cautions the user if any of the subprojects in the active project are not on the hard disk.




```
Dim SubProj As Subproject 

 

For Each SubProj in ActiveProject.Subprojects 

 If UCase(Left$(SubProj.Path, 1)) <> "C" Then 

 MsgBox Right$(SubProj.Path, InStrRev(SubProj.Path, "\") - 1) &amp; _ 

 " is not on your local hard disk.", vbExclamation 

 End If 

Next SubProj
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/subprojects-application-property-project%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/subprojects-count-property-project%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/subprojects-item-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/subprojects-parent-property-project%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/project-object-model%28Office.15%29.aspx)
