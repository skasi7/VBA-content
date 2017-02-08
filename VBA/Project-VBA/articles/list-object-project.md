---
title: List Object (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.List
ms.assetid: 3934c2e8-d810-6571-9a33-1d41edbab87a
---


# List Object (Project)

Represents a collection of strings or numbers that contain field identification numbers, field names, reports, resource filters, resource tables, resource views, task filters, task tables, task views, or views. (There is no collection for  **List** objects.) It can be accessed through the **List** properties of the appropriate objects.


## Example

 **Using the List Object**

Use a property such as the  **[ReportList](http://msdn.microsoft.com/library/project-reportlist-property-project%28Office.15%29.aspx)** property to return a **List** object. The following example displays a list of all the reports available in the active project.




```
Dim Items As Integer, ReportNames As String 
 
For Items = 1 To ActiveProject.ReportList.Count 
 ReportNames = ActiveProject.ReportList(Items) &amp; _ 
 ListSeparator &amp; " " &amp; ReportNames 
Next Items 
 
MsgBox Left$(ReportNames, Len(ReportNames) - Len(ListSeparator &amp; " "))
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/list-application-property-project%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/list-count-property-project%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/list-item-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/list-parent-property-project%28Office.15%29.aspx)|

