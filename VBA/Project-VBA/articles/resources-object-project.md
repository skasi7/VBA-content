---
title: Resources Object (Project)
ms.prod: PROJECTSERVER
ms.assetid: 84f8357a-358b-f2ae-e164-65c0c5abd383
---


# Resources Object (Project)

Contains a collection of  **[Resource](resource-object-project.md)** objects.


## Example

 **Using the Resources Collection**

Use  **Resources** ( _Index_ ), where _Index_ is the resource index number or resource name, to return a single **Resource** object. The following example lists the names of all resources in the active project.




```
Dim R As Long, Names As String 

 

For R = 1 To ActiveProject.Resources.Count 

 Names = ActiveProject.Resources(R).Name &amp; ", " &amp; Names 

Next R 

 

Names = Left$(Names, Len(Names) - Len(ListSeparator &amp; " ")) 

MsgBox Names
```

 **Using the Resources Collection**

Use the  **[Resources](http://msdn.microsoft.com/library/project-resources-property-project%28Office.15%29.aspx)** property to return a **Resources** collection. The following example generates the same list as the previous example, but does so by setting an object reference to `ActiveProject.Resources` , and then using `R` where `ActiveProject.Resources` is used.




```
Dim R As Resources, Temp As Long, Names As String 

 

Set R = ActiveProject.Resources 

 

For Temp = 1 To R.Count 

 Names = R(Temp).Name &amp; ", " &amp; Names 

Next Temp 

 

Names = Left$(Names, Len(Names) - Len(ListSeparator &amp; " ")) 

MsgBox Names
```

Use the  **[Add](http://msdn.microsoft.com/library/resources-add-method-project%28Office.15%29.aspx)** method to add a **Resource** object to the **Resources** collection. The following example adds a new resource named Matilda to the active project.




```
ActiveProject.Resources.Add "Matilda"
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/resources-add-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/resources-application-property-project%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/resources-count-property-project%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/resources-item-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/resources-parent-property-project%28Office.15%29.aspx)|
|[UniqueID](http://msdn.microsoft.com/library/resources-uniqueid-property-project%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/project-object-model%28Office.15%29.aspx)
