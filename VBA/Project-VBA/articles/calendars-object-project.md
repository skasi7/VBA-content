---
title: Calendars Object (Project)
ms.prod: PROJECTSERVER
ms.assetid: a96c7b96-f0ab-5ec3-3d16-facea61b8ee5
---


# Calendars Object (Project)

Contains a collection of  **[Calendar](calendar-object-project.md)** objects.


## Example

 **Using the Calendar Object**

Use  **BaseCalendars(** _Index_ **)**, where _Index_ is the calendar index number or calendar name, to return a single **Calendar** object.




```
MsgBox ActiveProject.BaseCalendars(1).Name
```

 **Using the Calendars Collection**

Use the  **[BaseCalendars](http://msdn.microsoft.com/library/project-basecalendars-property-project%28Office.15%29.aspx)** property to return a **Calendars** collection. The following example resets the properties of each base calendar in the active project to their default values.




```
Dim C As Calendar 

 

For Each C In ActiveProject.BaseCalendars 

 C.Reset 

Next C
```

Use the  **[BaseCalendarCreate](http://msdn.microsoft.com/library/application-basecalendarcreate-method-project%28Office.15%29.aspx)** method to add a **Calendar** object to the **Calendars** collection. The following example creates a new base calendar.




```
BaseCalendarCreate Name:="Base Holiday Calendar"
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/calendars-application-property-project%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/calendars-count-property-project%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/calendars-item-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/calendars-parent-property-project%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/project-object-model%28Office.15%29.aspx)
