---
title: Calendar Object (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.Calendar
ms.assetid: 2d3b0f05-4762-0058-15d4-47e1d2b9d9a9
---


# Calendar Object (Project)



Represents the calendar for a resource or project. The  **Calendar** object is a member of the **[Calendars](calendars-object-project.md)** collection.
 **Using the Calendar Object**
Use  **BaseCalendars(** _Index_ **)**, where _Index_ is the calendar index number or calendar name, to return a single **Calendar** object.
 **Using the Calendars Collection**
Use the  **[BaseCalendars](http://msdn.microsoft.com/library/project-basecalendars-property-project%28Office.15%29.aspx)** property to return a **Calendars** collection. The following example resets the properties of each base calendar in the active project to their default values.
Use the  **[BaseCalendarCreate](http://msdn.microsoft.com/library/application-basecalendarcreate-method-project%28Office.15%29.aspx)** method to add a **Calendar** object to the **Calendars** collection. The following example creates a new base calendar.

## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/calendar-delete-method-project%28Office.15%29.aspx)|
|[Period](http://msdn.microsoft.com/library/calendar-period-method-project%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/calendar-reset-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/calendar-application-property-project%28Office.15%29.aspx)|
|[BaseCalendar](http://msdn.microsoft.com/library/calendar-basecalendar-property-project%28Office.15%29.aspx)|
|[Enterprise](http://msdn.microsoft.com/library/calendar-enterprise-property-project%28Office.15%29.aspx)|
|[Exceptions](http://msdn.microsoft.com/library/calendar-exceptions-property-project%28Office.15%29.aspx)|
|[Guid](http://msdn.microsoft.com/library/calendar-guid-property-project%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/calendar-index-property-project%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/calendar-name-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/calendar-parent-property-project%28Office.15%29.aspx)|
|[ResourceGuid](http://msdn.microsoft.com/library/calendar-resourceguid-property-project%28Office.15%29.aspx)|
|[WeekDays](http://msdn.microsoft.com/library/calendar-weekdays-property-project%28Office.15%29.aspx)|
|[WorkWeeks](http://msdn.microsoft.com/library/calendar-workweeks-property-project%28Office.15%29.aspx)|
|[Years](http://msdn.microsoft.com/library/calendar-years-property-project%28Office.15%29.aspx)|

