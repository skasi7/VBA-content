---
title: CalendarModule Object (Outlook)
keywords: vbaol11.chm3194
f1_keywords:
- vbaol11.chm3194
ms.prod: OUTLOOK
api_name:
- Outlook.CalendarModule
ms.assetid: 9203024d-9cef-75e0-600f-f3899e24761a
---


# CalendarModule Object (Outlook)

Represents the  **Calendar** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **CalendarModule** object, derived from the **[NavigationModule](http://msdn.microsoft.com/library/navigationmodule-object-outlook%28Office.15%29.aspx)** object, provides access to the navigation groups contained in the **Calendar** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](http://msdn.microsoft.com/library/navigationmodules-getnavigationmodule-method-outlook%28Office.15%29.aspx)** method or the **[Item](http://msdn.microsoft.com/library/navigationmodules-item-method-outlook%28Office.15%29.aspx)** method of the **[Modules](http://msdn.microsoft.com/library/navigationpane-modules-property-outlook%28Office.15%29.aspx)** collection for the parent **[NavigationPane](http://msdn.microsoft.com/library/navigationpane-object-outlook%28Office.15%29.aspx)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](http://msdn.microsoft.com/library/navigationmodule-navigationmoduletype-property-outlook%28Office.15%29.aspx)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleCalendar**, you can then cast the **NavigationModule** object reference as a **CalendarModule** object to access the **[NavigationGroups](http://msdn.microsoft.com/library/calendarmodule-navigationgroups-property-outlook%28Office.15%29.aspx)** property for that navigation module.

You can use the  **[Visible](http://msdn.microsoft.com/library/calendarmodule-visible-property-outlook%28Office.15%29.aspx)** property to determine if the navigation module is visible and the **[Position](http://msdn.microsoft.com/library/calendarmodule-position-property-outlook%28Office.15%29.aspx)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](http://msdn.microsoft.com/library/calendarmodule-name-property-outlook%28Office.15%29.aspx)** property to return the display name of the **Calendar** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/calendarmodule-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/calendarmodule-class-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/calendarmodule-name-property-outlook%28Office.15%29.aspx)|
|[NavigationGroups](http://msdn.microsoft.com/library/calendarmodule-navigationgroups-property-outlook%28Office.15%29.aspx)|
|[NavigationModuleType](http://msdn.microsoft.com/library/calendarmodule-navigationmoduletype-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/calendarmodule-parent-property-outlook%28Office.15%29.aspx)|
|[Position](http://msdn.microsoft.com/library/calendarmodule-position-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/calendarmodule-session-property-outlook%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/calendarmodule-visible-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[CalendarModule Object Members](http://msdn.microsoft.com/library/calendarmodule-members-outlook%28Office.15%29.aspx)
