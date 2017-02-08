---
title: CalendarView Object (Outlook)
keywords: vbaol11.chm3208
f1_keywords:
- vbaol11.chm3208
ms.prod: OUTLOOK
api_name:
- Outlook.CalendarView
ms.assetid: 37e078b9-9fc6-5894-b043-06d7257666a8
---


# CalendarView Object (Outlook)

Represents a view that displays Outlook items in a calendar format.


## Remarks

The  **CalendarView** object, derived from the **[View](view-object-outlook.md)** object, allows you to create customizable views that allow you to display Outlook items within a calendar, in one of several different modes.

Outlook provides several built-in  **CalendarView** objects, and you can also create custom **CalendarView** objects. Use the **[Add](http://msdn.microsoft.com/library/views-add-method-outlook%28Office.15%29.aspx)** method of the **[Views](http://msdn.microsoft.com/library/views-object-outlook%28Office.15%29.aspx)** collection to add a new **CalendarView** to a **[Folder](folder-object-outlook.md)** object. Use the **[Standard](http://msdn.microsoft.com/library/timelineview-standard-property-outlook%28Office.15%29.aspx)** property to determine if an existing **CalendarView** object is built-in or custom.

The  **CalendarView** object supports several different view modes, depending on the desired layout and time period in which to display Outlook items. Use the **[CalendarViewMode](http://msdn.microsoft.com/library/calendarview-calendarviewmode-property-outlook%28Office.15%29.aspx)** property to set the view mode, the **[StartField](http://msdn.microsoft.com/library/calendarview-startfield-property-outlook%28Office.15%29.aspx)** property to specify the Outlook item property that contains the start date, and the **[EndField](http://msdn.microsoft.com/library/calendarview-endfield-property-outlook%28Office.15%29.aspx)** property to specify the Outlook item property that contains the end date for Outlook items to be displayed.

If you set the  **CalendarViewMode** property to any value other than **olCalendarViewMonth**, you can use the **[DayWeekFont](http://msdn.microsoft.com/library/ddb6f65d-72e2-d3f2-b10f-b3d8bc4d21b3%28Office.15%29.aspx)** and **[DayWeekTimeFont](http://msdn.microsoft.com/library/37ea6e1f-4148-3ab4-e0aa-48c49321ac91%28Office.15%29.aspx)** properties to configure the fonts used to display the day, date, and time labels in the view. Use the **[DayWeekTimeScale](http://msdn.microsoft.com/library/calendarview-dayweektimescale-property-outlook%28Office.15%29.aspx)** to configure the time scale used to display Outlook items within the view. If you set the **CalendarViewMode** to **olCalendarViewMultiDay**, you can use the **[DaysInMultiDayMode](http://msdn.microsoft.com/library/calendarview-daysinmultidaymode-property-outlook%28Office.15%29.aspx)** property to determine the number of days to display in the view.

If you set the  **CalendarViewMode** to **olCalendarViewMonth**, you can use the **[MonthFont](http://msdn.microsoft.com/library/b69d1690-d1a8-dbc0-3de4-86a8eb98a471%28Office.15%29.aspx)** property to configure the fonts used to display the month and day labels and the **[MonthShowEndTime](http://msdn.microsoft.com/library/calendarview-monthshowendtime-property-outlook%28Office.15%29.aspx)** to indicate whether the end time for is displayed in the view.

You can also configure how Outlook items appear within the  **CalendarView** object. Use the **[BoldSubjects](http://msdn.microsoft.com/library/calendarview-boldsubjects-property-outlook%28Office.15%29.aspx)** property to indicate whether subjects for Outlook items are displayed in bold and the **[BoldDatesWithItems](http://msdn.microsoft.com/library/calendarview-bolddateswithitems-property-outlook%28Office.15%29.aspx)** property to indicate whether dates in the Date Navigator that contain Outlook items are displayed in bold. Use the **[Filter](http://msdn.microsoft.com/library/calendarview-filter-property-outlook%28Office.15%29.aspx)** property to determine which Outlook items to display in the view.

The definition for each  **CalendarView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](http://msdn.microsoft.com/library/calendarview-xml-property-outlook%28Office.15%29.aspx)** property to work with the XML definition for the **CalendarView** object.

Use the  **[Apply](http://msdn.microsoft.com/library/calendarview-apply-method-outlook%28Office.15%29.aspx)** method to apply any changes made to the **CalendarView** object to the current view. Use the **[Save](http://msdn.microsoft.com/library/calendarview-save-method-outlook%28Office.15%29.aspx)** method to persist any changes made to the **CalendarView** object. Use the **[LockUserChanges](http://msdn.microsoft.com/library/calendarview-lockuserchanges-property-outlook%28Office.15%29.aspx)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **CalendarView** objects, but you cannot delete them. Use the **[Delete](http://msdn.microsoft.com/library/calendarview-delete-method-outlook%28Office.15%29.aspx)** method to delete a custom **CalendarView** object. Use the **[Reset](http://msdn.microsoft.com/library/calendarview-reset-method-outlook%28Office.15%29.aspx)** method to reset the properties of a built-in **CalendarView** object to their default values.


## Example

The following Visual Basic for Applications (VBA) example configures the current  **CalendarView** object to show a single day, using an 8-point Verdana font to display items and a 16-point Verdana font to display time values and the Tasks header within the view.


```
Sub ConfigureDayViewFonts() 
 Dim objView As CalendarView 
 
 ' Check if the current view is a calendar view. 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 olCalendarView Then 
 
 ' Obtain a CalendarView object reference for the 
 ' current calendar view. 
 Set objView = _ 
 Application.ActiveExplorer.CurrentView 
 
 With objView 
 ' Set the calendar view to show a 
 ' single day. 
 .CalendarViewMode = olCalendarViewDay 
 
 ' Set the DayWeekFont to 8-point Verdana. 
 .DayWeekFont.Name = "Verdana" 
 .DayWeekFont.Size = 8 
 
 ' Set the DayWeekTimeFont to 16-point Verdana. 
 .DayWeekTimeFont.Name = "Verdana" 
 .DayWeekTimeFont.Size = 16 
 
 ' Save the calendar view. 
 .Save 
 End With 
 End If 
End Sub 

```


## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[CalendarView Object Members](http://msdn.microsoft.com/library/calendarview-members-outlook%28Office.15%29.aspx)
