---
title: RecurrencePattern Object (Outlook)
keywords: vbaol11.chm268
f1_keywords:
- vbaol11.chm268
ms.prod: OUTLOOK
api_name:
- Outlook.RecurrencePattern
ms.assetid: 36c098f7-59fb-879a-5173-ed0260d13fa4
---


# RecurrencePattern Object (Outlook)

Represents the pattern of incidence of recurring appointments and tasks for the associated  **[AppointmentItem](appointmentitem-object-outlook.md)** and **[TaskItem](taskitem-object-outlook.md)** object.


## Remarks

Use the  **GetRecurrencePattern** method to return the **RecurrencePattern** object associated with an **AppointmentItem** or **TaskItem** object.

Calling  **GetRecurrencePattern** or **ClearRecurrencePattern** has the side effect of setting the **IsRecurring** property of the item accordingly. This property can be used as required for efficient filtering of the **[Items](items-object-outlook.md)** object.

The type of recurrence pattern is indicated by the  **[RecurrenceType](http://msdn.microsoft.com/library/recurrencepattern-recurrencetype-property-outlook%28Office.15%29.aspx)** property. The **RecurrenceType** property is the first property you should set.

The following properties are valid for all recurrence patterns:  **[EndTime](http://msdn.microsoft.com/library/recurrencepattern-endtime-property-outlook%28Office.15%29.aspx)**, **[Occurrences](http://msdn.microsoft.com/library/recurrencepattern-occurrences-property-outlook%28Office.15%29.aspx)**, **StartDate**, **[StartTime](http://msdn.microsoft.com/library/recurrencepattern-starttime-property-outlook%28Office.15%29.aspx)**, or **Type**.

The following table shows the properties that are valid for the different recurrence types. An error occurs if the item is saved and the property is null or contains an invalid value. Monthly and yearly patterns are only valid for a single day. Weekly patterns are only valid as the  **Or** of the **[DayOfWeekMask](http://msdn.microsoft.com/library/recurrencepattern-dayofweekmask-property-outlook%28Office.15%29.aspx)**.



|**RecurrenceType**|**Properties**|**Examples**|
|:-----|:-----|:-----|
|**olRecursDaily**|**[Duration](http://msdn.microsoft.com/library/recurrencepattern-duration-property-outlook%28Office.15%29.aspx)**, **EndTime**, **[Interval](http://msdn.microsoft.com/library/recurrencepattern-interval-property-outlook%28Office.15%29.aspx)**, **[NoEndDate](http://msdn.microsoft.com/library/recurrencepattern-noenddate-property-outlook%28Office.15%29.aspx)**, **Occurrences**, **[PatternStartDate](http://msdn.microsoft.com/library/recurrencepattern-patternstartdate-property-outlook%28Office.15%29.aspx)**, **[PatternEndDate](http://msdn.microsoft.com/library/recurrencepattern-patternenddate-property-outlook%28Office.15%29.aspx)**, **StartTime**|A value N for  **Interval** is every N days.|
|**olRecursWeekly**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N weeks. An example of **DayofWeekMask** is every Tuesday, Wednesday, and Thursday.|
|**olRecursMonthly**|**[DayOfMonth](http://msdn.microsoft.com/library/recurrencepattern-dayofmonth-property-outlook%28Office.15%29.aspx)**, **Duration**, **EndTime**, **Interval**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N months. A value N for **DayofMonth** is every Nth day of the month.|
|**olRecursMonthNth**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **[Instance](http://msdn.microsoft.com/library/recurrencepattern-instance-property-outlook%28Office.15%29.aspx)**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N months. An example of value N for **Instance** is every Nth Tuesday. An example of **DayofWeekMask** is every Tuesday and Wednesday.|
|**olRecursYearly**|**DayOfMonth**, **Duration**, **EndTime**, **Interval**, **[MonthOfYear](http://msdn.microsoft.com/library/recurrencepattern-monthofyear-property-outlook%28Office.15%29.aspx)**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **DayofMonth** is the Nth day of the month. An example of **MonthOfYear** is February.|
|**olRecursYearNth**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **Instance**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|An example of value N for  **Instance** is the Nth Tuesday. An example of **DayofWeekMask** is Tuesday, Wednesday, and Thursday. An example of **MonthOfYear** is February.|
When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes. This practice applies to the recurring  **AppointmentItem** object, and any **[Exception](http://msdn.microsoft.com/library/exception-object-outlook%28Office.15%29.aspx)** or **[RecurrencePattern](recurrencepattern-object-outlook.md)** object. To release a reference in Visual Basic for Applications (VBA) or Visual Basic, set that existing object to **Nothing**. In C#, explicitly release the memory for that object. For a code example, see the topic for the **AppointmentItem** object.

Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference, held by another add-in or Outlook, to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.


## Methods



|**Name**|
|:-----|
|[GetOccurrence](http://msdn.microsoft.com/library/recurrencepattern-getoccurrence-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/recurrencepattern-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/recurrencepattern-class-property-outlook%28Office.15%29.aspx)|
|[DayOfMonth](http://msdn.microsoft.com/library/recurrencepattern-dayofmonth-property-outlook%28Office.15%29.aspx)|
|[DayOfWeekMask](http://msdn.microsoft.com/library/recurrencepattern-dayofweekmask-property-outlook%28Office.15%29.aspx)|
|[Duration](http://msdn.microsoft.com/library/recurrencepattern-duration-property-outlook%28Office.15%29.aspx)|
|[EndTime](http://msdn.microsoft.com/library/recurrencepattern-endtime-property-outlook%28Office.15%29.aspx)|
|[Exceptions](http://msdn.microsoft.com/library/recurrencepattern-exceptions-property-outlook%28Office.15%29.aspx)|
|[Instance](http://msdn.microsoft.com/library/recurrencepattern-instance-property-outlook%28Office.15%29.aspx)|
|[Interval](http://msdn.microsoft.com/library/recurrencepattern-interval-property-outlook%28Office.15%29.aspx)|
|[MonthOfYear](http://msdn.microsoft.com/library/recurrencepattern-monthofyear-property-outlook%28Office.15%29.aspx)|
|[NoEndDate](http://msdn.microsoft.com/library/recurrencepattern-noenddate-property-outlook%28Office.15%29.aspx)|
|[Occurrences](http://msdn.microsoft.com/library/recurrencepattern-occurrences-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/recurrencepattern-parent-property-outlook%28Office.15%29.aspx)|
|[PatternEndDate](http://msdn.microsoft.com/library/recurrencepattern-patternenddate-property-outlook%28Office.15%29.aspx)|
|[PatternStartDate](http://msdn.microsoft.com/library/recurrencepattern-patternstartdate-property-outlook%28Office.15%29.aspx)|
|[RecurrenceType](http://msdn.microsoft.com/library/recurrencepattern-recurrencetype-property-outlook%28Office.15%29.aspx)|
|[Regenerate](http://msdn.microsoft.com/library/recurrencepattern-regenerate-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/recurrencepattern-session-property-outlook%28Office.15%29.aspx)|
|[StartTime](http://msdn.microsoft.com/library/recurrencepattern-starttime-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[RecurrencePattern Object Members](http://msdn.microsoft.com/library/recurrencepattern-members-outlook%28Office.15%29.aspx)
