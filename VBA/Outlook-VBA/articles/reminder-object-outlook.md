---
title: Reminder Object (Outlook)
keywords: vbaol11.chm3014
f1_keywords:
- vbaol11.chm3014
ms.prod: OUTLOOK
api_name:
- Outlook.Reminder
ms.assetid: b7364e48-51bc-b360-2154-e85e7779ece4
---


# Reminder Object (Outlook)

Represents an Outlook reminder.


## Remarks

Reminders allow users to keep track of upcoming appointments by scheduling a pop-up dialog box to appear at a given time. In addition to appointments, reminders can occur for tasks, contacts and e-mail messages.

Use  **[Reminders](http://msdn.microsoft.com/library/application-reminders-property-outlook%28Office.15%29.aspx)** ( _index_ ), where _index_ is the name or index number of the reminder, to return a single **Reminder** object.

Reminders are created programmatically when a new Microsoft Outlook item, such as an  **[AppointmentItem](appointmentitem-object-outlook.md)** object, is created and the item 's **[ReminderSet](http://msdn.microsoft.com/library/appointmentitem-reminderset-property-outlook%28Office.15%29.aspx)** property is set to **True**.

Use the  **Reminders** collection's **[Remove](http://msdn.microsoft.com/library/reminders-remove-method-outlook%28Office.15%29.aspx)** method to remove a **Reminder** object from the collection. Once a reminder is removed from its associated item, the **AppointmentItem** object's **ReminderSet** property is set to **False**.


## Example

The following example displays the caption of the first reminder in the collection.


```
Sub ViewReminderInfo() 
 
 'Displays information about first reminder in collection 
 
 
 
 Dim colReminders As Outlook.Reminders 
 
 Dim objRem As Reminder 
 
 
 
 Set colReminders = Application.Reminders 
 
 'If there are reminders, display message 
 
 If colReminders.Count <> 0 Then 
 
 Set objRem = colReminders.Item(1) 
 
 MsgBox "The caption of the first reminder in the collection is: " &amp; _ 
 
 objRem.Caption 
 
 Else 
 
 MsgBox "There are no reminders in the collection." 
 
 
 
 End If 
 
 
 
End Sub
```

The following example creates a new appointment item and sets the  **ReminderSet** property to **True**, adding a new **Reminder** object to the **Reminders** collection.




```
Sub AddAppt() 
 
 'Adds a new appointment and reminder to the reminders collection 
 
 Dim objApt As AppointmentItem 
 
 
 
 Set objApt = Application.CreateItem(olAppointmentItem) 
 
 objApt.ReminderSet = True 
 
 objApt.Subject = "Tuesday's meeting" 
 
 objApt.Save 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Dismiss](http://msdn.microsoft.com/library/reminder-dismiss-method-outlook%28Office.15%29.aspx)|
|[Snooze](http://msdn.microsoft.com/library/reminder-snooze-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/reminder-application-property-outlook%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/reminder-caption-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/reminder-class-property-outlook%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/reminder-isvisible-property-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/reminder-item-property-outlook%28Office.15%29.aspx)|
|[NextReminderDate](http://msdn.microsoft.com/library/reminder-nextreminderdate-property-outlook%28Office.15%29.aspx)|
|[OriginalReminderDate](http://msdn.microsoft.com/library/reminder-originalreminderdate-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/reminder-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/reminder-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Reminder Object Members](http://msdn.microsoft.com/library/reminder-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
