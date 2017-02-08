---
title: Recipients Object (Outlook)
keywords: vbaol11.chm225
f1_keywords:
- vbaol11.chm225
ms.prod: OUTLOOK
api_name:
- Outlook.Recipients
ms.assetid: 774f56b7-4de8-9584-60cd-4fbf361f4c85
---


# Recipients Object (Outlook)

Contains a collection of  **[Recipient](recipient-object-outlook.md)** objects for an Outlook item.


## Remarks

Use the  **Recipients** property to return the **Recipients** object of an **[AppointmentItem](appointmentitem-object-outlook.md)**, **[JournalItem](http://msdn.microsoft.com/library/journalitem-object-outlook%28Office.15%29.aspx)**, **[MailItem](http://msdn.microsoft.com/library/mailitem-object-outlook%28Office.15%29.aspx)**, **[MeetingItem](meetingitem-object-outlook.md)**, or **[TaskItem](taskitem-object-outlook.md)** object.

Use the  **[Add](http://msdn.microsoft.com/library/recipients-add-method-outlook%28Office.15%29.aspx)** method to create a new **Recipient** object and add it to the **Recipients** object. The **[Type](http://msdn.microsoft.com/library/recipient-type-property-outlook%28Office.15%29.aspx)** property of a new **Recipient** object is set to the default for the associated **AppointmentItem**, **JournalItem**, **MailItem**, or **TaskItem** object and must be reset to indicate another recipient type.

Use  **Recipients** ( _index_ ), where _index_ is the name or index number, to return a single **Recipient** object. The name can be a string representing the display name, the alias, or the full SMTP e-mail address of the recipient.


## Example

The following example creates a new  **MailItem** object and adds Jon Grande as the recipient using the default type ("To").


```
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande")
```

The following example creates the same  **MailItem** object as the preceding example, and then changes the type of the **Recipient** object from the default ("To") to CC.




```
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande") 
 
myRecipient.Type = olCC
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/recipients-add-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/recipients-item-method-outlook%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/recipients-remove-method-outlook%28Office.15%29.aspx)|
|[ResolveAll](http://msdn.microsoft.com/library/recipients-resolveall-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/recipients-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/recipients-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/recipients-count-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/recipients-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/recipients-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[Recipients Object Members](http://msdn.microsoft.com/library/recipients-members-outlook%28Office.15%29.aspx)
