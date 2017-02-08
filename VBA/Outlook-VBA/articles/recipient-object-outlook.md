---
title: Recipient Object (Outlook)
keywords: vbaol11.chm2339
f1_keywords:
- vbaol11.chm2339
ms.prod: OUTLOOK
api_name:
- Outlook.Recipient
ms.assetid: 8cee4d79-ec55-52a4-710b-6456944ca86d
---


# Recipient Object (Outlook)

Represents a user or resource in Outlook, generally a mail or mobile message addressee.


## Remarks

Use the  **[Recipients](http://msdn.microsoft.com/library/recipients-item-method-outlook%28Office.15%29.aspx)** ( _index_ ) method, where _index_ is the name or index number, to return a single **Recipient** object. The name can be a string that represents the display name, the alias, the full SMTP e-mail address, or the mobile phone number of the recipient. A good practice is to use the SMTP e-mail address for a mail message, and the mobile phone number for a mobile message.

Use the  **[Add](http://msdn.microsoft.com/library/recipients-add-method-outlook%28Office.15%29.aspx)** method to create a new **Recipient** object and add it to the **[Recipients](recipients-object-outlook.md)** object. The **[Type](http://msdn.microsoft.com/library/recipient-type-property-outlook%28Office.15%29.aspx)** property of a new **Recipient** object is set to the default value for the associated **[AppointmentItem](appointmentitem-object-outlook.md)**, **[JournalItem](http://msdn.microsoft.com/library/journalitem-object-outlook%28Office.15%29.aspx)**, **[MailItem](http://msdn.microsoft.com/library/mailitem-object-outlook%28Office.15%29.aspx)**, **[MeetingItem](meetingitem-object-outlook.md)**, or **[TaskItem](taskitem-object-outlook.md)** object and must be reset to indicate another recipient type.


## Example



The following Visual Basic for Applications (VBA) example creates a new  **MailItem** object and adds Jon Grande as the recipient by using the default type ("To").




```
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande")
```

The following VBA example creates the same  **MailItem** object as the preceding example, and then changes the type of the **Recipient** object from the default (To) to CC.




```
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande") 
 
myRecipient.Type = olCC
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/recipient-delete-method-outlook%28Office.15%29.aspx)|
|[FreeBusy](http://msdn.microsoft.com/library/recipient-freebusy-method-outlook%28Office.15%29.aspx)|
|[Resolve](http://msdn.microsoft.com/library/recipient-resolve-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/recipient-address-property-outlook%28Office.15%29.aspx)|
|[AddressEntry](http://msdn.microsoft.com/library/recipient-addressentry-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/recipient-application-property-outlook%28Office.15%29.aspx)|
|[AutoResponse](http://msdn.microsoft.com/library/recipient-autoresponse-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/recipient-class-property-outlook%28Office.15%29.aspx)|
|[DisplayType](http://msdn.microsoft.com/library/recipient-displaytype-property-outlook%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/recipient-entryid-property-outlook%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/recipient-index-property-outlook%28Office.15%29.aspx)|
|[MeetingResponseStatus](http://msdn.microsoft.com/library/recipient-meetingresponsestatus-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/recipient-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/recipient-parent-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/recipient-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[Resolved](http://msdn.microsoft.com/library/recipient-resolved-property-outlook%28Office.15%29.aspx)|
|[Sendable](http://msdn.microsoft.com/library/recipient-sendable-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/recipient-session-property-outlook%28Office.15%29.aspx)|
|[TrackingStatus](http://msdn.microsoft.com/library/recipient-trackingstatus-property-outlook%28Office.15%29.aspx)|
|[TrackingStatusTime](http://msdn.microsoft.com/library/recipient-trackingstatustime-property-outlook%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/recipient-type-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Recipient Object Members](http://msdn.microsoft.com/library/recipient-members-outlook%28Office.15%29.aspx)
[How to: Obtain the E-mail Address of a Recipient](http://msdn.microsoft.com/library/obtain-the-e-mail-address-of-a-recipient%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
