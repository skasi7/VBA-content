---
title: AppointmentItem Object (Outlook)
keywords: vbaol11.chm2988
f1_keywords:
- vbaol11.chm2988
ms.prod: OUTLOOK
api_name:
- Outlook.AppointmentItem
ms.assetid: 204a409d-654e-27aa-643a-8344c631b82d
---


# AppointmentItem Object (Outlook)

Represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder.


## Remarks

Use the  **[CreateItem](http://msdn.microsoft.com/library/application-createitem-method-outlook%28Office.15%29.aspx)** method to create an **AppointmentItem** object that represents a new appointment.

Use  **[Items](http://msdn.microsoft.com/library/items-item-method-outlook%28Office.15%29.aspx)** ( _index_ ), where _index_ is the index number of an appointment or a value used to match the default property of an appointment, to return a single **AppointmentItem** object from a Calendar folder.

You can also return an  **AppointmentItem** object from a **[MeetingItem](meetingitem-object-outlook.md)** object by using the **[GetAssociatedAppointment](http://msdn.microsoft.com/library/meetingitem-getassociatedappointment-method-outlook%28Office.15%29.aspx)** method.

When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes. This practice applies to the recurring  **AppointmentItem** object, and any **[Exception](http://msdn.microsoft.com/library/exception-object-outlook%28Office.15%29.aspx)** or **[RecurrencePattern](recurrencepattern-object-outlook.md)** object. To release a reference in Visual Basic for Applications (VBA) or Visual Basic, set that existing object to **Nothing**. In C#, explicitly release the memory for that object.

Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference, held by another add-in or Outlook, to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.

The following code example in VBA shows how to release and refresh references in order to obtain up-to-date data for a recurring appointment. The example obtains a set of appointment items from the Calendar folder. It assumes that the first item in the appointment collection is part of a recurring appointment. The example shows that a reference to the appointment collection obtained before an exception is created does not reflect the exception. The example then releases this reference and other existing appointment references, after which new references that point to the appointment collection reflect the exception.




```
Sub TestExceptions() 
 
 Dim oItems As Items 
 
 Dim oItemOriginal As AppointmentItem 
 
 Dim oItemNew As AppointmentItem 
 
 Dim rPattern As RecurrencePattern 
 
 Dim oEx As Exceptions 
 
 Dim oEx2 As Exceptions 
 
 Dim oOccurrence As AppointmentItem 
 
 Dim i As Long 
 
 
 
 ' This is the initial reference to an appointment collection. 
 
 Set oItems = _ 
 
 Outlook.Application.Session.GetDefaultFolder(olFolderCalendar).Items 
 
 
 
 ' This is the original reference to the first appointment in the 
 
 ' collection before an exception is created. 
 
 Set oItemOriginal = oItems.Item(1) 
 
 
 
 ' Code example assumes that the first appointment in the collection 
 
 ' is a recurring appointment. 
 
 Set oOccurrence = _ 
 
 oItemOriginal.GetRecurrencePattern().GetOccurrence(#2/28/2010 8:00:00 AM#) 
 
 
 
 ' Create an exception by changing the 2/28 occurrence to 3/3. 
 
 oOccurrence.Start = #3/3/2010 8:00:00 AM# 
 
 oOccurrence.Save 
 
 
 
 Stop 
 
 
 
 ' Preexisting reference to the first appointment in the collection 
 
 ' does not reflect the exception. 
 
 oItemOriginal.Save 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print oItemOriginal.subject 
 
 Debug.Print " Original item exceptions: " &amp; oEx.Count 
 
 
 
 ' Get a new reference based on the existing reference to the 
 
 ' appointment collection created before the exception. 
 
 ' The new reference does not reflect the exception. 
 
 Set oItemNew = oItems.Item(1) 
 
 oItemNew.Save 
 
 Set oEx2 = oItemNew.GetRecurrencePattern().Exceptions 
 
 Debug.Print " New item exceptions: " &amp; oEx2.Count 
 
 
 
 ' Same: preexisting reference to the first appointment in the collection 
 
 ' does not reflect the exception. 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print " Original item exceptions: " &amp; oEx.Count 
 
 
 
 ' Release all existing references to appointment items, 
 
 ' including the appointment collection, an exception, occurrence, 
 
 ' or any other appointment. 
 
 Debug.Print "REFRESH ITEM COLLECTION" 
 
 Set oItems = Nothing 
 
 Set oItemNew = Nothing 
 
 Set oEx = Nothing 
 
 Set oEx2 = Nothing 
 
 Set oOccurrence = Nothing 
 
 Set oItemOriginal = Nothing 
 
 Set rPattern = Nothing 
 
 
 
 ' Get new references to appointment items, including the appointment 
 
 ' collection, individual appointments, and exceptions. 
 
 Set oItems = _ 
 
 Outlook.Application.Session.GetDefaultFolder(olFolderCalendar).Items 
 
 Set oItemNew = oItems.Item(1) 
 
 
 
 ' If no other add-ins have the same recurring appointment open, 
 
 ' the new references reflect the current exception count. 
 
 Set oEx2 = oItemNew.GetRecurrencePattern().Exceptions 
 
 Debug.Print " New item exceptions: " &amp; oEx2.Count 
 
 
 
 Debug.Print "RE-GET ORIGINAL" 
 
 Set oItemOriginal = oItems.Item(1) 
 
 Set oEx = oItemOriginal.GetRecurrencePattern().Exceptions 
 
 Debug.Print " Original item exceptions: " &amp; oEx.Count 
 
End Sub
```


## Example

The following Visual Basic for Applications (VBA) example returns a new appointment.


```
Set myItem = Application.CreateItem(olAppointmentItem)
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/appointmentitem-afterwrite-event-outlook%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/appointmentitem-attachmentadd-event-outlook%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/appointmentitem-attachmentread-event-outlook%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/appointmentitem-attachmentremove-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/appointmentitem-beforeattachmentadd-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/appointmentitem-beforeattachmentpreview-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/appointmentitem-beforeattachmentread-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/appointmentitem-beforeattachmentsave-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/appointmentitem-beforeattachmentwritetotempfile-event-outlook%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/appointmentitem-beforeautosave-event-outlook%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/appointmentitem-beforechecknames-event-outlook%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/appointmentitem-beforedelete-event-outlook%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/appointmentitem-beforeread-event-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/appointmentitem-close-event-outlook%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/appointmentitem-customaction-event-outlook%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/appointmentitem-custompropertychange-event-outlook%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/appointmentitem-forward-event-outlook%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/appointmentitem-open-event-outlook%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/appointmentitem-propertychange-event-outlook%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/appointmentitem-read-event-outlook%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/appointmentitem-readcomplete-event-outlook%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/appointmentitem-reply-event-outlook%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/appointmentitem-replyall-event-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/appointmentitem-send-event-outlook%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/appointmentitem-unload-event-outlook%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/appointmentitem-write-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[ClearRecurrencePattern](http://msdn.microsoft.com/library/appointmentitem-clearrecurrencepattern-method-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/appointmentitem-close-method-outlook%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/appointmentitem-copy-method-outlook%28Office.15%29.aspx)|
|[CopyTo](http://msdn.microsoft.com/library/appointmentitem-copyto-method-outlook%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/appointmentitem-delete-method-outlook%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/appointmentitem-display-method-outlook%28Office.15%29.aspx)|
|[ForwardAsVcal](http://msdn.microsoft.com/library/appointmentitem-forwardasvcal-method-outlook%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/appointmentitem-getconversation-method-outlook%28Office.15%29.aspx)|
|[GetOrganizer](http://msdn.microsoft.com/library/appointmentitem-getorganizer-method-outlook%28Office.15%29.aspx)|
|[GetRecurrencePattern](http://msdn.microsoft.com/library/appointmentitem-getrecurrencepattern-method-outlook%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/appointmentitem-move-method-outlook%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/appointmentitem-printout-method-outlook%28Office.15%29.aspx)|
|[Respond](http://msdn.microsoft.com/library/appointmentitem-respond-method-outlook%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/appointmentitem-save-method-outlook%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/appointmentitem-saveas-method-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/appointmentitem-send-method-outlook%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/appointmentitem-showcategoriesdialog-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/appointmentitem-actions-property-outlook%28Office.15%29.aspx)|
|[AllDayEvent](http://msdn.microsoft.com/library/appointmentitem-alldayevent-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/appointmentitem-application-property-outlook%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/appointmentitem-attachments-property-outlook%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/appointmentitem-autoresolvedwinner-property-outlook%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/appointmentitem-billinginformation-property-outlook%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/appointmentitem-body-property-outlook%28Office.15%29.aspx)|
|[BusyStatus](http://msdn.microsoft.com/library/appointmentitem-busystatus-property-outlook%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/appointmentitem-categories-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/appointmentitem-class-property-outlook%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/appointmentitem-companies-property-outlook%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/appointmentitem-conflicts-property-outlook%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/appointmentitem-conversationid-property-outlook%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/appointmentitem-conversationindex-property-outlook%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/appointmentitem-conversationtopic-property-outlook%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/appointmentitem-creationtime-property-outlook%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/appointmentitem-downloadstate-property-outlook%28Office.15%29.aspx)|
|[Duration](http://msdn.microsoft.com/library/appointmentitem-duration-property-outlook%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/appointmentitem-end-property-outlook%28Office.15%29.aspx)|
|[EndInEndTimeZone](http://msdn.microsoft.com/library/appointmentitem-endinendtimezone-property-outlook%28Office.15%29.aspx)|
|[EndTimeZone](http://msdn.microsoft.com/library/appointmentitem-endtimezone-property-outlook%28Office.15%29.aspx)|
|[EndUTC](http://msdn.microsoft.com/library/appointmentitem-endutc-property-outlook%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/appointmentitem-entryid-property-outlook%28Office.15%29.aspx)|
|[ForceUpdateToAllAttendees](http://msdn.microsoft.com/library/appointmentitem-forceupdatetoallattendees-property-outlook%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/appointmentitem-formdescription-property-outlook%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/appointmentitem-getinspector-property-outlook%28Office.15%29.aspx)|
|[GlobalAppointmentID](http://msdn.microsoft.com/library/appointmentitem-globalappointmentid-property-outlook%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/appointmentitem-importance-property-outlook%28Office.15%29.aspx)|
|[InternetCodepage](http://msdn.microsoft.com/library/appointmentitem-internetcodepage-property-outlook%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/appointmentitem-isconflict-property-outlook%28Office.15%29.aspx)|
|[IsRecurring](http://msdn.microsoft.com/library/appointmentitem-isrecurring-property-outlook%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/appointmentitem-itemproperties-property-outlook%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/appointmentitem-lastmodificationtime-property-outlook%28Office.15%29.aspx)|
|[Location](http://msdn.microsoft.com/library/appointmentitem-location-property-outlook%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/appointmentitem-markfordownload-property-outlook%28Office.15%29.aspx)|
|[MeetingStatus](http://msdn.microsoft.com/library/appointmentitem-meetingstatus-property-outlook%28Office.15%29.aspx)|
|[MeetingWorkspaceURL](http://msdn.microsoft.com/library/appointmentitem-meetingworkspaceurl-property-outlook%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/appointmentitem-messageclass-property-outlook%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/appointmentitem-mileage-property-outlook%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/appointmentitem-noaging-property-outlook%28Office.15%29.aspx)|
|[OptionalAttendees](http://msdn.microsoft.com/library/appointmentitem-optionalattendees-property-outlook%28Office.15%29.aspx)|
|[Organizer](http://msdn.microsoft.com/library/appointmentitem-organizer-property-outlook%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/appointmentitem-outlookinternalversion-property-outlook%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/appointmentitem-outlookversion-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/appointmentitem-parent-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/appointmentitem-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[Recipients](http://msdn.microsoft.com/library/appointmentitem-recipients-property-outlook%28Office.15%29.aspx)|
|[RecurrenceState](http://msdn.microsoft.com/library/appointmentitem-recurrencestate-property-outlook%28Office.15%29.aspx)|
|[ReminderMinutesBeforeStart](http://msdn.microsoft.com/library/appointmentitem-reminderminutesbeforestart-property-outlook%28Office.15%29.aspx)|
|[ReminderOverrideDefault](http://msdn.microsoft.com/library/appointmentitem-reminderoverridedefault-property-outlook%28Office.15%29.aspx)|
|[ReminderPlaySound](http://msdn.microsoft.com/library/appointmentitem-reminderplaysound-property-outlook%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/appointmentitem-reminderset-property-outlook%28Office.15%29.aspx)|
|[ReminderSoundFile](http://msdn.microsoft.com/library/appointmentitem-remindersoundfile-property-outlook%28Office.15%29.aspx)|
|[ReplyTime](http://msdn.microsoft.com/library/appointmentitem-replytime-property-outlook%28Office.15%29.aspx)|
|[RequiredAttendees](http://msdn.microsoft.com/library/appointmentitem-requiredattendees-property-outlook%28Office.15%29.aspx)|
|[Resources](http://msdn.microsoft.com/library/appointmentitem-resources-property-outlook%28Office.15%29.aspx)|
|[ResponseRequested](http://msdn.microsoft.com/library/appointmentitem-responserequested-property-outlook%28Office.15%29.aspx)|
|[ResponseStatus](http://msdn.microsoft.com/library/appointmentitem-responsestatus-property-outlook%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/appointmentitem-rtfbody-property-outlook%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/appointmentitem-saved-property-outlook%28Office.15%29.aspx)|
|[SendUsingAccount](http://msdn.microsoft.com/library/appointmentitem-sendusingaccount-property-outlook%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/appointmentitem-sensitivity-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/appointmentitem-session-property-outlook%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/appointmentitem-size-property-outlook%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/appointmentitem-start-property-outlook%28Office.15%29.aspx)|
|[StartInStartTimeZone](http://msdn.microsoft.com/library/appointmentitem-startinstarttimezone-property-outlook%28Office.15%29.aspx)|
|[StartTimeZone](http://msdn.microsoft.com/library/appointmentitem-starttimezone-property-outlook%28Office.15%29.aspx)|
|[StartUTC](http://msdn.microsoft.com/library/appointmentitem-startutc-property-outlook%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/appointmentitem-subject-property-outlook%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/appointmentitem-unread-property-outlook%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/appointmentitem-userproperties-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[How to: Import Appointment XML Data into Outlook Appointment Objects](http://msdn.microsoft.com/library/import-appointment-xml-data-into-outlook-appointment-objects-outlook%28Office.15%29.aspx)
[AppointmentItem Object Members](http://msdn.microsoft.com/library/appointmentitem-members-outlook%28Office.15%29.aspx)
