---
title: AppointmentItem Properties (Outlook)
ms.prod: OUTLOOK
ms.assetid: 4a18f5fc-7f3a-492d-8b9c-1f280b338a89
---


# AppointmentItem Properties (Outlook)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](appointmentitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[AllDayEvent](appointmentitem-alldayevent-property-outlook.md)|Returns  **True** if the appointment is an all-day event (as opposed to a specified time). Read/write.|
|[Application](appointmentitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](appointmentitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoResolvedWinner](appointmentitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](appointmentitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](appointmentitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[BusyStatus](appointmentitem-busystatus-property-outlook.md)|Returns or sets an  **[OlBusyStatus](olbusystatus-enumeration-outlook.md)** constant indicating the busy status of the user for the appointment. Read/write.|
|[Categories](appointmentitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](appointmentitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](appointmentitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](appointmentitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ConversationID](appointmentitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[AppointmentItem](appointmentitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](appointmentitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](appointmentitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DownloadState](appointmentitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[Duration](appointmentitem-duration-property-outlook.md)|Returns or sets a  **Long** indicating the duration (in minutes) of the **[AppointmentItem](appointmentitem-object-outlook.md)** . Read/write.|
|[End](appointmentitem-end-property-outlook.md)|Returns or sets a  **Date** indicating the end date and time of an **[AppointmentItem](appointmentitem-object-outlook.md)** . Read/write.|
|[EndInEndTimeZone](appointmentitem-endinendtimezone-property-outlook.md)|Returns or sets a  **Date** value that represents the end date and time of the appointment expressed in the **[AppointmentItem.EndTimeZone](appointmentitem-endtimezone-property-outlook.md)** . Read/write.|
|[EndTimeZone](appointmentitem-endtimezone-property-outlook.md)|Returns or sets a  **[TimeZone](timezone-object-outlook.md)** value that corresponds to the end time of the appointment. Read/write.|
|[EndUTC](appointmentitem-endutc-property-outlook.md)|Returns or sets a  **Date** value that represents the end date and time of the appointment expressed in the Coordinated Univeral Time (UTC) standard. Read/write.|
|[EntryID](appointmentitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[ForceUpdateToAllAttendees](appointmentitem-forceupdatetoallattendees-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether updates to the[AppointmentItem](appointmentitem-object-outlook.md) object should be sent to all attendees. Read/write.|
|[FormDescription](appointmentitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](appointmentitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[GlobalAppointmentID](appointmentitem-globalappointmentid-property-outlook.md)|Returns a  **String** value that represents a unique global identifier for the **[AppointmentItem](appointmentitem-object-outlook.md)** object. Read-only.|
|[Importance](appointmentitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[InternetCodepage](appointmentitem-internetcodepage-property-outlook.md)|Returns or sets a  **Long** that determines the Internet code page used by the item. Read/write.|
|[IsConflict](appointmentitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item on the local computer is different from the copy on the server. Read-only.|
|[IsRecurring](appointmentitem-isrecurring-property-outlook.md)|Returns a  **Boolean** value that is **True** if the appointment is a recurring appointment. Read-only.|
|[ItemProperties](appointmentitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](appointmentitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[Location](appointmentitem-location-property-outlook.md)|Returns or sets a  **String** representing the specific office location (for example, Building 1 Room 1 or Suite 123) for the appointment. Read/write.|
|[MarkForDownload](appointmentitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MeetingStatus](appointmentitem-meetingstatus-property-outlook.md)|Returns or sets an  **[OlMeetingStatus](olmeetingstatus-enumeration-outlook.md)** constant specifying the meeting status of the appointment. Read/write.|
|[MeetingWorkspaceURL](appointmentitem-meetingworkspaceurl-property-outlook.md)|Returns a  **String** value that represents the URL for the Meeting Workspace that the appointment item is linked to. Read-only.|
|[MessageClass](appointmentitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](appointmentitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](appointmentitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OptionalAttendees](appointmentitem-optionalattendees-property-outlook.md)|Returns or sets a  **String** representing the display string of optional attendees names for the appointment. Read/write.|
|[Organizer](appointmentitem-organizer-property-outlook.md)|Returns a  **String** representing the name of the organizer of the appointment. Read-only.|
|[OutlookInternalVersion](appointmentitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](appointmentitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](appointmentitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](appointmentitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[AppointmentItem](appointmentitem-object-outlook.md)** object. Read-only.|
|[Recipients](appointmentitem-recipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the recipients for the Outlook item. Read-only.|
|[RecurrenceState](appointmentitem-recurrencestate-property-outlook.md)|Returns an  **[OlRecurrenceState](olrecurrencestate-enumeration-outlook.md)** constant indicating the recurrence property of the specified object. Read-only.|
|[ReminderMinutesBeforeStart](appointmentitem-reminderminutesbeforestart-property-outlook.md)|Returns or sets a  **Long** indicating the number of minutes the reminder should occur prior to the start of the appointment. Read/write.|
|[ReminderOverrideDefault](appointmentitem-reminderoverridedefault-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder overrides the default reminder behavior for the item. Read/write.|
|[ReminderPlaySound](appointmentitem-reminderplaysound-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder should play a sound when it occurs for this item. Read/write.|
|[ReminderSet](appointmentitem-reminderset-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a reminder has been set for this item. Read/write.|
|[ReminderSoundFile](appointmentitem-remindersoundfile-property-outlook.md)|Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.|
|[ReplyTime](appointmentitem-replytime-property-outlook.md)|Returns or sets a  **Date** indicating the reply time for the appointment. Read/write.|
|[RequiredAttendees](appointmentitem-requiredattendees-property-outlook.md)|Returns a semicolon-delimited  **String** of required attendee names for the meeting appointment. Read/write.|
|[Resources](appointmentitem-resources-property-outlook.md)|Returns a semicolon-delimited  **String** of resource names for the meeting. Read/write.|
|[ResponseRequested](appointmentitem-responserequested-property-outlook.md)|Returns a  **Boolean** that indicates **True** if the sender would like a response to the meeting request for the appointment. Read/write.|
|[ResponseStatus](appointmentitem-responsestatus-property-outlook.md)|Returns an  **[OlResponseStatus](olresponsestatus-enumeration-outlook.md)** constant indicating the overall status of the meeting for the current user for the appointment. Read-only.|
|[RTFBody](appointmentitem-rtfbody-property-outlook.md)|Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.|
|[Saved](appointmentitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[SendUsingAccount](appointmentitem-sendusingaccount-property-outlook.md)|Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account under which the **[AppointmentItem](appointmentitem-object-outlook.md)** is to be sent. Read/write.|
|[Sensitivity](appointmentitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Session](appointmentitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](appointmentitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Start](appointmentitem-start-property-outlook.md)|Returns or sets a  **Date** indicating the starting date and time for the Outlook item. Read/write.|
|[StartInStartTimeZone](appointmentitem-startinstarttimezone-property-outlook.md)|Returns or sets a  **Date** value that represents the start date and time of the appointment expressed in the **[AppointmentItem.StartTimeZone](appointmentitem-starttimezone-property-outlook.md)** . Read/write.|
|[StartTimeZone](appointmentitem-starttimezone-property-outlook.md)|Returns or sets a  **[TimeZone](timezone-object-outlook.md)** value that corresponds to the time zone for the start time of the appointment. Read/write.|
|[StartUTC](appointmentitem-startutc-property-outlook.md)|Returns or sets a  **Date** value that represents the start date and time of the appointment expressed in the Coordinated Univeral Time (UTC) standard. Read/write.|
|[Subject](appointmentitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[UnRead](appointmentitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](appointmentitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

