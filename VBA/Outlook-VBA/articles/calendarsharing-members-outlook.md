---
title: CalendarSharing Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 1b2b6233-9816-e3f2-5924-694ce30cc8ef
---


# CalendarSharing Members (Outlook)
Represents a set of utilities for sharing calendar information.

Represents a set of utilities for sharing calendar information.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)|Forwards calendar information from the parent  **[Folder](folder-object-outlook.md)** of the **[CalendarSharing](calendarsharing-object-outlook.md)** object as the payload of a **[MailItem](mailitem-object-outlook.md)** .|
|[SaveAsICal](calendarsharing-saveasical-method-outlook.md)|Exports calendar information from the parent  **[Folder](folder-object-outlook.md)** of the **[CalendarSharing](calendarsharing-object-outlook.md)** object as an iCalendar calendar (.ics) file.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](calendarsharing-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[CalendarDetail](calendarsharing-calendardetail-property-outlook.md)|Returns or sets an  **[OlCalendarDetail](olcalendardetail-enumeration-outlook.md)** value indicating the level of detail for calendar items included in the iCalendar (.ics) file created by the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** or **[SaveAsICal](calendarsharing-saveasical-method-outlook.md)** methods of the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.|
|[Class](calendarsharing-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[EndDate](calendarsharing-enddate-property-outlook.md)|Returns or sets a  **Date** value that represents the inclusive end date of the range of calendar items to be shared by the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.|
|[Folder](calendarsharing-folder-property-outlook.md)|Returns the  **[Folder](folder-object-outlook.md)** containing the calendar items to be shared by the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read-only.|
|[IncludeAttachments](calendarsharing-includeattachments-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether attachments for calendar items should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** or **[SaveAsICal](calendarsharing-saveasical-method-outlook.md)** methods of the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.|
|[IncludePrivateDetails](calendarsharing-includeprivatedetails-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether private details for calendar items should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** or **[SaveAsICal](calendarsharing-saveasical-method-outlook.md)** methods of the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.|
|[IncludeWholeCalendar](calendarsharing-includewholecalendar-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether all calendar items in the folder should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** or **[SaveAsICal](calendarsharing-saveasical-method-outlook.md)** methods of the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.|
|[Parent](calendarsharing-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[RestrictToWorkingHours](calendarsharing-restricttoworkinghours-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether calendar items that do not occur within working hours should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** or **[SaveAsICal](calendarsharing-saveasical-method-outlook.md)** methods of the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.|
|[Session](calendarsharing-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[StartDate](calendarsharing-startdate-property-outlook.md)|Returns or sets a  **Date** that represents the inclusive start date of the range of calendar items to be shared by the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.|

