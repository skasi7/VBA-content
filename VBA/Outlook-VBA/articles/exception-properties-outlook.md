---
title: Exception Properties (Outlook)
ms.prod: OUTLOOK
ms.assetid: 1225d8e6-0626-46de-a0af-51b77cf52146
---


# Exception Properties (Outlook)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](exception-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[AppointmentItem](exception-appointmentitem-property-outlook.md)|Returns the  **[AppointmentItem](appointmentitem-object-outlook.md)** object that is the exception. Not valid for deleted appointments. Read-only.|
|[Class](exception-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Deleted](exception-deleted-property-outlook.md)|Returns a  **Boolean** value that indicates whether the **[AppointmentItem](appointmentitem-object-outlook.md)** was deleted from the recurring pattern. Read-only|
|[OriginalDate](exception-originaldate-property-outlook.md)|Returns a  **Date** indicating the original date and time of an **[AppointmentItem](appointmentitem-object-outlook.md)** before it was altered. This property will return the original date even if the **AppointmentItem** has been deleted. However, it will not return the original time if deletion has occurred. Read-only.|
|[Parent](exception-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Session](exception-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|

