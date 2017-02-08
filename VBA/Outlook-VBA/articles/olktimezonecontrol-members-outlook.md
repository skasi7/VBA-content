---
title: OlkTimeZoneControl Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 350ded4c-0118-c278-dabe-c6139aeba1e9
---


# OlkTimeZoneControl Members (Outlook)
A control that supports a selection from a drop-down list of time zones.

A control that supports a selection from a drop-down list of time zones.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterUpdate](olktimezonecontrol-afterupdate-event-outlook.md)|Occurs after the data in the control has been changed through the user interface.|
|[BeforeUpdate](olktimezonecontrol-beforeupdate-event-outlook.md)|Occurs when the data in the control is changed through the user interface and is about to be saved to the item. |
|[Change](olktimezonecontrol-change-event-outlook.md)|Occurs when the  **[Value](olktimezonecontrol-value-property-outlook.md)** property changes.|
|[Click](olktimezonecontrol-click-event-outlook.md)|Occurs when the user clicks inside the control.|
|[DoubleClick](olktimezonecontrol-doubleclick-event-outlook.md)|Occurs when the user double-clicks inside the control.|
|[DropButtonClick](olktimezonecontrol-dropbuttonclick-event-outlook.md)|Occurs when the user clicks the drop button to expand the drop-down list in the time zone control, or when the  **[DropDown](olktimezonecontrol-dropdown-method-outlook.md)** method is called programmatically.|
|[Enter](olktimezonecontrol-enter-event-outlook.md)|Occurs when the control receives focus, immediately after the previous control's  **Exit** event.|
|[Exit](olktimezonecontrol-exit-event-outlook.md)|Occurs just after the focus passes from this control to another control on the same form.|
|[KeyDown](olktimezonecontrol-keydown-event-outlook.md)|Occurs when a user presses a key.|
|[KeyPress](olktimezonecontrol-keypress-event-outlook.md)|Occurs when the user presses an ANSI key.|
|[KeyUp](olktimezonecontrol-keyup-event-outlook.md)|Occurs when the user releases a key.|
|[MouseDown](olktimezonecontrol-mousedown-event-outlook.md)|Occurs when the user presses a mouse button on the control.|
|[MouseMove](olktimezonecontrol-mousemove-event-outlook.md)|Occurs after a mouse movement has been registered over the control.|
|[MouseUp](olktimezonecontrol-mouseup-event-outlook.md)|Occurs after the user releases a mouse button that has been pressed on the control.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[DropDown](olktimezonecontrol-dropdown-method-outlook.md)|Expands the drop-down portion of the time zone control.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AppointmentTimeField](olktimezonecontrol-appointmenttimefield-property-outlook.md)|Returns or sets an  **[OlAppointmentTimeField](olappointmenttimefield-enumeration-outlook.md)** constant that specifies the time field on the appointment that the control binds against. Read/write.|
|[BorderStyle](olktimezonecontrol-borderstyle-property-outlook.md)|Returns or sets an  **[OlBorderStyle](olborderstyle-enumeration-outlook.md)** constant that defines the style of the border around the control. Read/write.|
|[Enabled](olktimezonecontrol-enabled-property-outlook.md)|Returns or sets a  **Boolean** that indicates if the control is allowed to function. Read/write.|
|[Locked](olktimezonecontrol-locked-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether or not the control is locked from being changed. Read/write.|
|[MouseIcon](olktimezonecontrol-mouseicon-property-outlook.md)|Returns or sets a  **StdPicture** that represents a custom picture to the mouse cursor for this control. Read/write.|
|[MousePointer](olktimezonecontrol-mousepointer-property-outlook.md)|Returns or sets an  **[OlMousePointer](olmousepointer-enumeration-outlook.md)** constant that specifies the type of pointer displayed when the user positions the mouse over the control. Read/write.|
|[SelectedTimeZoneIndex](olktimezonecontrol-selectedtimezoneindex-property-outlook.md)|Returns or sets an index into the  **[Application.TimeZones](application-timezones-property-outlook.md)** collection that determines the selected time zone. Read/write.|
|[Value](olktimezonecontrol-value-property-outlook.md)|Returns or sets a  **Variant** that represents the content of the control. Read/write.|

