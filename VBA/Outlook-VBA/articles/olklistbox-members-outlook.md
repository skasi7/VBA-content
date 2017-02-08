---
title: OlkListBox Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: b8bed0b5-6994-1492-055e-4067b232f9c4
---


# OlkListBox Members (Outlook)
A control that supports displaying a scrollable list of items.

A control that supports displaying a scrollable list of items.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterUpdate](olklistbox-afterupdate-event-outlook.md)|Occurs after the data in the control has been changed through the user interface.|
|[BeforeUpdate](olklistbox-beforeupdate-event-outlook.md)|Occurs when the data in the control is changed through the user interface and is about to be saved to the item. |
|[Change](olklistbox-change-event-outlook.md)|Occurs when the selection in the list displayed by the control changes|
|[Click](olklistbox-click-event-outlook.md)|Occurs when the user clicks inside the control.|
|[DoubleClick](olklistbox-doubleclick-event-outlook.md)|Occurs when the user double-clicks inside the control.|
|[Enter](olklistbox-enter-event-outlook.md)|Occurs when the control receives focus, immediately after the previous control's  **Exit** event.|
|[Exit](olklistbox-exit-event-outlook.md)|Occurs just after the focus passes from this control to another control on the same form.|
|[KeyDown](olklistbox-keydown-event-outlook.md)|Occurs when a user presses a key.|
|[KeyPress](olklistbox-keypress-event-outlook.md)|Occurs when the user presses an ANSI key.|
|[KeyUp](olklistbox-keyup-event-outlook.md)|Occurs when the user releases a key.|
|[MouseDown](olklistbox-mousedown-event-outlook.md)|Occurs when the user presses a mouse button on the control.|
|[MouseMove](olklistbox-mousemove-event-outlook.md)|Occurs after a mouse movement has been registered over the control.|
|[MouseUp](olklistbox-mouseup-event-outlook.md)|Occurs after the user releases a mouse button that has been pressed on the control.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddItem](olklistbox-additem-method-outlook.md)|Adds an item to the list, optionally specifying an index for the new item to appear in the list.|
|[Clear](olklistbox-clear-method-outlook.md)|Removes all objects from the list.|
|[Copy](olklistbox-copy-method-outlook.md)|Copies the current selection in the drop-down list to the clipboard.|
|[GetItem](olklistbox-getitem-method-outlook.md)|Obtains a  **String** that represents an item at the specified location in the list.|
|[GetSelected](olklistbox-getselected-method-outlook.md)|Returns a  **Boolean** that indicates if the indexed item is currently selected.|
|[RemoveItem](olklistbox-removeitem-method-outlook.md)|Removes the specified item from the list.|
|[SetItem](olklistbox-setitem-method-outlook.md)|Sets the item at the specified location in the list to the specified value.|
|[SetSelected](olklistbox-setselected-method-outlook.md)|Sets the selected state of an item at the specified location in the list to the given  _Selected_ value.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[BackColor](olklistbox-backcolor-property-outlook.md)|Returns or sets a  **Long** that indicates the background color of the control. Read/write.|
|[BorderStyle](olklistbox-borderstyle-property-outlook.md)|Returns or sets an  **[OlBorderStyle](olborderstyle-enumeration-outlook.md)** constant that defines the style of the border around the control. Read/write.|
|[Enabled](olklistbox-enabled-property-outlook.md)|Returns or sets a  **Boolean** that indicates if the control is allowed to function. Read/write.|
|[Font](olklistbox-font-property-outlook.md)|Returns a  **StdFont** that represents the font used to render the text inside the control. Read-only.|
|[ForeColor](olklistbox-forecolor-property-outlook.md)|Returns or sets a  **Long** that indicates the foreground color of the control. Read/write.|
|[ListCount](olklistbox-listcount-property-outlook.md)|Returns a  **Long** that specifies the number of elements in the drop-down list of the list box control. Read-only.|
|[ListIndex](olklistbox-listindex-property-outlook.md)|Returns or sets a  **Long** that indicates the location of the currently selected element in the list of the combo box control. Read/write.|
|[Locked](olklistbox-locked-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether or not the control is locked from being changed. Read/write.|
|[MatchEntry](olklistbox-matchentry-property-outlook.md)|Returns or sets an  **[olMatchEntry](olmatchentry-enumeration-outlook.md)** constant that indicates how the control searches its list as the user types. Read/write.|
|[MouseIcon](olklistbox-mouseicon-property-outlook.md)|Returns or sets a  **StdPicture** that represents a custom picture to the mouse cursor for this control. Read/write.|
|[MousePointer](olklistbox-mousepointer-property-outlook.md)|Returns or sets an  **[OlMousePointer](olmousepointer-enumeration-outlook.md)** constant that specifies the type of pointer displayed when the user positions the mouse over the control. Read/write.|
|[MultiSelect](olklistbox-multiselect-property-outlook.md)|Returns or sets an  **[OlMultiSelect](olmultiselect-enumeration-outlook.md)** constant that specifies the range of items that can be selected in the list box control. Read/write.|
|[Text](olklistbox-text-property-outlook.md)|Returns or sets a  **String** that is the text displayed in the control. Read/write.|
|[TextAlign](olklistbox-textalign-property-outlook.md)|Returns or sets an  **[OlTextAlign](oltextalign-enumeration-outlook.md)** constant that specifies how text is aligned in the control. Read/write.|
|[TopIndex](olklistbox-topindex-property-outlook.md)|Returns or sets a  **Long** that represents the index of the item at the top of the displayed portion of the list. Read/write.|
|[Value](olklistbox-value-property-outlook.md)|Returns or sets a  **Variant** that represents the content selected in the list displayed by the control. Read/write.|

