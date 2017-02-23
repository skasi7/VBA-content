---
title: OlkComboBox Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 618de9e2-f5b9-40d9-239e-95aeb9dce092
---


# OlkComboBox Members (Outlook)
A control that supports the display of a selection from a drop-down list of all choices.

A control that supports the display of a selection from a drop-down list of all choices.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterUpdate](olkcombobox-afterupdate-event-outlook.md)|Occurs after the data in the control has been changed through the user interface.|
|[BeforeUpdate](olkcombobox-beforeupdate-event-outlook.md)|Occurs when the data in the control is changed through the user interface and is about to be saved to the item. |
|[Change](olkcombobox-change-event-outlook.md)|Occurs when the selection in the list displayed by the control changes.|
|[Click](olkcombobox-click-event-outlook.md)|Occurs when the user clicks inside the control.|
|[DoubleClick](olkcombobox-doubleclick-event-outlook.md)|Occurs when the user double-clicks inside the control.|
|[DropButtonClick](olkcombobox-dropbuttonclick-event-outlook.md)|Occurs when the user clicks the drop button to expand the drop-down list in the combo box control, or when the  **[DropDown](olkcombobox-dropdown-method-outlook.md)** method is called programmatically.|
|[Enter](olkcombobox-enter-event-outlook.md)|Occurs when the control receives focus, immediately after the previous control's  **Exit** event.|
|[Exit](olkcombobox-exit-event-outlook.md)|Occurs just after the focus passes from this control to another control on the same form.|
|[KeyDown](olkcombobox-keydown-event-outlook.md)|Occurs when a user presses a key.|
|[KeyPress](olkcombobox-keypress-event-outlook.md)|Occurs when the user presses an ANSI key.|
|[KeyUp](olkcombobox-keyup-event-outlook.md)|Occurs when the user releases a key.|
|[MouseDown](olkcombobox-mousedown-event-outlook.md)|Occurs when the user presses a mouse button on the control.|
|[MouseMove](olkcombobox-mousemove-event-outlook.md)|Occurs after a mouse movement has been registered over the control.|
|[MouseUp](olkcombobox-mouseup-event-outlook.md)|Occurs after the user releases a mouse button that has been pressed on the control.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddItem](olkcombobox-additem-method-outlook.md)|Adds an item to the list, optionally specifying an index for the new item to appear in the list.|
|[Clear](olkcombobox-clear-method-outlook.md)|Removes all objects from the list in the combo box.|
|[Copy](olkcombobox-copy-method-outlook.md)|Copies the contents of the control to the clipboard.|
|[Cut](olkcombobox-cut-method-outlook.md)|Removes the contents of the control and copies the contents to the clipboard.|
|[DropDown](olkcombobox-dropdown-method-outlook.md)|Expands the drop-down portion of the combo box.|
|[GetItem](olkcombobox-getitem-method-outlook.md)|Obtains a  **String** that represents an item at the specified location in the list of the combo box control.|
|[Paste](olkcombobox-paste-method-outlook.md)|Pastes the contents of the clipboard in the control.|
|[RemoveItem](olkcombobox-removeitem-method-outlook.md)|Removes the specified item from the list.|
|[SetItem](olkcombobox-setitem-method-outlook.md)|Sets the item at the specified location in the list of the combo box to the specified value.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AutoSize](olkcombobox-autosize-property-outlook.md)|Returns or sets a  **Boolean** that automatically sizes the control to display the entire contents. Read/write.|
|[AutoTab](olkcombobox-autotab-property-outlook.md)|Returns or sets a  **Boolean** that specifies if a tab is inserted automatically when the control has been filled to the maximum length specified. Read/write.|
|[AutoWordSelect](olkcombobox-autowordselect-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether a word or a character is the basic unit used to extend a selection. Read/write.|
|[BackColor](olkcombobox-backcolor-property-outlook.md)|Returns or sets a  **Long** that indicates the background color of the control. Read/write.|
|[BorderStyle](olkcombobox-borderstyle-property-outlook.md)|Returns or sets an  **[OlBorderStyle](olborderstyle-enumeration-outlook.md)** constant that defines the style of the border around the control. Read/write.|
|[DragBehavior](olkcombobox-dragbehavior-property-outlook.md)|Returns or sets an  **[OlDragBehavior Enumeration](oldragbehavior-enumeration-outlook.md)** constant that indicates whether the system enables the drag-and-drop feature for this control. Read/write.|
|[Enabled](olkcombobox-enabled-property-outlook.md)|Returns or sets a  **Boolean** that indicates if the control is allowed to function. Read/write.|
|[EnterFieldBehavior](olkcombobox-enterfieldbehavior-property-outlook.md)|Returns or sets an  **[olEnterFieldBehavior](olenterfieldbehavior-enumeration-outlook.md)** constant that specifies the selection behavior when entering the control. Read/write.|
|[Font](olkcombobox-font-property-outlook.md)|Returns a  **StdFont** that represents the font used to render the text inside the control. Read-only.|
|[ForeColor](olkcombobox-forecolor-property-outlook.md)|Returns or sets a  **Long** that indicates the foreground color of the control. Read/write.|
|[HideSelection](olkcombobox-hideselection-property-outlook.md)|Returns or sets a  **Boolean** that specifies if a selection is displayed or hidden for the control when the control loses focus. Read/write.|
|[ListCount](olkcombobox-listcount-property-outlook.md)|Returns a  **Long** that specifies the number of elements in the drop-down list of the combo box control. Read-only.|
|[ListIndex](olkcombobox-listindex-property-outlook.md)|Reurns or sets a  **Long** that indicates the location of the currently selected element in the list of the combo box control. Read/write.|
|[Locked](olkcombobox-locked-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether or not the control is locked from being changed. Read/write.|
|[MaxLength](olkcombobox-maxlength-property-outlook.md)|Returns or sets a  **Long** that specifies the maximum number of characters for the **[Value](olkcombobox-value-property-outlook.md)** of this control. Read/write.|
|[MouseIcon](olkcombobox-mouseicon-property-outlook.md)|Returns or sets a  **StdPicture** that represents a custom picture to the mouse cursor for this control. Read/write.|
|[MousePointer](olkcombobox-mousepointer-property-outlook.md)|Returns or sets an  **[OlMousePointer](olmousepointer-enumeration-outlook.md)** constant that specifies the type of pointer displayed when the user positions the mouse over the control. Read/write.|
|[SelectionMargin](olkcombobox-selectionmargin-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether the user can select a line of text by clicking in the region to the left of the text. Read/write.|
|[SelLength](olkcombobox-sellength-property-outlook.md)|Returns or sets a  **Long** that specifies the number of characters in the current selection. Read/write.|
|[SelStart](olkcombobox-selstart-property-outlook.md)|Returns or sets a  **Long** that specifies either the starting point of the selected text or the insertion point if no text has been selected. Read/write.|
|[SelText](olkcombobox-seltext-property-outlook.md)|Returns a  **String** that represents the selected portion of the value of the combo box. Read-only.|
|[Style](olkcombobox-style-property-outlook.md)|Returns or sets an  **[OlComboBoxStyle](olcomboboxstyle-enumeration-outlook.md)** constant to specify how the user can choose or set the control's value. Read/write.|
|[Text](olkcombobox-text-property-outlook.md)|Returns or sets a  **String** that is the text displayed in the control. Read/write.|
|[TextAlign](olkcombobox-textalign-property-outlook.md)|Returns or sets an  **[OlTextAlign](oltextalign-enumeration-outlook.md)** constant that specifies how text is aligned in the control. Read/write.|
|[TopIndex](olkcombobox-topindex-property-outlook.md)|Returns or sets a  **Long** that represents the index of the item at the top of the displayed portion of the list in the combo box. Read/write.|
|[Value](olkcombobox-value-property-outlook.md)|Returns or sets a  **Variant** that represents the content selected in the list displayed by the control. Read/write.|

