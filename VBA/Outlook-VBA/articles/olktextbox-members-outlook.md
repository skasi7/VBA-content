---
title: OlkTextBox Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: f4a5f9ea-15f7-164e-d7ca-77a0842105c8
---


# OlkTextBox Members (Outlook)
A control that supports a single or multiple-line data entry.

A control that supports a single or multiple-line data entry.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterUpdate](olktextbox-afterupdate-event-outlook.md)|Occurs after the data in the control has been changed through the user interface.|
|[BeforeUpdate](olktextbox-beforeupdate-event-outlook.md)|Occurs when the data in the control is changed through the user interface and is about to be saved to the item. |
|[Change](olktextbox-change-event-outlook.md)|Occurs when the  **[Value](olktextbox-value-property-outlook.md)** property changes.|
|[Click](olktextbox-click-event-outlook.md)|Occurs when the user clicks inside the control.|
|[DoubleClick](olktextbox-doubleclick-event-outlook.md)|Occurs when the user double-clicks inside the control.|
|[Enter](olktextbox-enter-event-outlook.md)|Occurs when the control receives focus, immediately after the previous control's  **Exit** event.|
|[Exit](olktextbox-exit-event-outlook.md)|Occurs just after the focus passes from this control to another control on the same form.|
|[KeyDown](olktextbox-keydown-event-outlook.md)|Occurs when a user presses a key.|
|[KeyPress](olktextbox-keypress-event-outlook.md)|Occurs when the user presses an ANSI key.|
|[KeyUp](olktextbox-keyup-event-outlook.md)|Occurs when the user releases a key.|
|[MouseDown](olktextbox-mousedown-event-outlook.md)|Occurs when the user presses a mouse button on the control.|
|[MouseMove](olktextbox-mousemove-event-outlook.md)|Occurs after a mouse movement has been registered over the control.|
|[MouseUp](olktextbox-mouseup-event-outlook.md)|Occurs after the user releases a mouse button that has been pressed on the control.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Clear](olktextbox-clear-method-outlook.md)|Removes the text in the text box.|
|[Copy](olktextbox-copy-method-outlook.md)|Copies the contents of the control to the clipboard.|
|[Cut](olktextbox-cut-method-outlook.md)|Removes the contents of the control and copies the contents to the clipboard.|
|[Paste](olktextbox-paste-method-outlook.md)|Pastes the contents of the clipboard in the control. |

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AutoSize](olktextbox-autosize-property-outlook.md)|Returns or sets a  **Boolean** that automatically sizes the control to display the entire contents. Read/write.|
|[AutoTab](olktextbox-autotab-property-outlook.md)|Returns or sets a  **Boolean** that specifies if a tab is inserted automatically when the control has been filled to the maximum length specified. Read/write.|
|[AutoWordSelect](olktextbox-autowordselect-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether a word or a character is the basic unit used to extend a selection. Read/write.|
|[BackColor](olktextbox-backcolor-property-outlook.md)|Returns or sets a  **Long** that indicates the background color of the control. Read/write.|
|[BorderStyle](olktextbox-borderstyle-property-outlook.md)|Returns or sets an  **[OlBorderStyle](olborderstyle-enumeration-outlook.md)** constant that defines the style of the border around the control. Read/write.|
|[DragBehavior](olktextbox-dragbehavior-property-outlook.md)|Returns or sets an  **[OlDragBehavior Enumeration](oldragbehavior-enumeration-outlook.md)** constant that indicates whether the system enables the drag-and-drop feature for this control. Read/write.|
|[Enabled](olktextbox-enabled-property-outlook.md)|Returns or sets a  **Boolean** that indicates if the control is allowed to function. Read/write.|
|[EnterFieldBehavior](olktextbox-enterfieldbehavior-property-outlook.md)|Returns or sets an  **[olEnterFieldBehavior](olenterfieldbehavior-enumeration-outlook.md)** constant that specifies the selection behavior when entering the control. Read/write.|
|[EnterKeyBehavior](olktextbox-enterkeybehavior-property-outlook.md)|Returns or sets a  **Boolean** that defines the way the **ENTER** key behaves in the control. Read/write.|
|[Font](olktextbox-font-property-outlook.md)|Returns a  **StdFont** that represents the font used to render the text inside the control. Read-only.|
|[ForeColor](olktextbox-forecolor-property-outlook.md)|Returns or sets a  **Long** that indicates the foreground color of the control. Read/write.|
|[HideSelection](olktextbox-hideselection-property-outlook.md)|Returns or sets a  **Boolean** that specifies if a selection is displayed or hidden for the control when the control loses focus. Read/write.|
|[IntegralHeight](olktextbox-integralheight-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether this control displays full lines of text. Read/write.|
|[Locked](olktextbox-locked-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether or not the control is locked from being changed. Read/write.|
|[MaxLength](olktextbox-maxlength-property-outlook.md)|Returns or sets a  **Long** that specifies the maximum number of characters for the **[Value](olktextbox-value-property-outlook.md)** of this control. Read/write.|
|[MouseIcon](olktextbox-mouseicon-property-outlook.md)|Returns or sets a  **StdPicture** that represents a custom picture to the mouse cursor for this control. Read/write.|
|[MousePointer](olktextbox-mousepointer-property-outlook.md)|Returns or sets an  **[OlMousePointer](olmousepointer-enumeration-outlook.md)** constant that specifies the type of pointer displayed when the user positions the mouse over the control. Read/write.|
|[MultiLine](olktextbox-multiline-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether a control can accept and display multiple lines of text. Read/write.|
|[PasswordChar](olktextbox-passwordchar-property-outlook.md)|Returns or sets a  **String** that specifies a placeholder character that is to be displayed repetitively as a string instead of the actual characters entered in the text box. Read/write.|
|[Scrollbars](olktextbox-scrollbars-property-outlook.md)|Returns or sets an  **[olScrollBars](olscrollbars-enumeration-outlook.md)** constant that specifies whether the control has a vertical scroll bar, horizontal scroll bar, or both. Read/write.|
|[SelectionMargin](olktextbox-selectionmargin-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether the user can select a line of text by clicking in the region to the left of the text. Read/write.|
|[SelLength](olktextbox-sellength-property-outlook.md)|Returns or sets a  **Long** that specifies the number of characters in the current selection. Read/write.|
|[SelStart](olktextbox-selstart-property-outlook.md)|Returns or sets a  **Long** that specifies either the starting point of the selected text or the insertion point if no text has been selected. Read/write.|
|[SelText](olktextbox-seltext-property-outlook.md)|Returns a  **String** that represents the currently selected portion of the value of the text box. Read-only.|
|[TabKeyBehavior](olktextbox-tabkeybehavior-property-outlook.md)|Returns or sets a  **Boolean** that specifies how the control responds to the **TAB** key. Read/write.|
|[Text](olktextbox-text-property-outlook.md)|Returns or sets a  **String** that is the text displayed in the control. Read/write.|
|[TextAlign](olktextbox-textalign-property-outlook.md)|Returns or sets an  **[OlTextAlign](oltextalign-enumeration-outlook.md)** constant that specifies how text is aligned in the control. Read/write.|
|[Value](olktextbox-value-property-outlook.md)|Returns or sets a  **Variant** that represents the content of the control. Read/write.|
|[WordWrap](olktextbox-wordwrap-property-outlook.md)|Returns or sets a  **Boolean** that specifies whether the contents of a control automatically wrap at the end of a line. Read/write.|

