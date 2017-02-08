---
title: Form Events (Access)
ms.prod: ACCESS
ms.assetid: 2938861a-c822-4276-8b32-0784405929e2
---


# Form Events (Access)
This object has the following events:

## Events



|**Name**|**Description**|
|:-----|:-----|
|[Activate](form-activate-event-access.md)|The Activate event occurs when a form receives the focus and becomes the active window.|
|[AfterDelConfirm](form-afterdelconfirm-event-access.md)|The  **AfterDelConfirm** event occurs after the user confirms the deletions and the records are actually deleted or when the deletions are canceled.|
|[AfterFinalRender](form-afterfinalrender-event-access.md)|Occurs after all elements in the specified PivotChart view have been rendered.|
|[AfterInsert](form-afterinsert-event-access.md)|The  **AfterInsert** event occurs after a new record is added.|
|[AfterLayout](form-afterlayout-event-access.md)|Occurs after all charts in the specfied PivotChart view have been laid out, but before they have been rendered.|
|[AfterRender](form-afterrender-event-access.md)|Occurs after the object represented by the  _chartObject_ argument has been rendered.|
|[AfterUpdate](form-afterupdate-event-access.md)|The  **AfterUpdate** event occurs after changed data in a control or record is updated.|
|[ApplyFilter](form-applyfilter-event-access.md)|Occurs when a filter is applied to a form.|
|[BeforeDelConfirm](form-beforedelconfirm-event-access.md)|The  **BeforeDelConfirm** event occurs after the user deletes to the buffer one or more records, but before Microsoft Access displays a dialog box asking the user to confirm the deletions.|
|[BeforeInsert](form-beforeinsert-event-access.md)|The BeforeInsert event occurs when the user types the first character in a new record, but before the record is actually created.|
|[BeforeQuery](form-beforequery-event-access.md)|Occurs when the specified PivotTable view queries its data source.|
|[BeforeRender](form-beforerender-event-access.md)|Occurs before any object in the specified PivotChart view has been rendered.|
|[BeforeScreenTip](form-beforescreentip-event-access.md)|Occurs before a ScreenTip is displayed for an element in a PivotChart view or PivotTable view.|
|[BeforeUpdate](form-beforeupdate-event-access.md)|The  **BeforeUpdate** event occurs before changed data in a control or record is updated.|
|[Click](form-click-event-access.md)|The  **Click** event occurs when the user presses and then releases a mouse button over an object.|
|[Close](form-close-event-access.md)|The  **Close** event occurs when a form is closed and removed from the screen.|
|[CommandBeforeExecute](form-commandbeforeexecute-event-access.md)|Occurs before a specified command is executed. Use this event when you want to impose certain restrictions before a particular command is executed.|
|[CommandChecked](form-commandchecked-event-access.md)|Occurs when the specified Microsoft Office Web Component determines whether the specified command is checked.|
|[CommandEnabled](form-commandenabled-event-access.md)|Occurs when the specified Microsoft Office Web Component determines whether the specified command is enabled.|
|[CommandExecute](form-commandexecute-event-access.md)|Occurs after the specified command is executed. Use this event when you want to execute a set of commands after a particular command is executed.|
|[Current](form-current-event-access.md)|Occurs when the focus moves to a record, making it the current record, or when the form is refreshed or requeried.|
|[DataChange](form-datachange-event-access.md)|Occurs when certain properties are changed or when certain methods are executed in the specified PivotTable view.|
|[DataSetChange](form-datasetchange-event-access.md)|Occurs whenever the specified PivotTable view is data-bound and the data set changes — for example, when a filter operation takes place. This event also occurs when initial data is available from the data source.|
|[DblClick](form-dblclick-event-access.md)|The  **DblClick** event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.|
|[Deactivate](form-deactivate-event-access.md)|The  **Deactivate** event occurs when a form loses the focus to a Table, Query, Form, Report, Macro, or Module window, or to the Database window.|
|[Delete](form-delete-event-access.md)|Occurs when the user performs some action, such as pressing the DEL key, to delete a record, but before the record is actually deleted.|
|[Dirty](form-dirty-event-access.md)|The Dirty event occurs when the contents of the specified control changes.|
|[Error](form-error-event-access.md)|The Error event occurs when a run-time error is produced in Microsoft Access when a form has the focus.|
|[Filter](form-filter-event-access.md)|Occurs when the user opens a filter window by clicking  **Filter by Form**,  **Advanced Filter/Sort**, or  **Server Filter By Form**.|
|[GotFocus](form-gotfocus-event-access.md)|The  **GotFocus** event occurs when the specified object receives the focus.|
|[KeyDown](form-keydown-event-access.md)|The  **KeyDown** event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyPress](form-keypress-event-access.md)|The  **KeyPress** event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[KeyUp](form-keyup-event-access.md)|The  **KeyUp** event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
|[Load](form-load-event-access.md)|Occurs when a form is opened and its records are displayed.|
|[LostFocus](form-lostfocus-event-access.md)|The  **LostFocus** event occurs when the specified object loses the focus.|
|[MouseDown](form-mousedown-event-access.md)|The  **MouseDown** event occurs when the user presses a mouse button.|
|[MouseMove](form-mousemove-event-access.md)|The  **MouseMove** event occurs when the user moves the mouse.|
|[MouseUp](form-mouseup-event-access.md)|The  **MouseUp** event occurs when the user releases a mouse button.|
|[MouseWheel](form-mousewheel-event-access.md)|Occurs when the user rolls the mouse wheel in Form View, Split Form View, Datasheet View, Layout View, PivotChart View, or PivotTable View.|
|[OnConnect](form-onconnect-event-access.md)|Occurs when the specified PivotTable view connects to a data source.|
|[OnDisconnect](form-ondisconnect-event-access.md)|Occurs when the specified PivotTable view disconnects from a data source.|
|[Open](form-open-event-access.md)|The  **Open** event occurs when a form is opened, but before the first record is displayed.|
|[PivotTableChange](form-pivottablechange-event-access.md)|Occurs whenever the specified PivotTable view field, field set, or total is added or deleted.|
|[Query](form-query-event-access.md)|Occurs whenever the specified PivotTable view query becomes necessary. The query may not occur immediately; it may be delayed until the new data is displayed.|
|[Resize](form-resize-event-access.md)|The  **Resize** event occurs when a form is opened and whenever the size of a form changes.|
|[SelectionChange](form-selectionchange-event-access.md)|Occurs whenever the user makes a new selection in a PivotChart view or PivotTable view.|
|[Timer](form-timer-event-access.md)|The  **Timer** event occurs for a form at regular intervals as specified by the form's **[TimerInterval](form-timerinterval-property-access.md)** property.|
|[Undo](form-undo-event-access.md)|Occurs when the user undoes a change.|
|[Unload](form-unload-event-access.md)|The  **Unload** event occurs after a form is closed but before it's removed from the screen. When the form is reloaded, Microsoft Access redisplays the form and reinitializes the contents of all its controls.|
|[ViewChange](form-viewchange-event-access.md)|Occurs whenever the specified PivotChart view or PivotTable view is redrawn.|

