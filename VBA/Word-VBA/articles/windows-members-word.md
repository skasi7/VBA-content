---
title: Windows Members (Word)
ms.prod: WORD
ms.assetid: 4a0863e6-b72c-fc50-95ac-3e9a0d231626
---


# Windows Members (Word)
A collection of  **[Window](window-object-word.md)** objects that represent all the available windows. The **Windows** collection for the **Application** object contains all the windows in the application, whereas the **Windows** collection for the **Document** object contains only the windows that display the specified document.

A collection of  **[Window](window-object-word.md)** objects that represent all the available windows. The **Windows** collection for the **Application** object contains all the windows in the application, whereas the **Windows** collection for the **Document** object contains only the windows that display the specified document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](windows-add-method-word.md)|Returns a  **Window** object that represents a new window of a document.|
|[Arrange](windows-arrange-method-word.md)|Arranges all open document windows in the application workspace.|
|[BreakSideBySide](windows-breaksidebyside-method-word.md)|Ends side by side mode if two windows are in side by side mode. Returns a  **Boolean** that represents whether the method was successful.|
|[CompareSideBySideWith](windows-comparesidebysidewith-method-word.md)|Opens two windows in side by side mode. Returns a **Boolean** .|
|[Item](windows-item-method-word.md)|Returns an individual  **Window** object in a collection.|
|[ResetPositionsSideBySide](windows-resetpositionssidebyside-method-word.md)|Resets two document windows that are in the  **Compare side by side with** view mode.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](windows-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Count](windows-count-property-word.md)|Returns a  **Long** that represents the number of windows in the collection. Read-only.|
|[Creator](windows-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Parent](windows-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Windows** object.|
|[SyncScrollingSideBySide](windows-syncscrollingsidebyside-property-word.md)| **True** enables scrolling the contents of the windows at the same time. Read/write **Boolean** .|

