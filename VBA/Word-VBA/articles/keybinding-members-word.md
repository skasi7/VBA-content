---
title: KeyBinding Members (Word)
ms.prod: WORD
ms.assetid: ff0776e1-3695-a392-992b-9d5a772449dc
---


# KeyBinding Members (Word)
Represents a custom key assignment in the current context. The  **KeyBinding** object is a member of the **KeyBindings** collection.

Represents a custom key assignment in the current context. The  **KeyBinding** object is a member of the **KeyBindings** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Clear](keybinding-clear-method-word.md)|Removes the specified key binding from the  **KeyBindings** collection and resets a built-in command to its default key assignment.|
|[Disable](keybinding-disable-method-word.md)|Removes the specified key combination if it is currently assigned to a command. After you use this method, the key combination has no effect.|
|[Execute](keybinding-execute-method-word.md)|Runs the command associated with the specified key combination.|
|[Rebind](keybinding-rebind-method-word.md)|Changes the command assigned to the specified key binding.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](keybinding-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Command](keybinding-command-property-word.md)|Returns the command assigned to the specified key combination. Read-only  **String** .|
|[CommandParameter](keybinding-commandparameter-property-word.md)|Returns the command parameter assigned to the specified shortcut key. Read-only  **String** .|
|[Context](keybinding-context-property-word.md)|Returns an  **Object** that represents the storage location of the specified key binding. Read-only.|
|[Creator](keybinding-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[KeyCategory](keybinding-keycategory-property-word.md)|Returns the type of item assigned to the specified key binding. Read-only  **WdKeyCategory** .|
|[KeyCode](keybinding-keycode-property-word.md)|Returns a unique number for the first key in the specified key binding. Read-only  **Long** .|
|[KeyCode2](keybinding-keycode2-property-word.md)|Returns a unique number for the second key in the specified key binding. Read-only  **Long** .|
|[KeyString](keybinding-keystring-property-word.md)|Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A). Read-only  **String** .|
|[Parent](keybinding-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **KeyBinding** object.|
|[Protected](keybinding-protected-property-word.md)| **True** if you cannot change the specified key binding in the **Customize Keyboard** dialog box. Read-only **Boolean** .|

