---
title: AddIns Members (Word)
ms.prod: WORD
ms.assetid: 351dc3b6-6fb1-7d68-16d7-e377b433130a
---


# AddIns Members (Word)
A collection of  **AddIn** objects that represents all the add-ins available to Word, regardless of whether or not they are currently loaded. The **AddIns** collection includes global templates or Word add-in libraries (WLLs) displayed in the **Templates and Add-ins** dialog box.

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](addins-add-method-word.md)|Returns an  **[AddIn](addin-object-word.md)** object that represents an add-in added to the list of available add-ins.|
|[Item](addins-item-method-word.md)|Returns an individual object in a collection.|
|[Unload](addins-unload-method-word.md)|Unloads all loaded add-ins and, depending on the value of the  _RemoveFromList_ argument, removes them from the **AddIns** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](addins-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Count](addins-count-property-word.md)|Returns the number of  **[AddIn](addin-object-word.md)** objects in the **AddIns** collection. Read-only **Long** .|
|[Creator](addins-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Parent](addins-parent-property-word.md)|Returns an  **Object** that represents the parent object of the **Addins** collection. This is usually an **[Application](application-object-word.md)** object.|

