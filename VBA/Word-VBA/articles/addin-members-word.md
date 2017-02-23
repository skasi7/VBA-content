---
title: AddIn Members (Word)
ms.prod: WORD
ms.assetid: 7bffb4a9-f948-fc97-342e-d4d46fa48913
---


# AddIn Members (Word)
Represents a single add-in, either installed or not installed. The  **AddIn** object is a member of the **[AddIns](addins-object-word.md)** collection. The **AddIns** collection contains all the add-ins available to Microsoft Word, regardless of whether they are currently loaded. The **AddIns** collection includes global templates or Word add-in libraries (WLLs) displayed in the **Templates and Add-ins** dialog box.

Represents a single add-in, either installed or not installed. The  **AddIn** object is a member of the **[AddIns](addins-object-word.md)** collection. The **AddIns** collection contains all the add-ins available to Microsoft Word, regardless of whether they are currently loaded. The **AddIns** collection includes global templates or Word add-in libraries (WLLs) displayed in the **Templates and Add-ins** dialog box.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](addin-delete-method-word.md)|Deletes the specified add-in.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](addin-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Autoload](addin-autoload-property-word.md)| **True** if the specified add-in is automatically loaded when Word is started. Add-ins located in the Startup folder in the Word program folder are automatically loaded. Read-only **Boolean** .|
|[Compiled](addin-compiled-property-word.md)| **True** if the specified add-in is a Word add-in library (WLL). **False** if the add-in is a template. Read-only **Boolean** .|
|[Creator](addin-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .|
|[Index](addin-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[Installed](addin-installed-property-word.md)| **True** if the specified add-in is installed (loaded). Add-ins that are loaded are selected in the **Templates and Add-ins** dialog box. Read/write **Boolean** .|
|[Name](addin-name-property-word.md)|Returns the name of an add-in. Read-only  **String** .|
|[Parent](addin-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **AddIn** object.|
|[Path](addin-path-property-word.md)|Returns the location of an installed add-in. Read-only  **String** .|

