---
title: StorageItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 450983cc-543f-a832-d9bb-06911b0b0ce4
---


# StorageItem Members (Outlook)
A message object in MAPI that is always saved as a hidden item in the parent folder and stores private data for Outlook solutions.

A message object in MAPI that is always saved as a hidden item in the parent folder and stores private data for Outlook solutions.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](storageitem-delete-method-outlook.md)|Permanently removes the  **[StorageItem](storageitem-object-outlook.md)** object from the parent folder.|
|[Save](storageitem-save-method-outlook.md)|Saves the  **[StorageItem](storageitem-object-outlook.md)** .|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](storageitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](storageitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[Body](storageitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[Class](storageitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[CreationTime](storageitem-creationtime-property-outlook.md)|Returns a  **DateTime** value that indicates the creation time for the **[StorageItem](storageitem-object-outlook.md)** . Read-only.|
|[Creator](storageitem-creator-property-outlook.md)|Returns and sets the solution that created the  **[StorageItem](storageitem-object-outlook.md)** object. Read/write.|
|[EntryID](storageitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[LastModificationTime](storageitem-lastmodificationtime-property-outlook.md)|Returns a  **DateTime** value specifying the date and time that the Outlook item was last modified. Read-only.|
|[Parent](storageitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](storageitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[StorageItem](storageitem-object-outlook.md)** object. Read-only.|
|[Session](storageitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](storageitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the **[StorageItem](storageitem-object-outlook.md)** . Read-only.|
|[Subject](storageitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[UserProperties](storageitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

