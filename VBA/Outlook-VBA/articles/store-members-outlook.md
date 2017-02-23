---
title: Store Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 84c1d423-e507-0b3b-6570-33829b94be04
---


# Store Members (Outlook)
Represents a file on the local computer or a network drive that stores e-mail messages and other items for an account in the current profile.

Represents a file on the local computer or a network drive that stores e-mail messages and other items for an account in the current profile.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[GetDefaultFolder](store-getdefaultfolder-method-outlook.md)|Returns a  **[Folder](folder-object-outlook.md)** object that represents the default folder in the store and that is of the type specified by the _FolderType_ argument.|
|[GetRootFolder](store-getrootfolder-method-outlook.md)|Returns a  **[Folder](folder-object-outlook.md)** object representing the root-level folder of the **[Store](store-object-outlook.md)** . Read-only.|
|[GetRules](store-getrules-method-outlook.md)|Returns a  **[Rules](rules-object-outlook.md)** collection object that contains the **[Rule](rule-object-outlook.md)** objects defined for the current session.|
|[GetSearchFolders](store-getsearchfolders-method-outlook.md)|Returns a  **[Folders](folders-object-outlook.md)** collection object that represents the search folders defined for the **[Store](store-object-outlook.md)** object.|
|[GetSpecialFolder](store-getspecialfolder-method-outlook.md)|Returns a  **[Folder](folder-object-outlook.md)** object for a special folder specified by _FolderType_ in a given store.|
|[RefreshQuotaDisplay](store-refreshquotadisplay-method-outlook.md)|Refreshes the store quota information that is displayed in the status bar in the explorer window.|
|[CreateUnifiedGroup](store-createunifiedgroup-method-outlook.md)|Enables a unified group to be created.|
|[DeleteUnifiedGroup](store-deleteunifiedgroup-method-outlook.md)|Enables a unified group to be deleted.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](store-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Categories](store-categories-property-outlook.md)|Returns a  **[Categories](categories-object-outlook.md)** collection that represents all of the categories that are defined for the **[Store](store-object-outlook.md)** . Read-only.|
|[Class](store-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[DisplayName](store-displayname-property-outlook.md)|Returns a  **String** representing the display name of the **[Store](store-object-outlook.md)** object. Read-only.|
|[ExchangeStoreType](store-exchangestoretype-property-outlook.md)|Returns a constant in the  **[OlExchangeStoreType](olexchangestoretype-enumeration-outlook.md)** enumeration that indicates the type of an Exchange store. Read-only.|
|[FilePath](store-filepath-property-outlook.md)|Returns a  **String** representing the full file path for a Personal Folders File (.pst) or an Offline Folder File (.ost) store. Read-only.|
|[IsCachedExchange](store-iscachedexchange-property-outlook.md)|Returns a  **Boolean** that indicates if the **[Store](store-object-outlook.md)** is a cached Exchange store. Read-only.|
|[IsConversationEnabled](store-isconversationenabled-property-outlook.md)|Returns a  **Boolean** value that is **True** if the store supports Conversation view. Read-only.|
|[IsDataFileStore](store-isdatafilestore-property-outlook.md)|Returns a  **Boolean** that indicates if the **[Store](store-object-outlook.md)** is a store for an Outlook data file, which is either a Personal Folders File (.pst) or an Offline Folder File (.ost). Read-only.|
|[IsInstantSearchEnabled](store-isinstantsearchenabled-property-outlook.md)|Returns a  **Boolean** that indicates whether Instant Search is enabled and operational on a store. Read-only.|
|[IsOpen](store-isopen-property-outlook.md)|Returns a  **Boolean** that indicates if the **[Store](store-object-outlook.md)** is open. Read-only.|
|[Parent](store-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](store-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[Store](store-object-outlook.md)** object. Read-only.|
|[Session](store-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[StoreID](store-storeid-property-outlook.md)|Returns a  **String** identifying the **[Store](store-object-outlook.md)** . Read-only.|

