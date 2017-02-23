---
title: NameSpace Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: d7a978a3-a2c8-6195-c5f8-af8773500456
---


# NameSpace Members (Outlook)
Represents an abstract root object for any data source.

Represents an abstract root object for any data source.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AutoDiscoverComplete](namespace-autodiscovercomplete-event-outlook.md)|Occurs after Microsoft Outlook has finished accessing the auto-discovery service of the Microsoft Exchange server that hosts the primary Exchange account and has the related information available in  **[NameSpace.AutoDiscoverXml](namespace-autodiscoverxml-property-outlook.md)** .|
|[OptionsPagesAdd](namespace-optionspagesadd-event-outlook.md)|Occurs whenever the  **Properties** dialog box for a folder is opened.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddStore](namespace-addstore-method-outlook.md)|Adds a Personal Folders (.pst) file to the current profile.|
|[AddStoreEx](namespace-addstoreex-method-outlook.md)|Adds a Personal Folders file (.pst) in the specified format to the current profile.|
|[CompareEntryIDs](namespace-compareentryids-method-outlook.md)|Returns a  **Boolean** value that indicates if two entry ID values refer to the same Outlook item.|
|[CreateContactCard](namespace-createcontactcard-method-outlook.md)|Creates an instance of a  **[ContactCard](contactcard-object-office.md)** object for the contact that is specified by the _AddressEntry_ parameter.|
|[CreateRecipient](namespace-createrecipient-method-outlook.md)|Creates a  **[Recipient](recipient-object-outlook.md)** object.|
|[CreateSharingItem](namespace-createsharingitem-method-outlook.md)|Creates a new  **[SharingItem](sharingitem-object-outlook.md)** object.|
|[Dial](namespace-dial-method-outlook.md)|Displays the  **New Call** dialog box that allows users to dial the primary phone number of a specified contact.|
|[GetAddressEntryFromID](namespace-getaddressentryfromid-method-outlook.md)|Returns an  **[AddressEntry](addressentry-object-outlook.md)** object that represents the address entry for the specified _ID_ .|
|[GetDefaultFolder](namespace-getdefaultfolder-method-outlook.md)|Returns a  **[Folder](folder-object-outlook.md)** object that represents the default folder of the requested type for the current profile; for example, obtains the default **Calendar** folder for the user who is currently logged on.|
|[GetFolderFromID](namespace-getfolderfromid-method-outlook.md)|Returns a  **[Folder](folder-object-outlook.md)** object identified by the specified entry ID (if valid).|
|[GetGlobalAddressList](namespace-getglobaladdresslist-method-outlook.md)|Returns an  **[AddressList](addresslist-object-outlook.md)** object that represents the Exchange Global Address List.|
|[GetItemFromID](namespace-getitemfromid-method-outlook.md)|Returns a Microsoft Outlook item identified by the specified entry ID (if valid). |
|[GetRecipientFromID](namespace-getrecipientfromid-method-outlook.md)|Returns the  **[Recipient](recipient-object-outlook.md)** object that is identified by the specified entry ID (if valid).|
|[GetSelectNamesDialog](namespace-getselectnamesdialog-method-outlook.md)|Obtains a  **[SelectNamesDialog](selectnamesdialog-object-outlook.md)** object for the current session.|
|[GetSharedDefaultFolder](namespace-getshareddefaultfolder-method-outlook.md)|Returns a  **[Folder](folder-object-outlook.md)** object that represents the specified default folder for the specified user.|
|[GetStoreFromID](namespace-getstorefromid-method-outlook.md)|Returns a  **[Store](store-object-outlook.md)** object that represents the store specified by _ID_ .|
|[Logoff](namespace-logoff-method-outlook.md)|Logs the user off from the current MAPI session.|
|[Logon](namespace-logon-method-outlook.md)|Logs the user on to MAPI, obtaining a MAPI session.|
|[OpenSharedFolder](namespace-opensharedfolder-method-outlook.md)|Opens a shared folder referenced through a URL or file name.|
|[OpenSharedItem](namespace-openshareditem-method-outlook.md)|Opens a shared item from a specified path or URL.|
|[PickFolder](namespace-pickfolder-method-outlook.md)|Displays the  **Pick Folder** dialog box.|
|[RemoveStore](namespace-removestore-method-outlook.md)|Removes a Personal Folders file (.pst) from the current MAPI profile or session.|
|[SendAndReceive](namespace-sendandreceive-method-outlook.md)|Initiates immediate delivery of all undelivered messages submitted in the current session, and immediate receipt of mail for all accounts in the current profile. |

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Accounts](namespace-accounts-property-outlook.md)|Returns an  **[Accounts](accounts-object-outlook.md)** collection object that represents all the **[Account](account-object-outlook.md)** objects in the current profile. Read-only.|
|[AddressLists](namespace-addresslists-property-outlook.md)|Returns an  **[AddressLists](addresslists-object-outlook.md)** collection representing a collection of the address lists available for this session. Read-only.|
|[Application](namespace-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[AutoDiscoverConnectionMode](namespace-autodiscoverconnectionmode-property-outlook.md)|Returns an  **[OlAutoDiscoverConnectionMode](olautodiscoverconnectionmode-enumeration-outlook.md)** constant that specifies the type of connection for auto-discovery of the Microsoft Exchange server that hosts the primary Exchange account. Read-only.|
|[AutoDiscoverXml](namespace-autodiscoverxml-property-outlook.md)|Returns a  **String** that represents information in XML retrieved from the auto-discovery service for the Microsoft Exchange server that hosts the primary Exchange account. Read-only.|
|[Categories](namespace-categories-property-outlook.md)|Returns or sets a  **[Categories](categories-object-outlook.md)** object that represents the set of **[Category](category-object-outlook.md)** objects that are available to the namespace. Read/write.|
|[Class](namespace-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[CurrentProfileName](namespace-currentprofilename-property-outlook.md)|Returns a  **String** representing the name of the current profile. Read-only.|
|[CurrentUser](namespace-currentuser-property-outlook.md)|Returns the display name of the currently logged-on user as a  **[Recipient](recipient-object-outlook.md)** object. Read-only.|
|[DefaultStore](namespace-defaultstore-property-outlook.md)|Returns a  **[Store](store-object-outlook.md)** object representing the default Store for the profile. Read-only.|
|[ExchangeConnectionMode](namespace-exchangeconnectionmode-property-outlook.md)|Returns an  **[OlExchangeConnectionMode](olexchangeconnectionmode-enumeration-outlook.md)** constant that indicates the connection mode of the user's primary Exchange account. Read-only.|
|[ExchangeMailboxServerName](namespace-exchangemailboxservername-property-outlook.md)|Returns a  **String** value that represents the name of the Exchange server that hosts the primary Exchange account mailbox. Read-only.|
|[ExchangeMailboxServerVersion](namespace-exchangemailboxserverversion-property-outlook.md)|Returns a  **String** value that represents the full version number of the Exchange server that hosts the primary Exchange account mailbox. Read-only.|
|[Folders](namespace-folders-property-outlook.md)|Returns the  **[Folders](folders-object-outlook.md)** collection that represents all the folders contained in the specified **[NameSpace](namespace-object-outlook.md)** . Read-only.|
|[Offline](namespace-offline-property-outlook.md)|Returns a  **Boolean** indicating **True** if Outlook is offline (not connected to an Exchange server), and **False** if online (connected to an Exchange server). Read-only.|
|[Parent](namespace-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Session](namespace-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Stores](namespace-stores-property-outlook.md)|Returns a  **[Stores](stores-object-outlook.md)** collection object that represents all the **[Store](store-object-outlook.md)** objects in the current profile. Read-only.|
|[SyncObjects](namespace-syncobjects-property-outlook.md)|Returns a  **[SyncObjects](syncobjects-object-outlook.md)** collection containing all Send\Receive groups. Read-only.|
|[Type](namespace-type-property-outlook.md)|Returns a  **String** indicating the type of the specified object. Read-only.|

