---
title: NameSpace Object (Outlook)
keywords: vbaol11.chm3000
f1_keywords:
- vbaol11.chm3000
ms.prod: OUTLOOK
api_name:
- Outlook.NameSpace
ms.assetid: f0dcaa19-07f5-5d42-a3bf-2e42b7885644
---


# NameSpace Object (Outlook)

Represents an abstract root object for any data source.


## Remarks

The object itself provides methods for logging in and out, accessing storage objects directly by ID, accessing certain special default folders directly, and accessing data sources owned by other users.

Use  **[GetNameSpace](http://msdn.microsoft.com/library/application-getnamespace-method-outlook%28Office.15%29.aspx)** ("MAPI") to return the Outlook **NameSpace** object from the **[Application](http://msdn.microsoft.com/library/application-object-outlook%28Office.15%29.aspx)** object.

The only data source supported is MAPI, which allows access to all Outlook data stored in the user's mail stores.


## Events



|**Name**|
|:-----|
|[AutoDiscoverComplete](http://msdn.microsoft.com/library/namespace-autodiscovercomplete-event-outlook%28Office.15%29.aspx)|
|[OptionsPagesAdd](http://msdn.microsoft.com/library/namespace-optionspagesadd-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddStore](http://msdn.microsoft.com/library/namespace-addstore-method-outlook%28Office.15%29.aspx)|
|[AddStoreEx](http://msdn.microsoft.com/library/namespace-addstoreex-method-outlook%28Office.15%29.aspx)|
|[CompareEntryIDs](http://msdn.microsoft.com/library/namespace-compareentryids-method-outlook%28Office.15%29.aspx)|
|[CreateContactCard](http://msdn.microsoft.com/library/namespace-createcontactcard-method-outlook%28Office.15%29.aspx)|
|[CreateRecipient](http://msdn.microsoft.com/library/namespace-createrecipient-method-outlook%28Office.15%29.aspx)|
|[CreateSharingItem](http://msdn.microsoft.com/library/namespace-createsharingitem-method-outlook%28Office.15%29.aspx)|
|[Dial](http://msdn.microsoft.com/library/namespace-dial-method-outlook%28Office.15%29.aspx)|
|[GetAddressEntryFromID](http://msdn.microsoft.com/library/namespace-getaddressentryfromid-method-outlook%28Office.15%29.aspx)|
|[GetDefaultFolder](http://msdn.microsoft.com/library/namespace-getdefaultfolder-method-outlook%28Office.15%29.aspx)|
|[GetFolderFromID](http://msdn.microsoft.com/library/namespace-getfolderfromid-method-outlook%28Office.15%29.aspx)|
|[GetGlobalAddressList](http://msdn.microsoft.com/library/namespace-getglobaladdresslist-method-outlook%28Office.15%29.aspx)|
|[GetItemFromID](http://msdn.microsoft.com/library/namespace-getitemfromid-method-outlook%28Office.15%29.aspx)|
|[GetRecipientFromID](http://msdn.microsoft.com/library/namespace-getrecipientfromid-method-outlook%28Office.15%29.aspx)|
|[GetSelectNamesDialog](http://msdn.microsoft.com/library/namespace-getselectnamesdialog-method-outlook%28Office.15%29.aspx)|
|[GetSharedDefaultFolder](http://msdn.microsoft.com/library/namespace-getshareddefaultfolder-method-outlook%28Office.15%29.aspx)|
|[GetStoreFromID](http://msdn.microsoft.com/library/namespace-getstorefromid-method-outlook%28Office.15%29.aspx)|
|[Logoff](http://msdn.microsoft.com/library/namespace-logoff-method-outlook%28Office.15%29.aspx)|
|[Logon](http://msdn.microsoft.com/library/namespace-logon-method-outlook%28Office.15%29.aspx)|
|[OpenSharedFolder](http://msdn.microsoft.com/library/namespace-opensharedfolder-method-outlook%28Office.15%29.aspx)|
|[OpenSharedItem](http://msdn.microsoft.com/library/namespace-openshareditem-method-outlook%28Office.15%29.aspx)|
|[PickFolder](http://msdn.microsoft.com/library/namespace-pickfolder-method-outlook%28Office.15%29.aspx)|
|[RemoveStore](http://msdn.microsoft.com/library/namespace-removestore-method-outlook%28Office.15%29.aspx)|
|[SendAndReceive](http://msdn.microsoft.com/library/namespace-sendandreceive-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Accounts](http://msdn.microsoft.com/library/namespace-accounts-property-outlook%28Office.15%29.aspx)|
|[AddressLists](http://msdn.microsoft.com/library/namespace-addresslists-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/namespace-application-property-outlook%28Office.15%29.aspx)|
|[AutoDiscoverConnectionMode](http://msdn.microsoft.com/library/namespace-autodiscoverconnectionmode-property-outlook%28Office.15%29.aspx)|
|[AutoDiscoverXml](http://msdn.microsoft.com/library/namespace-autodiscoverxml-property-outlook%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/namespace-categories-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/namespace-class-property-outlook%28Office.15%29.aspx)|
|[CurrentProfileName](http://msdn.microsoft.com/library/namespace-currentprofilename-property-outlook%28Office.15%29.aspx)|
|[CurrentUser](http://msdn.microsoft.com/library/namespace-currentuser-property-outlook%28Office.15%29.aspx)|
|[DefaultStore](http://msdn.microsoft.com/library/namespace-defaultstore-property-outlook%28Office.15%29.aspx)|
|[ExchangeConnectionMode](http://msdn.microsoft.com/library/namespace-exchangeconnectionmode-property-outlook%28Office.15%29.aspx)|
|[ExchangeMailboxServerName](http://msdn.microsoft.com/library/namespace-exchangemailboxservername-property-outlook%28Office.15%29.aspx)|
|[ExchangeMailboxServerVersion](http://msdn.microsoft.com/library/namespace-exchangemailboxserverversion-property-outlook%28Office.15%29.aspx)|
|[Folders](http://msdn.microsoft.com/library/namespace-folders-property-outlook%28Office.15%29.aspx)|
|[Offline](http://msdn.microsoft.com/library/namespace-offline-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/namespace-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/namespace-session-property-outlook%28Office.15%29.aspx)|
|[Stores](http://msdn.microsoft.com/library/namespace-stores-property-outlook%28Office.15%29.aspx)|
|[SyncObjects](http://msdn.microsoft.com/library/namespace-syncobjects-property-outlook%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/namespace-type-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[NameSpace Object Members](http://msdn.microsoft.com/library/namespace-members-outlook%28Office.15%29.aspx)
[How to: Obtain and Log On to an Instance of Outlook](http://msdn.microsoft.com/library/obtain-and-log-on-to-an-instance-of-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
