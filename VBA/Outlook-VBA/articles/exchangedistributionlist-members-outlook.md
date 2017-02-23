---
title: ExchangeDistributionList Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 89105487-3e5b-ee8b-02e0-33ad42bd2fbe
---


# ExchangeDistributionList Members (Outlook)
The  **ExchangeDistributionList** object provides detailed information about an **[AddressEntry](addressentry-object-outlook.md)** that represents an Exchange distribution list.

The  **ExchangeDistributionList** object provides detailed information about an **[AddressEntry](addressentry-object-outlook.md)** that represents an Exchange distribution list.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](exchangedistributionlist-delete-method-outlook.md)|Deletes the  **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object from the **[AddressEntries](addressentries-object-outlook.md)** collection object to which it belongs.|
|[Details](exchangedistributionlist-details-method-outlook.md)|Displays a modal dialog box that provides detailed information about an  **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object.|
|[GetContact](exchangedistributionlist-getcontact-method-outlook.md)|Returns  **Null** ( **Nothing** in Visual Basic) because the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object does not correspond to a contact in a Contacts Address Book.|
|[GetExchangeDistributionList](exchangedistributionlist-getexchangedistributionlist-method-outlook.md)|Returns the  **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object.|
|[GetExchangeDistributionListMembers](exchangedistributionlist-getexchangedistributionlistmembers-method-outlook.md)|Returns an  **[AddressEntries](addressentries-object-outlook.md)** collection that represents the members of the Exchange distribution list.|
|[GetExchangeUser](exchangedistributionlist-getexchangeuser-method-outlook.md)|Returns  **Null** ( **Nothing** in Visual Basic) because the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object does not correspond to an **[ExchangeUser](exchangeuser-object-outlook.md)** object.|
|[GetFreeBusy](exchangedistributionlist-getfreebusy-method-outlook.md)|Returns  **Null** ( **Nothing** in Visual Basic) because free-busy information is available only to individual users and not **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** objects.|
|[GetMemberOfList](exchangedistributionlist-getmemberoflist-method-outlook.md)|Returns an  **[AddressEntries](addressentries-object-outlook.md)** collection object that contains all the **[AddressEntry](addressentry-object-outlook.md)** objects representing Exchange Distribution Lists of which the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** is a member.|
|[GetOwners](exchangedistributionlist-getowners-method-outlook.md)|Returns an  **[AddressEntries](addressentries-object-outlook.md)** collection object that contains all the owners of the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** .|
|[Update](exchangedistributionlist-update-method-outlook.md)|Posts a change to the  **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object in the messaging system.|
|[GetUnifiedGroup](exchangedistributionlist-getunifiedgroup-method-outlook.md)|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](exchangedistributionlist-isunifiedgroup-method-outlook.md). |
|[GetUnifiedGroupFromStore](exchangedistributionlist-getunifiedgroupfromstore-method-outlook.md)|Determines if the object is a unified group (by way of a call to [IsUnifiedGroup](exchangedistributionlist-isunifiedgroup-method-outlook.md)) and returns the  **Outlook.Folder** object associated with the group using the[GetUnifiedGroup](exchangedistributionlist-getunifiedgroup-method-outlook.md) and **GetUnifiedGroupFromStore** methods.|
|[IsUnifiedGroup](exchangedistributionlist-isunifiedgroup-method-outlook.md)|Determines if the object is a unified group.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Address](exchangedistributionlist-address-property-outlook.md)|Returns or sets a  **String** representing the X400 e-mail address of the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read/write.|
|[AddressEntryUserType](exchangedistributionlist-addressentryusertype-property-outlook.md)|Returns  **olExchangeDistributionListAddressEntry** which is a constant from the **[OlAddressEntryUserType](oladdressentryusertype-enumeration-outlook.md)** enumeration representing the user type of the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read-only.|
|[Alias](exchangedistributionlist-alias-property-outlook.md)|Returns a  **String** representing the alias for the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read-only.|
|[Application](exchangedistributionlist-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent application (Outlook) for the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object. Read-only.|
|[Class](exchangedistributionlist-class-property-outlook.md)|Returns a constant in the  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** enumeration indicating the class of the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object. Read-only.|
|[Comments](exchangedistributionlist-comments-property-outlook.md)|Returns a  **String** representing the comments for the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read/write.|
|[DisplayType](exchangedistributionlist-displaytype-property-outlook.md)|Returns  **olDistList** which is a constant from the **[OlDisplayType](oldisplaytype-enumeration-outlook.md)** enumeration representing the nature of the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read-only.|
|[ID](exchangedistributionlist-id-property-outlook.md)|Returns a  **String** representing the unique identifier for the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read-only.|
|[Name](exchangedistributionlist-name-property-outlook.md)|Returns or sets a  **String** value that represents the display name for the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object. Read/write.|
|[Parent](exchangedistributionlist-parent-property-outlook.md)|Returns the parent  **Object** of the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object. Read-only.|
|[PrimarySmtpAddress](exchangedistributionlist-primarysmtpaddress-property-outlook.md)|Returns a  **String** representing the primary Simple Mail Transfer Protocol (SMTP) address for the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read-only.|
|[PropertyAccessor](exchangedistributionlist-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object. Read-only.|
|[Session](exchangedistributionlist-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Type](exchangedistributionlist-type-property-outlook.md)|Returns a  **String** representing the type of entry for the **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** . Read/write.|

