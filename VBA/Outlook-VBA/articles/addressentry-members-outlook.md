---
title: AddressEntry Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 74c88069-aec4-952b-556f-03873fbb488b
---


# AddressEntry Members (Outlook)
Represents a person, group, or public folder to which the messaging system can deliver messages.

Represents a person, group, or public folder to which the messaging system can deliver messages.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](addressentry-delete-method-outlook.md)|Deletes an object from the collection.|
|[Details](addressentry-details-method-outlook.md)|Displays a modeless dialog box that provides detailed information about an  **[AddressEntry](addressentry-object-outlook.md)** object.|
|[GetContact](addressentry-getcontact-method-outlook.md)|Returns a  **[ContactItem](contactitem-object-outlook.md)** object that represents the **[AddressEntry](addressentry-object-outlook.md)** , if the **AddressEntry** corresponds to a contact in an Outlook Contacts Address Book (CAB).|
|[GetExchangeDistributionList](addressentry-getexchangedistributionlist-method-outlook.md)|Returns an  **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object that represents the **[AddressEntry](addressentry-object-outlook.md)** if the **AddressEntry** belongs to an Exchange **[AddressList](addresslist-object-outlook.md)** object such as the Global Address List (GAL) and corresponds to an Exchange distribution list.|
|[GetExchangeUser](addressentry-getexchangeuser-method-outlook.md)|Returns an  **[ExchangeUser](exchangeuser-object-outlook.md)** object that represents the **[AddressEntry](addressentry-object-outlook.md)** if the **AddressEntry** belongs to an Exchange **[AddressList](addresslist-object-outlook.md)** object such as the Global Address List (GAL) and corresponds to an Exchange user.|
|[GetFreeBusy](addressentry-getfreebusy-method-outlook.md)|Returns a  **String** value that represents the availability of the individual user for a period of 30 days from the start date, beginning at midnight of the date specified.|
|[Update](addressentry-update-method-outlook.md)|Posts a change to the  **[AddressEntry](addressentry-object-outlook.md)** object in the messaging system.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Address](addressentry-address-property-outlook.md)|Returns or sets a  **String** representing the e-mail address of the **[AddressEntry](addressentry-object-outlook.md)** . Read/write.|
|[AddressEntryUserType](addressentry-addressentryusertype-property-outlook.md)|Returns a constant from the  **[OlAddressEntryUserType](oladdressentryusertype-enumeration-outlook.md)** enumeration representing the user type of the **[AddressEntry](addressentry-object-outlook.md)** . Read-only.|
|[Application](addressentry-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Class](addressentry-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[DisplayType](addressentry-displaytype-property-outlook.md)|Returns a constant belonging to the  **[OlDisplayType](oldisplaytype-enumeration-outlook.md)** enumeration that describes the nature of the **[AddressEntry](addressentry-object-outlook.md)** . Read-only.|
|[ID](addressentry-id-property-outlook.md)|Returns a  **String** representing the unique identifier for the object. Read-only.|
|[Name](addressentry-name-property-outlook.md)|Returns or sets a  **String** value that represents the display name for the object. Read/write.|
|[Parent](addressentry-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](addressentry-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[AddressEntry](addressentry-object-outlook.md)** object. Read-only.|
|[Session](addressentry-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Type](addressentry-type-property-outlook.md)|Returns or sets a  **String** representing the type of entry for this address such as an Internet Address, MacMail Address, or Microsoft Mail Address. Read/write.|

