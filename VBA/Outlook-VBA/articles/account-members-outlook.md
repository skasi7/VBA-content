---
title: Account Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 37759c57-d1ec-775c-cbe6-75c8f314d196
---


# Account Members (Outlook)
The  **Account** object represents an account that is defined for the current profile.

The  **Account** object represents an account that is defined for the current profile.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[GetAddressEntryFromID](account-getaddressentryfromid-method-outlook.md)|Returns an  **[AddressEntry](addressentry-object-outlook.md)** object that represents the address entry specified by the given entry ID.|
|[GetRecipientFromID](account-getrecipientfromid-method-outlook.md)|Returns the **[Recipient](recipient-object-outlook.md)** object that is identified by the given entry ID.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AccountType](account-accounttype-property-outlook.md)|Returns a constant in the  **[OlAccountType](olaccounttype-enumeration-outlook.md)** enumeration that indicates the type of the **[Account](account-object-outlook.md)** . Read-only.|
|[Application](account-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[AutoDiscoverConnectionMode](account-autodiscoverconnectionmode-property-outlook.md)|Returns an  **[OlAutoDiscoverConnectionMode](olautodiscoverconnectionmode-enumeration-outlook.md)** constant that specifies the type of connection to use for the auto-discovery service of the Microsoft Exchange server that hosts the account mailbox. Read-only.|
|[AutoDiscoverXml](account-autodiscoverxml-property-outlook.md)|Returns a  **String** that represents information in XML retrieved from the auto-discovery service of the Microsoft Exchange Server that is associated with the account. Read-only.|
|[Class](account-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[CurrentUser](account-currentuser-property-outlook.md)|Returns a  **[Recipient](recipient-object-outlook.md)** object that represents the current user identity for the account. Read-only.|
|[DeliveryStore](account-deliverystore-property-outlook.md)|Returns a  **[Store](store-object-outlook.md)** object that represents the default delivery store for the account. Read-only.|
|[DisplayName](account-displayname-property-outlook.md)|Returns a  **String** representing the display name of the e-mail **[Account](account-object-outlook.md)** . Read-only.|
|[ExchangeConnectionMode](account-exchangeconnectionmode-property-outlook.md)|Returns an  **[OlExchangeConnectionMode](olexchangeconnectionmode-enumeration-outlook.md)** constant that indicates the current connection mode for the Microsoft Exchange Server that hosts the account mailbox. Read-only|
|[ExchangeMailboxServerName](account-exchangemailboxservername-property-outlook.md)|Returns a  **String** value that represents the name of the Microsoft Exchange Server that hosts the account mailbox. Read-only.|
|[ExchangeMailboxServerVersion](account-exchangemailboxserverversion-property-outlook.md)|Returns a  **String** value that represents the full version number of the Microsoft Exchange Server that hosts the account mailbox. Read-only.|
|[Parent](account-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Session](account-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[SmtpAddress](account-smtpaddress-property-outlook.md)|Returns a  **String** representing the Simple Mail Transfer Protocol (SMTP) address for the **[Account](account-object-outlook.md)** . Read-only.|
|[UserName](account-username-property-outlook.md)|Returns a  **String** representing the user name for the **[Account](account-object-outlook.md)** . Read-only.|

