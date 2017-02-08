---
title: ExchangeUser Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: b9489e9d-0b8e-1c8d-d5df-8def4b1ee5e8
---


# ExchangeUser Members (Outlook)
Provides detailed information about an  **[AddressEntry](addressentry-object-outlook.md)** that represents a Microsoft Exchange mailbox user.

Provides detailed information about an  **[AddressEntry](addressentry-object-outlook.md)** that represents a Microsoft Exchange mailbox user.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](exchangeuser-delete-method-outlook.md)|Deletes the  **[ExchangeUser](exchangeuser-object-outlook.md)** object from the **[AddressEntries](addressentries-object-outlook.md)** collection object to which it belongs.|
|[Details](exchangeuser-details-method-outlook.md)|Displays a modal dialog box that provides detailed information about an  **[ExchangeUser](exchangeuser-object-outlook.md)** object.|
|[GetContact](exchangeuser-getcontact-method-outlook.md)|Returns  **Null** ( **Nothing** in Visual Basic) because the **[ExchangeUser](exchangeuser-object-outlook.md)** object does not correspond to a contact in a Contacts Address Book.|
|[GetDirectReports](exchangeuser-getdirectreports-method-outlook.md)|Obtains an  **[AddressEntries](addressentries-object-outlook.md)** collection object that contains all the users directly reporting to the Exchange user.|
|[GetExchangeDistributionList](exchangeuser-getexchangedistributionlist-method-outlook.md)|Returns  **Null** ( **Nothing** in Visual Basic) because the **[ExchangeUser](exchangeuser-object-outlook.md)** object does not correspond to an **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** object.|
|[GetExchangeUser](exchangeuser-getexchangeuser-method-outlook.md)|Returns the  **[ExchangeUser](exchangeuser-object-outlook.md)** object.|
|[GetExchangeUserManager](exchangeuser-getexchangeusermanager-method-outlook.md)|Returns an  **[ExchangeUser](exchangeuser-object-outlook.md)** object that represents the manager of the Exchange user.|
|[GetFreeBusy](exchangeuser-getfreebusy-method-outlook.md)|Obtains a  **String** representing the availability of the **[ExchangeUser](exchangeuser-object-outlook.md)** for a period of 30 days from the start date, beginning at midnight of the date specified.|
|[GetMemberOfList](exchangeuser-getmemberoflist-method-outlook.md)|Returns an  **[AddressEntries](addressentries-object-outlook.md)** collection object that contains the **[AddressEntry](addressentry-object-outlook.md)** objects representing all the Exchange distribution lists to which the user belongs.|
|[GetPicture](exchangeuser-getpicture-method-outlook.md)|Obtains an  **[IPictureDisp](http://msdn.microsoft.com/en-us/library/ms680762%28VS.85%29.aspx)** object that represents the picture of the Microsoft Exchange user that is displayed in Microsoft Outlook.|
|[Update](exchangeuser-update-method-outlook.md)|Posts a change to the  **[ExchangeUser](exchangeuser-object-outlook.md)** object in the messaging system.|
|[GetUnifiedGroup](exchangeuser-getunifiedgroup-method-outlook.md)|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](exchangeuser-isunifiedgroup-method-outlook.md).|
|[GetUnifiedGroupFromStore](exchangeuser-getunifiedgroupfromstore-method-outlook.md)|Determines if the object is a unified group, by way of a call to [IsUnifiedGroup](exchangeuser-isunifiedgroup-method-outlook.md).|
|[IsUnifiedGroup](exchangeuser-isunifiedgroup-method-outlook.md)|Determines if the object is a unified group.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Address](exchangeuser-address-property-outlook.md)|Returns or sets a  **String** representing the X400 e-mail address of the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[AddressEntryUserType](exchangeuser-addressentryusertype-property-outlook.md)|Returns  **olExchangeUserAddressEntry** which is a constant from the **[OlAddressEntryUserType](oladdressentryusertype-enumeration-outlook.md)** enumeration representing the user type of the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read-only.|
|[Alias](exchangeuser-alias-property-outlook.md)|Returns a  **String** representing the alias for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read-only.|
|[Application](exchangeuser-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent application (Outlook) for the **[ExchangeUser](exchangeuser-object-outlook.md)** object. Read-only.|
|[AssistantName](exchangeuser-assistantname-property-outlook.md)|Returns a  **String** representing the name of the assistant for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[BusinessTelephoneNumber](exchangeuser-businesstelephonenumber-property-outlook.md)|Returns a  **String** representing the business telephone number for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[City](exchangeuser-city-property-outlook.md)|Returns a  **String** representing the city for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[Class](exchangeuser-class-property-outlook.md)|Returns a constant in the  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** enumeration indicating the class of the **[ExchangeUser](exchangeuser-object-outlook.md)** object. Read-only.|
|[Comments](exchangeuser-comments-property-outlook.md)|Returns a  **String** representing the comments for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[CompanyName](exchangeuser-companyname-property-outlook.md)|Returns a  **String** representing the name of the company for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[Department](exchangeuser-department-property-outlook.md)|Returns a  **String** representing the department for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[DisplayType](exchangeuser-displaytype-property-outlook.md)|Returns  **olUser** which is a constant from the **[OlDisplayType](oldisplaytype-enumeration-outlook.md)** enumeration representing the nature of the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read-only.|
|[FirstName](exchangeuser-firstname-property-outlook.md)|Returns a  **String** representing the first name of the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[ID](exchangeuser-id-property-outlook.md)|Returns a  **String** representing the unique identifier for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read-only.|
|[JobTitle](exchangeuser-jobtitle-property-outlook.md)|Returns a  **String** representing the job title of the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[LastName](exchangeuser-lastname-property-outlook.md)|Returns a  **String** representing the last name of the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[MobileTelephoneNumber](exchangeuser-mobiletelephonenumber-property-outlook.md)|Returns a  **String** representing the mobile telephone number for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[Name](exchangeuser-name-property-outlook.md)|Returns or sets a  **String** value that represents the display name for the **[ExchangeUser](exchangeuser-object-outlook.md)** object. Read/write.|
|[OfficeLocation](exchangeuser-officelocation-property-outlook.md)|Returns a  **String** representing the office location for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[Parent](exchangeuser-parent-property-outlook.md)|Returns the parent  **Object** of the **[ExchangeUser](exchangeuser-object-outlook.md)** object. Read-only.|
|[PostalCode](exchangeuser-postalcode-property-outlook.md)|Returns a  **String** representing the postal code for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[PrimarySmtpAddress](exchangeuser-primarysmtpaddress-property-outlook.md)|Returns a  **String** representing the primary Simple Mail Transfer Protocol (SMTP) address for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read-only.|
|[PropertyAccessor](exchangeuser-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[ExchangeUser](exchangeuser-object-outlook.md)** object. Read-only.|
|[Session](exchangeuser-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[StateOrProvince](exchangeuser-stateorprovince-property-outlook.md)|Returns a  **String** representing the state or province for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[StreetAddress](exchangeuser-streetaddress-property-outlook.md)|Returns a  **String** representing the street address for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[Type](exchangeuser-type-property-outlook.md)|Returns a  **String** representing the type of entry for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[YomiCompanyName](exchangeuser-yomicompanyname-property-outlook.md)|Returns a  **String** representing the Japanese phonetic rendering (yomigana) of the company name for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[YomiDepartment](exchangeuser-yomidepartment-property-outlook.md)|Returns a  **String** representing the Japanese phonetic rendering (yomigana) of the department name for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[YomiDisplayName](exchangeuser-yomidisplayname-property-outlook.md)|Returns a  **String** representing the Japanese phonetic rendering (yomigana) of the Exchange display name for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[YomiFirstName](exchangeuser-yomifirstname-property-outlook.md)|Returns a  **String** representing the Japanese phonetic rendering (yomigana) of the first name for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|
|[YomiLastName](exchangeuser-yomilastname-property-outlook.md)|Returns a  **String** representing the Japanese phonetic rendering (yomigana) of the last name for the **[ExchangeUser](exchangeuser-object-outlook.md)** . Read/write.|

