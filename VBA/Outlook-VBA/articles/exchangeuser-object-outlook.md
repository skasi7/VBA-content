---
title: ExchangeUser Object (Outlook)
keywords: vbaol11.chm3158
f1_keywords:
- vbaol11.chm3158
ms.prod: OUTLOOK
api_name:
- Outlook.ExchangeUser
ms.assetid: 6ec117d1-7fdb-aa36-b567-1242f8238df0
---


# ExchangeUser Object (Outlook)

Provides detailed information about an  **[AddressEntry](addressentry-object-outlook.md)** that represents a Microsoft Exchange mailbox user.


## Remarks

 **ExchangeUser** is derived from the **AddressEntry** object, and is returned instead of an **AddressEntry** when the caller performs a query interface on the **AddressEntry** object.

This object provides first-class access to properties applicable to Exchange users such as  **[FirstName](http://msdn.microsoft.com/library/exchangeuser-firstname-property-outlook%28Office.15%29.aspx)**, **[JobTitle](http://msdn.microsoft.com/library/exchangeuser-jobtitle-property-outlook%28Office.15%29.aspx)**, **[LastName](http://msdn.microsoft.com/library/exchangeuser-lastname-property-outlook%28Office.15%29.aspx)**, and **[OfficeLocation](http://msdn.microsoft.com/library/exchangeuser-officelocation-property-outlook%28Office.15%29.aspx)**. You can also access other properties specific to the Exchange user that are not exposed in the object model through the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object. Note that some of the explicit built-in properties are read-write properties. Setting these properties requires the code to be running under an appropriate Exchange administrator account; without sufficient permissions, calling the **[ExchangeUser.Update](http://msdn.microsoft.com/library/exchangeuser-update-method-outlook%28Office.15%29.aspx)** method will result in a "permission denied" error.


## Example

The following code sample shows how to obtain the business phone number, office location, and job title for all entries in the Exchange Global Address List.


```
Sub DemoAE() 
 
 Dim colAL As Outlook.AddressLists 
 
 Dim oAL As Outlook.AddressList 
 
 Dim colAE As Outlook.AddressEntries 
 
 Dim oAE As Outlook.AddressEntry 
 
 Dim oExUser As Outlook.ExchangeUser 
 
 Set colAL = Application.Session.AddressLists 
 
 For Each oAL In colAL 
 
 'Address list is an Exchange Global Address List 
 
 If oAL.AddressListType = olExchangeGlobalAddressList Then 
 
 Set colAE = oAL.AddressEntries 
 
 For Each oAE In colAE 
 
 If oAE.AddressEntryUserType = _ 
 
 olExchangeUserAddressEntry Then 
 
 Set oExUser = oAE.GetExchangeUser 
 
 Debug.Print(oExUser.JobTitle) 
 
 Debug.Print(oExUser.OfficeLocation) 
 
 Debug.Print(oExUser.BusinessTelephoneNumber) 
 
 End If 
 
 Next 
 
 End If 
 
 Next 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/exchangeuser-delete-method-outlook%28Office.15%29.aspx)|
|[Details](http://msdn.microsoft.com/library/exchangeuser-details-method-outlook%28Office.15%29.aspx)|
|[GetContact](http://msdn.microsoft.com/library/exchangeuser-getcontact-method-outlook%28Office.15%29.aspx)|
|[GetDirectReports](http://msdn.microsoft.com/library/exchangeuser-getdirectreports-method-outlook%28Office.15%29.aspx)|
|[GetExchangeDistributionList](http://msdn.microsoft.com/library/exchangeuser-getexchangedistributionlist-method-outlook%28Office.15%29.aspx)|
|[GetExchangeUser](http://msdn.microsoft.com/library/exchangeuser-getexchangeuser-method-outlook%28Office.15%29.aspx)|
|[GetExchangeUserManager](http://msdn.microsoft.com/library/exchangeuser-getexchangeusermanager-method-outlook%28Office.15%29.aspx)|
|[GetFreeBusy](http://msdn.microsoft.com/library/exchangeuser-getfreebusy-method-outlook%28Office.15%29.aspx)|
|[GetMemberOfList](http://msdn.microsoft.com/library/exchangeuser-getmemberoflist-method-outlook%28Office.15%29.aspx)|
|[GetPicture](http://msdn.microsoft.com/library/exchangeuser-getpicture-method-outlook%28Office.15%29.aspx)|
|[Update](http://msdn.microsoft.com/library/exchangeuser-update-method-outlook%28Office.15%29.aspx)|
|[GetUnifiedGroup](http://msdn.microsoft.com/library/exchangeuser-getunifiedgroup-method-outlook%28Office.15%29.aspx)|
|[GetUnifiedGroupFromStore](http://msdn.microsoft.com/library/exchangeuser-getunifiedgroupfromstore-method-outlook%28Office.15%29.aspx)|
|[IsUnifiedGroup](http://msdn.microsoft.com/library/exchangeuser-isunifiedgroup-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/exchangeuser-address-property-outlook%28Office.15%29.aspx)|
|[AddressEntryUserType](http://msdn.microsoft.com/library/exchangeuser-addressentryusertype-property-outlook%28Office.15%29.aspx)|
|[Alias](http://msdn.microsoft.com/library/exchangeuser-alias-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/exchangeuser-application-property-outlook%28Office.15%29.aspx)|
|[AssistantName](http://msdn.microsoft.com/library/exchangeuser-assistantname-property-outlook%28Office.15%29.aspx)|
|[BusinessTelephoneNumber](http://msdn.microsoft.com/library/exchangeuser-businesstelephonenumber-property-outlook%28Office.15%29.aspx)|
|[City](http://msdn.microsoft.com/library/exchangeuser-city-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/exchangeuser-class-property-outlook%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/exchangeuser-comments-property-outlook%28Office.15%29.aspx)|
|[CompanyName](http://msdn.microsoft.com/library/exchangeuser-companyname-property-outlook%28Office.15%29.aspx)|
|[Department](http://msdn.microsoft.com/library/exchangeuser-department-property-outlook%28Office.15%29.aspx)|
|[DisplayType](http://msdn.microsoft.com/library/exchangeuser-displaytype-property-outlook%28Office.15%29.aspx)|
|[FirstName](http://msdn.microsoft.com/library/exchangeuser-firstname-property-outlook%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/exchangeuser-id-property-outlook%28Office.15%29.aspx)|
|[JobTitle](http://msdn.microsoft.com/library/exchangeuser-jobtitle-property-outlook%28Office.15%29.aspx)|
|[LastName](http://msdn.microsoft.com/library/exchangeuser-lastname-property-outlook%28Office.15%29.aspx)|
|[MobileTelephoneNumber](http://msdn.microsoft.com/library/exchangeuser-mobiletelephonenumber-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/exchangeuser-name-property-outlook%28Office.15%29.aspx)|
|[OfficeLocation](http://msdn.microsoft.com/library/exchangeuser-officelocation-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/exchangeuser-parent-property-outlook%28Office.15%29.aspx)|
|[PostalCode](http://msdn.microsoft.com/library/exchangeuser-postalcode-property-outlook%28Office.15%29.aspx)|
|[PrimarySmtpAddress](http://msdn.microsoft.com/library/exchangeuser-primarysmtpaddress-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/exchangeuser-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/exchangeuser-session-property-outlook%28Office.15%29.aspx)|
|[StateOrProvince](http://msdn.microsoft.com/library/exchangeuser-stateorprovince-property-outlook%28Office.15%29.aspx)|
|[StreetAddress](http://msdn.microsoft.com/library/exchangeuser-streetaddress-property-outlook%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/exchangeuser-type-property-outlook%28Office.15%29.aspx)|
|[YomiCompanyName](http://msdn.microsoft.com/library/exchangeuser-yomicompanyname-property-outlook%28Office.15%29.aspx)|
|[YomiDepartment](http://msdn.microsoft.com/library/exchangeuser-yomidepartment-property-outlook%28Office.15%29.aspx)|
|[YomiDisplayName](http://msdn.microsoft.com/library/exchangeuser-yomidisplayname-property-outlook%28Office.15%29.aspx)|
|[YomiFirstName](http://msdn.microsoft.com/library/exchangeuser-yomifirstname-property-outlook%28Office.15%29.aspx)|
|[YomiLastName](http://msdn.microsoft.com/library/exchangeuser-yomilastname-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[ExchangeUser Object Members](http://msdn.microsoft.com/library/exchangeuser-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
