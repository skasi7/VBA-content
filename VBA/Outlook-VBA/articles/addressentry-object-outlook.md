---
title: AddressEntry Object (Outlook)
keywords: vbaol11.chm2037
f1_keywords:
- vbaol11.chm2037
ms.prod: OUTLOOK
api_name:
- Outlook.AddressEntry
ms.assetid: d4a0a85e-8bab-bc56-57bc-d70c3c570c8e
---


# AddressEntry Object (Outlook)

Represents a person, group, or public folder to which the messaging system can deliver messages.


## Remarks

The  **AddressEntry** object is an address in an **[AddressEntries](addressentries-object-outlook.md)** object. Each **AddressEntry** object in the **AddressEntries** object holds information that represents a person, group, or public folder to which the messaging system can deliver messages.

Use  **AddressEntries** ( _index_ ), where _index_ is the index number of an address entry or a value used to match the default property of an address entry, to return a single **AddressEntry** object.


## Example

The following example sets a reference to an  **AddressEntry** object.


```
Set myAddressEntry = myRecipient.AddressEntry 
 

```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/addressentry-delete-method-outlook%28Office.15%29.aspx)|
|[Details](http://msdn.microsoft.com/library/addressentry-details-method-outlook%28Office.15%29.aspx)|
|[GetContact](http://msdn.microsoft.com/library/addressentry-getcontact-method-outlook%28Office.15%29.aspx)|
|[GetExchangeDistributionList](http://msdn.microsoft.com/library/addressentry-getexchangedistributionlist-method-outlook%28Office.15%29.aspx)|
|[GetExchangeUser](http://msdn.microsoft.com/library/addressentry-getexchangeuser-method-outlook%28Office.15%29.aspx)|
|[GetFreeBusy](http://msdn.microsoft.com/library/addressentry-getfreebusy-method-outlook%28Office.15%29.aspx)|
|[Update](http://msdn.microsoft.com/library/addressentry-update-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/addressentry-address-property-outlook%28Office.15%29.aspx)|
|[AddressEntryUserType](http://msdn.microsoft.com/library/addressentry-addressentryusertype-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/addressentry-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/addressentry-class-property-outlook%28Office.15%29.aspx)|
|[DisplayType](http://msdn.microsoft.com/library/addressentry-displaytype-property-outlook%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/addressentry-id-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/addressentry-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/addressentry-parent-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/addressentry-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/addressentry-session-property-outlook%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/addressentry-type-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[AddressEntry Object Members](http://msdn.microsoft.com/library/addressentry-members-outlook%28Office.15%29.aspx)
