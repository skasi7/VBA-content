---
title: AddressList Object (Outlook)
keywords: vbaol11.chm2022
f1_keywords:
- vbaol11.chm2022
ms.prod: OUTLOOK
api_name:
- Outlook.AddressList
ms.assetid: 84611afe-48b1-185b-df4b-0f004e7436ff
---


# AddressList Object (Outlook)

Represents an address book that contains a set of  **[AddressEntry](addressentry-object-outlook.md)** objects.


## Remarks

The  **AddressList** object is an address book that contains a set of **[AddressEntry](addressentry-object-outlook.md)** objects.

The  **AddressList** object supplies a list of address entries to which a messaging system can deliver messages. An **AddressList** object represents one address book container available under the transport provider's address book hierarchy for the current session. The entire hierarchy is available through the parent **[AddressLists](http://msdn.microsoft.com/library/addresslists-object-outlook%28Office.15%29.aspx)** object.


## Example

The following example retrieves an  **AddressList** object that represents the Personal Address List.


```
Set myAddressList = Application.Session.AddressLists("Personal Address Book")
```


## Methods



|**Name**|
|:-----|
|[GetContactsFolder](http://msdn.microsoft.com/library/addresslist-getcontactsfolder-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddressEntries](http://msdn.microsoft.com/library/addresslist-addressentries-property-outlook%28Office.15%29.aspx)|
|[AddressListType](http://msdn.microsoft.com/library/addresslist-addresslisttype-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/addresslist-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/addresslist-class-property-outlook%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/addresslist-id-property-outlook%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/addresslist-index-property-outlook%28Office.15%29.aspx)|
|[IsInitialAddressList](http://msdn.microsoft.com/library/addresslist-isinitialaddresslist-property-outlook%28Office.15%29.aspx)|
|[IsReadOnly](http://msdn.microsoft.com/library/addresslist-isreadonly-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/addresslist-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/addresslist-parent-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/addresslist-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[ResolutionOrder](http://msdn.microsoft.com/library/addresslist-resolutionorder-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/addresslist-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[AddressList Object Members](http://msdn.microsoft.com/library/addresslist-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
