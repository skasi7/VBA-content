---
title: AddressEntries Object (Outlook)
keywords: vbaol11.chm24
f1_keywords:
- vbaol11.chm24
ms.prod: OUTLOOK
api_name:
- Outlook.AddressEntries
ms.assetid: db91b717-07c6-d1f2-c545-b766ee1f0c6b
---


# AddressEntries Object (Outlook)

Contains a collection of addresses for an  **[AddressList](addresslist-object-outlook.md)** object.


## Remarks

The object may contain zero or more  **[AddressEntry](addressentry-object-outlook.md)** objects and provides access to the entries in a transport provider's address book container.


## Example

The following example sets a reference to an  **AddressEntries** object.






```
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myAddressList = myNameSpace.AddressLists("Personal Address Book") 
 
Set myAddressEntries = myAddressList.AddressEntries
```

You can also index directly into the  **AddressEntries** object, returning an **AddressEntry** object.




```
Set myAddressEntry = myAddressList.AddressEntries(index)
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/addressentries-add-method-outlook%28Office.15%29.aspx)|
|[GetFirst](http://msdn.microsoft.com/library/addressentries-getfirst-method-outlook%28Office.15%29.aspx)|
|[GetLast](http://msdn.microsoft.com/library/addressentries-getlast-method-outlook%28Office.15%29.aspx)|
|[GetNext](http://msdn.microsoft.com/library/addressentries-getnext-method-outlook%28Office.15%29.aspx)|
|[GetPrevious](http://msdn.microsoft.com/library/addressentries-getprevious-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/addressentries-item-method-outlook%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/addressentries-sort-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/addressentries-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/addressentries-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/addressentries-count-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/addressentries-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/addressentries-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[AddressEntries Object Members](http://msdn.microsoft.com/library/addressentries-members-outlook%28Office.15%29.aspx)
