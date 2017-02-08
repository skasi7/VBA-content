---
title: PropertyAccessor Object (Outlook)
keywords: vbaol11.chm3157
f1_keywords:
- vbaol11.chm3157
ms.prod: OUTLOOK
api_name:
- Outlook.PropertyAccessor
ms.assetid: 2fc91e13-703c-3ec9-9066-ffee7144306c
---


# PropertyAccessor Object (Outlook)

Provides the ability to create, get, set, and delete properties on objects.


## Remarks

Use the  **PropertyAccessor** object to get and set item-level properties that are not explicitly exposed in the Outlook object model, or properties for the following non-item objects: **[AddressEntry](addressentry-object-outlook.md)**, **[AddressList](addresslist-object-outlook.md)**, **[Attachment](http://msdn.microsoft.com/library/attachment-object-outlook%28Office.15%29.aspx)**, **[ExchangeDistributionList](http://msdn.microsoft.com/library/exchangedistributionlist-object-outlook%28Office.15%29.aspx)**, **[ExchangeUser](exchangeuser-object-outlook.md)**, **[Folder](folder-object-outlook.md)**, **[Recipient](recipient-object-outlook.md)**, and **[Store](store-object-outlook.md)**.

To get or set multiple custom properties, use the  **PropertyAccessor** object instead of the **[UserProperties](userproperties-object-outlook.md)** object for better performance.

For more information on using the  **PropertyAccessor** object, see[Properties Overview](http://msdn.microsoft.com/library/properties-overview%28Office.15%29.aspx).


## Example

The following code sample demonstrates how to use the  **[PropertyAccessor.GetProperty](http://msdn.microsoft.com/library/propertyaccessor-getproperty-method-outlook%28Office.15%29.aspx)** method to read a MAPI property that belongs to a **[MailItem](http://msdn.microsoft.com/library/mailitem-object-outlook%28Office.15%29.aspx)** but that is not exposed in the Outlook object model, **PR_TRANSPORT_MESSAGE_HEADERS**.


```
Sub DemoPropertyAccessorGetProperty() 
 
 Dim PropName, Header As String 
 
 Dim oMail As Object 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 'Get first item in the inbox 
 
 Set oMail = _ 
 
 Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 'PR_TRANSPORT_MESSAGE_HEADERS 
 
 PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E" 
 
 'Obtain an instance of PropertyAccessor class 
 
 Set oPA = oMail.PropertyAccessor 
 
 'Call GetProperty 
 
 Header = oPA.GetProperty(PropName) 
 
 Debug.Print (Header) 
 
End Sub
```

The next code sample demonstrates how the  **[PropertyAccessor.SetProperties](http://msdn.microsoft.com/library/propertyaccessor-setproperties-method-outlook%28Office.15%29.aspx)** method sets the values of multiple properties. If a property does not exist, then **SetProperties** will create the property as long as the parent object supports the creation of those properties. If the object supports an explicit **Save** operation, then the properties are saved to the object when the explicit **Save** operation is called. If the object does not support an explicit **Save** operation, then the properties are saved to the object when **SetProperties** is called.




```
Sub DemoPropertyAccessorSetProperties() 
 
 Dim PropNames(), myValues() As Variant 
 
 Dim arrErrors As Variant 
 
 Dim prop1, prop2, prop3, prop4 As String 
 
 Dim i As Integer 
 
 Dim oMail As Outlook.MailItem 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 'Get first item in the inbox 
 
 Set oMail = _ 
 
 Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 'Names for properties using the MAPI string namespace 
 
 prop1 = "http://schemas.microsoft.com/mapi/string/" &amp; _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mylongprop" 
 
 prop2 = "http://schemas.microsoft.com/mapi/string/" &amp; _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mystringprop" 
 
 prop3 = "http://schemas.microsoft.com/mapi/string/" &amp; _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mydateprop" 
 
 prop4 = "http://schemas.microsoft.com/mapi/string/" &amp; _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/myboolprop" 
 
 PropNames = Array(prop1, prop2, prop3, prop4) 
 
 myValues = Array(1020, "111-222-Kudo", Now(), False) 
 
 'Set values with SetProperties call 
 
 'If the properties do not exist, then SetProperties 
 
 'adds the properties to the object when saved. 
 
 'The type of the property is the type of the element 
 
 'passed in myValues array. 
 
 Set oPA = oMail.PropertyAccessor 
 
 arrErrors = oPA.SetProperties(PropNames, myValues) 
 
 If Not (IsEmpty(arrErrors)) Then 
 
 'Examine the arrErrors array to determine if any 
 
 'elements contain errors 
 
 For i = LBound(arrErrors) To UBound(arrErrors) 
 
 'Examine the type of the element 
 
 If IsError(arrErrors(i)) Then 
 
 Debug.Print (CVErr(arrErrors(i))) 
 
 End If 
 
 Next 
 
 End If 
 
 'Save the item 
 
 oMail.Save 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[BinaryToString](http://msdn.microsoft.com/library/propertyaccessor-binarytostring-method-outlook%28Office.15%29.aspx)|
|[DeleteProperties](http://msdn.microsoft.com/library/propertyaccessor-deleteproperties-method-outlook%28Office.15%29.aspx)|
|[DeleteProperty](http://msdn.microsoft.com/library/propertyaccessor-deleteproperty-method-outlook%28Office.15%29.aspx)|
|[GetProperties](http://msdn.microsoft.com/library/propertyaccessor-getproperties-method-outlook%28Office.15%29.aspx)|
|[GetProperty](http://msdn.microsoft.com/library/propertyaccessor-getproperty-method-outlook%28Office.15%29.aspx)|
|[LocalTimeToUTC](http://msdn.microsoft.com/library/propertyaccessor-localtimetoutc-method-outlook%28Office.15%29.aspx)|
|[SetProperties](http://msdn.microsoft.com/library/propertyaccessor-setproperties-method-outlook%28Office.15%29.aspx)|
|[SetProperty](http://msdn.microsoft.com/library/propertyaccessor-setproperty-method-outlook%28Office.15%29.aspx)|
|[StringToBinary](http://msdn.microsoft.com/library/propertyaccessor-stringtobinary-method-outlook%28Office.15%29.aspx)|
|[UTCToLocalTime](http://msdn.microsoft.com/library/propertyaccessor-utctolocaltime-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/propertyaccessor-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/propertyaccessor-class-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/propertyaccessor-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/propertyaccessor-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[PropertyAccessor Object Members](http://msdn.microsoft.com/library/propertyaccessor-members-outlook%28Office.15%29.aspx)
