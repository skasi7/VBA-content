---
title: UserProperty Object (Outlook)
keywords: vbaol11.chm212
f1_keywords:
- vbaol11.chm212
ms.prod: OUTLOOK
api_name:
- Outlook.UserProperty
ms.assetid: c94f642f-4368-d775-a79f-ce6c39bfe1fd
---


# UserProperty Object (Outlook)

Represents a custom property of an Outlook item.


## Remarks

Use  **[UserProperties](http://msdn.microsoft.com/library/mailitem-userproperties-property-outlook%28Office.15%29.aspx)** ( _index_ ), where _index_ is a name or index number, to return a single **UserProperty** object.

Use the  **[Add](http://msdn.microsoft.com/library/userproperties-add-method-outlook%28Office.15%29.aspx)** method to create a new **UserProperty** for an item and add it to the **[UserProperties](userproperties-object-outlook.md)** object. The **Add** method allows you to specify a name and type for the new property.




 **Note**  When you create a custom property, a field is added in the folder that contains the item (using the same name as the property). That field can be used as a column in folder views.


## Example

The following example adds a custom text property named MyPropName.


```
Set myProp = myItem.UserProperties.Add("MyPropName", olText)
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/userproperty-delete-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/userproperty-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/userproperty-class-property-outlook%28Office.15%29.aspx)|
|[Formula](http://msdn.microsoft.com/library/userproperty-formula-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/userproperty-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/userproperty-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/userproperty-session-property-outlook%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/userproperty-type-property-outlook%28Office.15%29.aspx)|
|[ValidationFormula](http://msdn.microsoft.com/library/userproperty-validationformula-property-outlook%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/userproperty-validationtext-property-outlook%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/userproperty-value-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[UserProperty Object Members](http://msdn.microsoft.com/library/userproperty-members-outlook%28Office.15%29.aspx)
