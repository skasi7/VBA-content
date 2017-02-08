---
title: UserProperties Object (Outlook)
keywords: vbaol11.chm202
f1_keywords:
- vbaol11.chm202
ms.prod: OUTLOOK
api_name:
- Outlook.UserProperties
ms.assetid: 20b49c86-d74f-9bda-382c-559af278c148
---


# UserProperties Object (Outlook)

Contains  **[UserProperty](userproperty-object-outlook.md)** objects that represent the custom properties of an Outlook item.


## Remarks

Use the  **UserProperties** property to return the **UserProperties** object for an Outlook item. This applies to all Outlook items except for the **[NoteItem](http://msdn.microsoft.com/library/noteitem-object-outlook%28Office.15%29.aspx)**.

Use the  **[Add](http://msdn.microsoft.com/library/userproperties-add-method-outlook%28Office.15%29.aspx)** method to create a new **UserProperty** for an item and add it to the **UserProperties** object. The **Add** method allows you to specify a name and type for the new property. When you create a new property, it can also be added as a custom field to the folder that contains the item (using the same name as the property) by setting the _AddToFolderFields_ parameter to **True** when calling the **Add** method. That field can then be used as a column in folder views.

Use  **UserProperties** ( _index_ ), where _index_ is a name or one-based index number, to return a single **[UserProperty](userproperty-object-outlook.md)** object.

You can use the  **[UserDefinedProperties](http://msdn.microsoft.com/library/folder-userdefinedproperties-property-outlook%28Office.15%29.aspx)** property of the **[Folder](folder-object-outlook.md)** object to retrieve and examine the definitions of custom item-level properties that a folder can display in a view.

To get or set multiple custom properties, use the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object instead of the **UserProperties** object for better performance.


## Example

The following example adds a custom text property named MyPropName to myItem.


```
Set myProp = myItem.UserProperties.Add("MyPropName", olText)
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/userproperties-add-method-outlook%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/userproperties-find-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/userproperties-item-method-outlook%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/userproperties-remove-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/userproperties-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/userproperties-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/userproperties-count-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/userproperties-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/userproperties-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[UserProperties Object Members](http://msdn.microsoft.com/library/userproperties-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
