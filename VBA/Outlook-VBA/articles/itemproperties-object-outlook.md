---
title: ItemProperties Object (Outlook)
keywords: vbaol11.chm530
f1_keywords:
- vbaol11.chm530
ms.prod: OUTLOOK
api_name:
- Outlook.ItemProperties
ms.assetid: 34a110ed-6617-72da-1e98-a9773c705b40
---


# ItemProperties Object (Outlook)

A collection of all properties associated with the item.


## Remarks

Use the  **[ItemProperties](http://msdn.microsoft.com/library/mailitem-itemproperties-property-outlook%28Office.15%29.aspx)** property to return the **ItemProperties** collection. Use **ItemProperties.Item** ( _index_ ), where _index_ is the name of the object or the numeric position of the item within the collection, to return a single **[ItemProperty](itemproperty-object-outlook.md)** object.


 **Note**  The  **ItemProperties** collection is zero-based, meaning that the first item in the collection is referenced by 0.

Use the  **[Add](http://msdn.microsoft.com/library/itemproperties-add-method-outlook%28Office.15%29.aspx)** method to add a new item property to the **ItemProperties** collection. Use the **[Remove](http://msdn.microsoft.com/library/itemproperties-remove-method-outlook%28Office.15%29.aspx)** method to remove an item property from the **ItemProperties** collection.


 **Note**   You can only add or remove custom properties. Custom properties are denoted by the **[IsUserProperty](http://msdn.microsoft.com/library/itemproperty-isuserproperty-property-outlook%28Office.15%29.aspx)**.


## Example

The following example creates a new  **[MailItem](http://msdn.microsoft.com/library/mailitem-object-outlook%28Office.15%29.aspx)** object and stores its **ItemProperties** collection in a variable called `objItems`.


```
Sub ItemProperty() 
 
 'Creates a new MailItem and access its properties 
 
 Dim objMail As MailItem 
 
 Dim objItems As ItemProperties 
 
 Dim objItem As ItemProperty 
 
 
 
 'Create the mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 'Create a reference to the item properties collection 
 
 Set objItems = objMail.ItemProperties 
 
 'Create a reference to the item property page 
 
 Set objItem = objItems.item(0) 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/itemproperties-add-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/itemproperties-item-method-outlook%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/itemproperties-remove-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/itemproperties-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/itemproperties-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/itemproperties-count-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/itemproperties-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/itemproperties-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[ItemProperties Object Members](http://msdn.microsoft.com/library/itemproperties-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
