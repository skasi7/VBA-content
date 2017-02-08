---
title: ItemProperty Object (Outlook)
keywords: vbaol11.chm517
f1_keywords:
- vbaol11.chm517
ms.prod: OUTLOOK
api_name:
- Outlook.ItemProperty
ms.assetid: 3570d1f9-40ed-0a99-f63c-141134418c3b
---


# ItemProperty Object (Outlook)

Represents information about a given item property for a Microsoft Outlook item object.


## Remarks

 Each item property defines a certain attribute of the item, such as the name, type, or value of the item. The **ItemProperty** object is a member of the **[ItemProperties](itemproperties-object-outlook.md)** collection.

Use  **ItemProperties.Item** ( _index_ ), where _index_ is the object's numeric position within the collection or it's name to return a single **ItemProperty** object.


## Example

The following example creates a reference to the first  **ItemProperty** object in the **ItemProperties** collection.


```
Sub NewMail() 
 
 'Creates a new MailItem and references the ItemProperties collection. 
 
 Dim objMail As MailItem 
 
 Dim objitems As ItemProperties 
 
 Dim objitem As ItemProperty 
 
 
 
 'Create a new mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 'Create a reference to the ItemProperties collection 
 
 Set objitems = objMail.ItemProperties 
 
 'Create reference to the first object in the collection 
 
 Set objitem = objitems.item(0) 
 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/itemproperty-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/itemproperty-class-property-outlook%28Office.15%29.aspx)|
|[IsUserProperty](http://msdn.microsoft.com/library/itemproperty-isuserproperty-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/itemproperty-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/itemproperty-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/itemproperty-session-property-outlook%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/itemproperty-type-property-outlook%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/itemproperty-value-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[ItemProperty Object Members](http://msdn.microsoft.com/library/itemproperty-members-outlook%28Office.15%29.aspx)
