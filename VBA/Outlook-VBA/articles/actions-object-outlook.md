---
title: Actions Object (Outlook)
keywords: vbaol11.chm144
f1_keywords:
- vbaol11.chm144
ms.prod: OUTLOOK
api_name:
- Outlook.Actions
ms.assetid: b0903aa4-9b75-5311-d0a5-5ff4a5e29c79
---


# Actions Object (Outlook)

Contains a collection of  **[Action](action-object-outlook.md)** objects that represent all the specialized actions that can be executed on an Outlook item.


## Remarks

Use the  **Actions** property of any Outlook item, such as **[MailItem](http://msdn.microsoft.com/library/mailitem-object-outlook%28Office.15%29.aspx)**, to return the **Actions** object.

Use  **Actions** ( _index_ ), where _index_ is the name of an available action, to return a single **Action** object.


## Example

The following Visual Basic for Applications (VBA) example uses the Reply action of a particular item to send a reply.


```
myItem = CreateItem(olMailItem) 
 
Set myReply = myItem.Actions("Reply").Execute
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/actions-add-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/actions-item-method-outlook%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/actions-remove-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/actions-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/actions-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/actions-count-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/actions-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/actions-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[Actions Object Members](http://msdn.microsoft.com/library/actions-members-outlook%28Office.15%29.aspx)
