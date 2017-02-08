---
title: Action Object (Outlook)
keywords: vbaol11.chm9
f1_keywords:
- vbaol11.chm9
ms.prod: OUTLOOK
api_name:
- Outlook.Action
ms.assetid: 22bd8d4a-9cf4-bd37-011b-8da3dfadf761
---


# Action Object (Outlook)

Represents a specialized action (for example, the voting options response) that can be executed on an Outlook item.


## Remarks

The  **Action** object is a member of the **[Actions](actions-object-outlook.md)** collection.

Use  **[Actions](http://msdn.microsoft.com/library/mailitem-actions-property-outlook%28Office.15%29.aspx)** ( _index_ ), where _index_ is the name of an available action, to return a single **Action** object from the **Actions** collection object of an Outlook item, such as **[MailItem](http://msdn.microsoft.com/library/mailitem-object-outlook%28Office.15%29.aspx)**.


## Example

The following Visual Basic for Applications (VBA) example uses the Reply action of a particular item to send a reply.


```
myItem = CreateItem(olMailItem) 
 
Set myReply = myItem.Actions("Reply").Execute
```

The following Visual Basic for Applications example does the same thing, using a different reply style for the reply.




```
myItem = CreateItem(olMailItem) 
 
myItem.Actions("Reply").ReplyStyle = _ 
 
 olIncludeOriginalText 
 
Set myReply = myItem.Actions("Reply").Execute
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/action-delete-method-outlook%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/action-execute-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/action-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/action-class-property-outlook%28Office.15%29.aspx)|
|[CopyLike](http://msdn.microsoft.com/library/action-copylike-property-outlook%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/action-enabled-property-outlook%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/action-messageclass-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/action-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/action-parent-property-outlook%28Office.15%29.aspx)|
|[Prefix](http://msdn.microsoft.com/library/action-prefix-property-outlook%28Office.15%29.aspx)|
|[ReplyStyle](http://msdn.microsoft.com/library/action-replystyle-property-outlook%28Office.15%29.aspx)|
|[ResponseStyle](http://msdn.microsoft.com/library/action-responsestyle-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/action-session-property-outlook%28Office.15%29.aspx)|
|[ShowOn](http://msdn.microsoft.com/library/action-showon-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Action Object Members](http://msdn.microsoft.com/library/action-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
