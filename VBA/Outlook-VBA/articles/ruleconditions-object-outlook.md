---
title: RuleConditions Object (Outlook)
keywords: vbaol11.chm3172
f1_keywords:
- vbaol11.chm3172
ms.prod: OUTLOOK
api_name:
- Outlook.RuleConditions
ms.assetid: e8e9a05a-b36b-add2-b294-8cdc5a97e119
---


# RuleConditions Object (Outlook)

Contains a set of  **[RuleCondition](http://msdn.microsoft.com/library/rulecondition-object-outlook%28Office.15%29.aspx)** objects or objects derived from **RuleCondition**, representing the conditions or exception conditions that must be satisfied in order for the **[Rule](rule-object-outlook.md)** to execute.


## Remarks

The  **RuleConditions** object include both rule conditions and rule exceptions. The type of rule condition that can be added to a **RuleConditions** collection depends upon the **[Rule.RuleType](http://msdn.microsoft.com/library/rule-ruletype-property-outlook%28Office.15%29.aspx)**.

The  **RuleConditions** object is a fixed collection. A **RuleCondition** object or a type that is derived from the **RuleCondition** object cannot be added or removed from the **RuleConditions** object.

The Rules object model provides partial parity with the Rules and Alerts Wizard in the Outlook user interface. It supports the most commonly used rule actions and conditions. Although it does not support creating rules with any rule action or rule condition that the Wizard supports, you can still enumerate and enable these rule actions and conditions in existing rules. 

For more information on rule conditions, see [Specifying Rule Conditions](http://msdn.microsoft.com/library/specifying-rule-conditions%28Office.15%29.aspx) and[How to: Create a Rule to Move Specific E-mails to a Folder](http://msdn.microsoft.com/library/create-a-rule-to-move-specific-e-mails-to-a-folder%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[Item](http://msdn.microsoft.com/library/ruleconditions-item-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Account](http://msdn.microsoft.com/library/ruleconditions-account-property-outlook%28Office.15%29.aspx)|
|[AnyCategory](http://msdn.microsoft.com/library/ruleconditions-anycategory-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/ruleconditions-application-property-outlook%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/ruleconditions-body-property-outlook%28Office.15%29.aspx)|
|[BodyOrSubject](http://msdn.microsoft.com/library/ruleconditions-bodyorsubject-property-outlook%28Office.15%29.aspx)|
|[Category](http://msdn.microsoft.com/library/ruleconditions-category-property-outlook%28Office.15%29.aspx)|
|[CC](http://msdn.microsoft.com/library/ruleconditions-cc-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/ruleconditions-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/ruleconditions-count-property-outlook%28Office.15%29.aspx)|
|[FormName](http://msdn.microsoft.com/library/ruleconditions-formname-property-outlook%28Office.15%29.aspx)|
|[From](http://msdn.microsoft.com/library/ruleconditions-from-property-outlook%28Office.15%29.aspx)|
|[FromAnyRSSFeed](http://msdn.microsoft.com/library/ruleconditions-fromanyrssfeed-property-outlook%28Office.15%29.aspx)|
|[FromRssFeed](http://msdn.microsoft.com/library/ruleconditions-fromrssfeed-property-outlook%28Office.15%29.aspx)|
|[HasAttachment](http://msdn.microsoft.com/library/ruleconditions-hasattachment-property-outlook%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/ruleconditions-importance-property-outlook%28Office.15%29.aspx)|
|[MeetingInviteOrUpdate](http://msdn.microsoft.com/library/ruleconditions-meetinginviteorupdate-property-outlook%28Office.15%29.aspx)|
|[MessageHeader](http://msdn.microsoft.com/library/ruleconditions-messageheader-property-outlook%28Office.15%29.aspx)|
|[NotTo](http://msdn.microsoft.com/library/ruleconditions-notto-property-outlook%28Office.15%29.aspx)|
|[OnLocalMachine](http://msdn.microsoft.com/library/ruleconditions-onlocalmachine-property-outlook%28Office.15%29.aspx)|
|[OnlyToMe](http://msdn.microsoft.com/library/ruleconditions-onlytome-property-outlook%28Office.15%29.aspx)|
|[OnOtherMachine](http://msdn.microsoft.com/library/ruleconditions-onothermachine-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/ruleconditions-parent-property-outlook%28Office.15%29.aspx)|
|[RecipientAddress](http://msdn.microsoft.com/library/ruleconditions-recipientaddress-property-outlook%28Office.15%29.aspx)|
|[SenderAddress](http://msdn.microsoft.com/library/ruleconditions-senderaddress-property-outlook%28Office.15%29.aspx)|
|[SenderInAddressList](http://msdn.microsoft.com/library/ruleconditions-senderinaddresslist-property-outlook%28Office.15%29.aspx)|
|[SentTo](http://msdn.microsoft.com/library/ruleconditions-sentto-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/ruleconditions-session-property-outlook%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/ruleconditions-subject-property-outlook%28Office.15%29.aspx)|
|[ToMe](http://msdn.microsoft.com/library/ruleconditions-tome-property-outlook%28Office.15%29.aspx)|
|[ToOrCc](http://msdn.microsoft.com/library/ruleconditions-toorcc-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[RuleConditions Object Members](http://msdn.microsoft.com/library/ruleconditions-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
