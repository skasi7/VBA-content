---
title: RuleActions Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: ea4c7acb-2ce2-ecf9-046f-2eb48d4935bb
---


# RuleActions Members (Outlook)
The  **RuleActions** object contains a set of **[RuleAction](ruleaction-object-outlook.md)** objects or objects derived from **RuleAction** , representing the actions that are executed on a **[Rule](rule-object-outlook.md)** object.

The  **RuleActions** object contains a set of **[RuleAction](ruleaction-object-outlook.md)** objects or objects derived from **RuleAction** , representing the actions that are executed on a **[Rule](rule-object-outlook.md)** object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Item](ruleactions-item-method-outlook.md)|Obtains a  **[RuleAction](ruleaction-object-outlook.md)** object specified by _Index_ which is a numerical index into the **[RuleActions](ruleactions-object-outlook.md)** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](ruleactions-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[AssignToCategory](ruleactions-assigntocategory-property-outlook.md)|Returns an  **[AssignToCategoryRuleAction](assigntocategoryruleaction-object-outlook.md)** object with **[AssignToCategoryRuleAction.ActionType](assigntocategoryruleaction-actiontype-property-outlook.md)** being **olRuleAssignToCategory** . Read-only.|
|[CC](ruleactions-cc-property-outlook.md)|Returns a  **[SendRuleAction](sendruleaction-object-outlook.md)** object with **[SendRuleAction.ActionType](sendruleaction-actiontype-property-outlook.md)** being **olRuleActionCcMessage** . Read-only.|
|[Class](ruleactions-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[ClearCategories](ruleactions-clearcategories-property-outlook.md)|Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with a **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** of **olRuleActionClearCategories** . Read-only.|
|[CopyToFolder](ruleactions-copytofolder-property-outlook.md)|Returns a  **[MoveOrCopyRuleAction](moveorcopyruleaction-object-outlook.md)** object with **[MoveOrCopyRuleAction.ActionType](moveorcopyruleaction-actiontype-property-outlook.md)** being **olRuleActionCopyToFolder** . Read-only.|
|[Count](ruleactions-count-property-outlook.md)|Returns a  **Long** indicating the count of objects in the specified collection. Read-only.|
|[Delete](ruleactions-delete-property-outlook.md)|Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionDelete** . Read-only.|
|[DeletePermanently](ruleactions-deletepermanently-property-outlook.md)|Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionDeletePermanently** . Read-only.|
|[DesktopAlert](ruleactions-desktopalert-property-outlook.md)|Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionDesktopAlert** . Read-only.|
|[Forward](ruleactions-forward-property-outlook.md)|Returns a  **[SendRuleAction](sendruleaction-object-outlook.md)** object with **[SendRuleAction.ActionType](sendruleaction-actiontype-property-outlook.md)** being **olRuleActionForward** . Read-only.|
|[ForwardAsAttachment](ruleactions-forwardasattachment-property-outlook.md)|Returns a  **[SendRuleAction](sendruleaction-object-outlook.md)** object with **[SendRuleAction.ActionType](sendruleaction-actiontype-property-outlook.md)** being **olRuleActionForwardAsAttachment** . Read-only.|
|[MarkAsTask](ruleactions-markastask-property-outlook.md)|Returns a  **[MarkAsTaskRuleAction](markastaskruleaction-object-outlook.md)** object with **[MarkAsTaskRuleAction.ActionType](markastaskruleaction-actiontype-property-outlook.md)** being **olRuleActionMarkAsTask** . Read-only.|
|[MoveToFolder](ruleactions-movetofolder-property-outlook.md)|Returns a  **[MoveOrCopyRuleAction](moveorcopyruleaction-object-outlook.md)** object with **[MoveOrCopyRuleAction.ActionType](moveorcopyruleaction-actiontype-property-outlook.md)** being **olRuleActionMoveToFolder** . Read-only.|
|[NewItemAlert](ruleactions-newitemalert-property-outlook.md)|Returns a  **[NewItemAlertRuleAction](newitemalertruleaction-object-outlook.md)** object with **[ActionType](newitemalertruleaction-actiontype-property-outlook.md)** being **olRuleActionNewItemAlert** . Read-only.|
|[NotifyDelivery](ruleactions-notifydelivery-property-outlook.md)|Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionNotifyDelivery** . Read-only.|
|[NotifyRead](ruleactions-notifyread-property-outlook.md)|Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionNotifyRead** . Read-only.|
|[Parent](ruleactions-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PlaySound](ruleactions-playsound-property-outlook.md)|Returns a  **[PlaySoundRuleAction](playsoundruleaction-object-outlook.md)** object with **[PlaySoundRuleAction.ActionType](playsoundruleaction-actiontype-property-outlook.md)** being **olRuleActionNotifyRead** . Read-only.|
|[Redirect](ruleactions-redirect-property-outlook.md)|Returns a  **[SendRuleAction](sendruleaction-object-outlook.md)** object with **[SendRuleAction.ActionType](sendruleaction-actiontype-property-outlook.md)** being **olRuleActionRedirect** . Read-only.|
|[Session](ruleactions-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Stop](ruleactions-stop-property-outlook.md)|Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionStop** . Read-only.|

