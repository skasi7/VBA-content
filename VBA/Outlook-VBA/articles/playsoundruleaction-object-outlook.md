---
title: PlaySoundRuleAction Object (Outlook)
keywords: vbaol11.chm3169
f1_keywords:
- vbaol11.chm3169
ms.prod: OUTLOOK
api_name:
- Outlook.PlaySoundRuleAction
ms.assetid: 6a7a1f78-640e-8ffc-558c-c26b87638d64
---


# PlaySoundRuleAction Object (Outlook)

Represents an action that plays a .wav file sound.


## Remarks

 **PlaySoundRuleAction** is derived from the **[RuleAction](ruleaction-object-outlook.md)** object. Each rule is associated with a **[RuleActions](ruleactions-object-outlook.md)** object which has a **[PlaySound](ruleactions-playsound-property-outlook.md)** property. The **PlaySound** property always returns a **PlaySoundRuleAction** object. If the rule has an enabled rule action that plays a sound file, then **[PlaySoundRuleAction.Enabled](playsoundruleaction-enabled-property-outlook.md)** property would be **True**.

For more information on specifying rule actions, see [Specify Rule Actions](http://msdn.microsoft.com/library/specifying-rule-actions%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[ActionType](playsoundruleaction-actiontype-property-outlook.md)|
|[Application](playsoundruleaction-application-property-outlook.md)|
|[Class](playsoundruleaction-class-property-outlook.md)|
|[Enabled](playsoundruleaction-enabled-property-outlook.md)|
|[FilePath](playsoundruleaction-filepath-property-outlook.md)|
|[Parent](playsoundruleaction-parent-property-outlook.md)|
|[Session](playsoundruleaction-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
