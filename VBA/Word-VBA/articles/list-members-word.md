---
title: List Members (Word)
ms.prod: WORD
ms.assetid: 939e2533-7d59-bc78-0e89-53e4f204da49
---


# List Members (Word)
Represents a single list format that's been applied to specified paragraphs in a document. The  **List** object is a member of the **Lists** collection.

Represents a single list format that's been applied to specified paragraphs in a document. The  **List** object is a member of the **Lists** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyListTemplate](list-applylisttemplate-method-word.md)|Applies a set of list-formatting characteristics to the specified  **ListFormat** object.|
|[ApplyListTemplateWithLevel](list-applylisttemplatewithlevel-method-word.md)|Applies a set of list-formatting characteristics, optionally for a specified level.|
|[CanContinuePreviousList](list-cancontinuepreviouslist-method-word.md)|Returns a  **[WdContinue](wdcontinue-enumeration-word.md)** constant ( **wdContinueDisabled** , **wdResetList** , or **wdContinueList** ) that indicates whether the formatting from the previous list can be continued.|
|[ConvertNumbersToText](list-convertnumberstotext-method-word.md)|Changes the list numbers and LISTNUM fields in the specified  **List** object.|
|[CountNumberedItems](list-countnumbereditems-method-word.md)|Returns the number of bulleted or numbered items and LISTNUM fields in the specified  **List** object.|
|[RemoveNumbers](list-removenumbers-method-word.md)|Removes numbers or bullets from the specified list.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](list-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](list-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[ListParagraphs](list-listparagraphs-property-word.md)|Returns a  **[ListParagraphs](listparagraphs-object-word.md)** collection that represents all the numbered paragraphs in the list, document, or range. Read-only.|
|[Parent](list-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **List** object.|
|[Range](list-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that is contained in the specified object.|
|[SingleListTemplate](list-singlelisttemplate-property-word.md)| **True** if the entire list uses the same list template. Read-only **Boolean** .|
|[StyleName](list-stylename-property-word.md)|Returns the name of the style applied to the specified AutoText entry. Read-only  **String** .|

