---
title: ListFormat Members (Word)
ms.prod: WORD
ms.assetid: daf87b14-29a3-c5d9-ab43-8465237c02da
---


# ListFormat Members (Word)
Represents the list formatting attributes that can be applied to the paragraphs in a range.

Represents the list formatting attributes that can be applied to the paragraphs in a range.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyBulletDefault](listformat-applybulletdefault-method-word.md)|Adds bullets and formatting to the paragraphs in the range for the specified  **ListFormat** object.|
|[ApplyListTemplate](listformat-applylisttemplate-method-word.md)|Applies a set of list-formatting characteristics to the specified  **ListFormat** object.|
|[ApplyListTemplateWithLevel](listformat-applylisttemplatewithlevel-method-word.md)|Applies a set of list-formatting characteristics, optionally for a specified level.|
|[ApplyNumberDefault](listformat-applynumberdefault-method-word.md)|Adds the default numbering scheme to the paragraphs in the range for the specified  **ListFormat** object.|
|[ApplyOutlineNumberDefault](listformat-applyoutlinenumberdefault-method-word.md)|Adds the default outline-numbering scheme to the paragraphs in the range for the specified  **ListFormat** object.|
|[CanContinuePreviousList](listformat-cancontinuepreviouslist-method-word.md)|Returns a  **WdContinue** constant ( **wdContinueDisabled** , **wdResetList** , or **wdContinueList** ) that indicates whether the formatting from the previous list can be continued.|
|[ConvertNumbersToText](listformat-convertnumberstotext-method-word.md)|Changes the list numbers and LISTNUM fields in the specified  **ListFormat** object to text.|
|[CountNumberedItems](listformat-countnumbereditems-method-word.md)|Returns the number of bulleted or numbered items and LISTNUM fields in the specified  **ListFormat** object.|
|[ListIndent](listformat-listindent-method-word.md)|Increases the list level of the paragraphs in the range for the specified  **ListFormat** object, in increments of one level.|
|[ListOutdent](listformat-listoutdent-method-word.md)|Decreases the list level of the paragraphs in the range for the specified  **ListFormat** object, in increments of one level.|
|[RemoveNumbers](listformat-removenumbers-method-word.md)|Removes numbers or bullets from the specified list.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](listformat-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](listformat-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[List](listformat-list-property-word.md)|Returns a  **[List](list-object-word.md)** object that represents the first formatted list contained in the specified **ListFormat** object.|
|[ListLevelNumber](listformat-listlevelnumber-property-word.md)|Returns or sets the list level for the first paragraph in the specified  **ListFormat** object. Read/write **Long** .|
|[ListPictureBullet](listformat-listpicturebullet-property-word.md)|Returns the  **[InlineShape](inlineshape-object-word.md)** object that represents the picture used as a bullet in a picture bulleted list.|
|[ListString](listformat-liststring-property-word.md)|Returns a  **String** that represents the appearance of the list value of the first paragraph in the range for the specified **ListFormat** object. For example, the second paragraph in an alphabetical list would return B. Read-only.|
|[ListTemplate](listformat-listtemplate-property-word.md)|Returns a  **ListTemplate** object that represents the list formatting for the specified **ListFormat** object.|
|[ListType](listformat-listtype-property-word.md)|Returns the type of lists that are contained in the range for the specified  **ListFormat** object. Read-only **WdListType** .|
|[ListValue](listformat-listvalue-property-word.md)|Returns the numeric value of the first paragraph in the range for the specified  **ListFormat** object. Read-only **Long** .|
|[Parent](listformat-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **ListFormat** object.|
|[SingleList](listformat-singlelist-property-word.md)| **True** if the specified **ListFormat** object contains only one list. Read-only **Boolean** .|
|[SingleListTemplate](listformat-singlelisttemplate-property-word.md)| **True** if the entire **ListFormat** object uses the same list template. Read-only **Boolean** .|

