---
title: Characters Members (Visio)
ms.prod: VISIO
ms.assetid: 16505f00-ddfd-a9c9-cf6d-4058b0642a0e
---


# Characters Members (Visio)
Represents a shape's text with the text fields expanded to the number of characters they display in a drawing window.

Represents a shape's text with the text fields expanded to the number of characters they display in a drawing window.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[TextChanged](characters-textchanged-event-visio.md)|Occurs after the text of a shape is changed in a document.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddCustomField](characters-addcustomfield-method-visio.md)|Replaces the text represented by a  **Characters** object with a custom formula field that uses universal syntax.|
|[AddCustomFieldU](characters-addcustomfieldu-method-visio.md)|Replaces the text represented by a  **Characters** object with a custom formula field that uses universal syntax.|
|[AddField](characters-addfield-method-visio.md)|Replaces the text represented by a  **Characters** object with a new field of the category, code, and format you specify.|
|[AddFieldEx](characters-addfieldex-method-visio.md)|Replaces the text represented by a  **Characters** object with a new field of the category, code, format, language ID, and calendar ID you specify.|
|[Copy](characters-copy-method-visio.md)|Copies a text range to the Clipboard.|
|[Cut](characters-cut-method-visio.md)|Deletes a text range and places it on the Clipboard.|
|[Delete](characters-delete-method-visio.md)|Deletes an object or selection.|
|[Paste](characters-paste-method-visio.md)|Pastes the text range on the Clipboard into an object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](characters-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[Begin](characters-begin-property-visio.md)|Gets or sets the beginning index of a  **Characters** object, which represents a range of text in a shape. Read/write.|
|[CharCount](characters-charcount-property-visio.md)|Returns the number of characters in an object. Read-only.|
|[CharProps](characters-charprops-property-visio.md)|Sets a character property of a  **Characters** object to a new value. Write-only.|
|[CharPropsRow](characters-charpropsrow-property-visio.md)|Returns the index of the row in the Character section of a ShapeSheet window that contains character formatting information for a  **Characters** object. Read-only.|
|[ContainingMasterID](characters-containingmasterid-property-visio.md)|Returns the ID of the  **Master** object that contains an object. Read-only.|
|[ContainingPageID](characters-containingpageid-property-visio.md)|Returns the ID of the page that contains an object. Read-only.|
|[Document](characters-document-property-visio.md)|Gets the  **Document** object that is associated with an object. Read-only.|
|[End](characters-end-property-visio.md)|Returns or sets the ending index of the indicated  **Characters** object representing a range of text in a shape. Read/write.|
|[EventList](characters-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[FieldCategory](characters-fieldcategory-property-visio.md)|Returns the field category for a field represented by an object. Read-only.|
|[FieldCode](characters-fieldcode-property-visio.md)|Returns the field code for a field represented by an object. Read-only.|
|[FieldFormat](characters-fieldformat-property-visio.md)|Returns the field format for a field represented by an object. Read-only.|
|[FieldFormula](characters-fieldformula-property-visio.md)|Returns the formula of the custom field represented by an object. Read-only.|
|[FieldFormulaU](characters-fieldformulau-property-visio.md)|Returns the universal-syntax formula of the custom field represented by an object. Read-only.|
|[IsField](characters-isfield-property-visio.md)|Determines whether a  **Characters** object represents the expanded text of a single field with no additional non-field characters. Read-only.|
|[ObjectType](characters-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[ParaProps](characters-paraprops-property-visio.md)|Sets the paragraph property of a  **Characters** object to a new value. Read/write.|
|[ParaPropsRow](characters-parapropsrow-property-visio.md)|Returns the index of the row in the Paragraph section of a ShapeSheet window that contains paragraph-formatting information for a  **Characters** object. Read-only.|
|[PersistsEvents](characters-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[RunBegin](characters-runbegin-property-visio.md)|Returns the beginning index of a type of run?a sequence of characters that share a particular attribute, such as character, paragraph, or tab formatting; or a word, paragraph, or field. Read-only.|
|[RunEnd](characters-runend-property-visio.md)|Returns the ending index of a type of run?a sequence of characters that share a particular attribute, such as character, paragraph, or tab formatting; or a word, paragraph, or field. Read-only.|
|[Shape](characters-shape-property-visio.md)|Returns the  **Shape** object that owns a **Cell** , **Characters** , **Row** , or **Section** object or that is associated with a **Hyperlink** or **OLEObject** object or with the **Hyperlinks** collection. Read-only.|
|[Stat](characters-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[TabPropsRow](characters-tabpropsrow-property-visio.md)|Returns the index of the row in the Tabs section of the ShapeSheet that contains tab formatting information for a  **Characters** object. Read-only.|
|[Text](characters-text-property-visio.md)|Returns the range of text represented by a  **Characters** object, which may be a subset of the shape's text depending on the values of the **Characters** object's **Begin** and **End** properties.Read/write.|

