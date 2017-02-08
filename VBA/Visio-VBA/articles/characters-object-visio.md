---
title: Characters Object (Visio)
keywords: vis_sdr.chm10050
f1_keywords:
- vis_sdr.chm10050
ms.prod: VISIO
api_name:
- Visio.Characters
ms.assetid: aaff009b-c665-c2ea-8494-e917126d8491
---


# Characters Object (Visio)

Represents a shape's text with the text fields expanded to the number of characters they display in a drawing window.


## Remarks

To retrieve a  **Characters** object, use the **Characters** property of a **Shape** object.

The default property of a  **Characters** object is **Text**.

The  **Begin** and **End** properties of a **Characters** object determine the range of the shape's text that is represented by the **Characters** object. Initially, the range contains all of the shape's text; you can set the **Begin** and **End** properties to specify a subrange of the text.

After you retrieve a  **Characters** object, you can use its **Text** property to retrieve or set the shape's text. Use the **Copy**, **Cut**, or **Paste** method to copy, cut, or paste the **Characters** object's text to or from the Clipboard. Use the **CharProps** or **ParaProps** property to change the **Characters** object's formatting.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this object maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVCharacters**
    

## Events



|**Name**|
|:-----|
|[TextChanged](http://msdn.microsoft.com/library/characters-textchanged-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddCustomField](http://msdn.microsoft.com/library/characters-addcustomfield-method-visio%28Office.15%29.aspx)|
|[AddCustomFieldU](http://msdn.microsoft.com/library/characters-addcustomfieldu-method-visio%28Office.15%29.aspx)|
|[AddField](http://msdn.microsoft.com/library/characters-addfield-method-visio%28Office.15%29.aspx)|
|[AddFieldEx](http://msdn.microsoft.com/library/characters-addfieldex-method-visio%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/characters-copy-method-visio%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/characters-cut-method-visio%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/characters-delete-method-visio%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/characters-paste-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/characters-application-property-visio%28Office.15%29.aspx)|
|[Begin](http://msdn.microsoft.com/library/characters-begin-property-visio%28Office.15%29.aspx)|
|[CharCount](http://msdn.microsoft.com/library/characters-charcount-property-visio%28Office.15%29.aspx)|
|[CharProps](http://msdn.microsoft.com/library/characters-charprops-property-visio%28Office.15%29.aspx)|
|[CharPropsRow](http://msdn.microsoft.com/library/characters-charpropsrow-property-visio%28Office.15%29.aspx)|
|[ContainingMasterID](http://msdn.microsoft.com/library/characters-containingmasterid-property-visio%28Office.15%29.aspx)|
|[ContainingPageID](http://msdn.microsoft.com/library/characters-containingpageid-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/characters-document-property-visio%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/characters-end-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/characters-eventlist-property-visio%28Office.15%29.aspx)|
|[FieldCategory](http://msdn.microsoft.com/library/characters-fieldcategory-property-visio%28Office.15%29.aspx)|
|[FieldCode](http://msdn.microsoft.com/library/characters-fieldcode-property-visio%28Office.15%29.aspx)|
|[FieldFormat](http://msdn.microsoft.com/library/characters-fieldformat-property-visio%28Office.15%29.aspx)|
|[FieldFormula](http://msdn.microsoft.com/library/characters-fieldformula-property-visio%28Office.15%29.aspx)|
|[FieldFormulaU](http://msdn.microsoft.com/library/characters-fieldformulau-property-visio%28Office.15%29.aspx)|
|[IsField](http://msdn.microsoft.com/library/characters-isfield-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/characters-objecttype-property-visio%28Office.15%29.aspx)|
|[ParaProps](http://msdn.microsoft.com/library/characters-paraprops-property-visio%28Office.15%29.aspx)|
|[ParaPropsRow](http://msdn.microsoft.com/library/characters-parapropsrow-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/characters-persistsevents-property-visio%28Office.15%29.aspx)|
|[RunBegin](http://msdn.microsoft.com/library/characters-runbegin-property-visio%28Office.15%29.aspx)|
|[RunEnd](http://msdn.microsoft.com/library/characters-runend-property-visio%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/characters-shape-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/characters-stat-property-visio%28Office.15%29.aspx)|
|[TabPropsRow](http://msdn.microsoft.com/library/characters-tabpropsrow-property-visio%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/characters-text-property-visio%28Office.15%29.aspx)|

