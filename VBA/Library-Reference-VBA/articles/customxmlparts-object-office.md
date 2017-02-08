---
title: CustomXMLParts Object (Office)
keywords: vbaof11.chm300000
f1_keywords:
- vbaof11.chm300000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CustomXMLParts
ms.assetid: 98c1c58e-a08d-6304-8626-1e6705917da3
---


# CustomXMLParts Object (Office)

Represents a collection of  **CustomXMLPart** objects.


## Remarks

There are three default parts that are always created with a document. These are 'Cover pages', 'Doc properties' and 'App properties'. The last two were in previous versions of Microsoft Word but are now provided in XML form in the  **CustomXMLParts** object collection


## Example

The following example adds a node to a  **CustomXMLPart** object that is part of the **CustomXMLParts** object collection.


```
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## Events



|**Name**|
|:-----|
|[PartAfterAdd](http://msdn.microsoft.com/library/customxmlparts-partafteradd-event-office%28Office.15%29.aspx)|
|[PartAfterLoad](http://msdn.microsoft.com/library/customxmlparts-partafterload-event-office%28Office.15%29.aspx)|
|[PartBeforeDelete](http://msdn.microsoft.com/library/customxmlparts-partbeforedelete-event-office%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/customxmlparts-add-method-office%28Office.15%29.aspx)|
|[SelectByID](http://msdn.microsoft.com/library/customxmlparts-selectbyid-method-office%28Office.15%29.aspx)|
|[SelectByNamespace](http://msdn.microsoft.com/library/customxmlparts-selectbynamespace-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/customxmlparts-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/customxmlparts-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/customxmlparts-creator-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/customxmlparts-item-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/customxmlparts-parent-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[CustomXMLParts Object Members](http://msdn.microsoft.com/library/customxmlparts-members-office%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
