---
title: TempVars Object (Access)
keywords: vbaac10.chm14073
f1_keywords:
- vbaac10.chm14073
ms.prod: ACCESS
api_name:
- Access.TempVars
ms.assetid: aa81b18b-5e9f-ae44-cbcf-55cf6e37b7f6
---


# TempVars Object (Access)

Represents the collection of  **[TempVar](tempvar-object-access.md)** objects.


## Remarks

Use the  **[Add](http://msdn.microsoft.com/library/tempvars-add-method-access%28Office.15%29.aspx)** method or the[SetTempVar](http://msdn.microsoft.com/library/settempvar-macro-action%28Office.15%29.aspx) macro action to create a **TempVar** object.

Use the  **[Remove](http://msdn.microsoft.com/library/tempvars-remove-method-access%28Office.15%29.aspx)** method or the[RemoveTempVar](http://msdn.microsoft.com/library/removealltempvars-macro-action%28Office.15%29.aspx) macro action to delete a **TempVar** object from the **TempVars** collection.

Use the  **[RemoveAll](http://msdn.microsoft.com/library/tempvars-removeall-method-access%28Office.15%29.aspx)** method or[RemoveAllTempVars](http://msdn.microsoft.com/library/removealltempvars-macro-action%28Office.15%29.aspx) macro action to delete all **TempVar** objects from the **TempVars** collection.

The  **TempVars** collection can store up to 255 **TempVar** objects. If you do not remove a **TempVar** object, it will remain in memory until you close the database. It is a good practice to remove **TempVar** object variables when you are finished using them.

To refer to a  **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:


-  **TempVar** ![name]
    

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/tempvars-add-method-access%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/tempvars-remove-method-access%28Office.15%29.aspx)|
|[RemoveAll](http://msdn.microsoft.com/library/tempvars-removeall-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/tempvars-application-property-access%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/tempvars-count-property-access%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/tempvars-item-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/tempvars-parent-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[TempVars Object Members](http://msdn.microsoft.com/library/tempvars-members-access%28Office.15%29.aspx)
