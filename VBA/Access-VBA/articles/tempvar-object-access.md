---
title: TempVar Object (Access)
keywords: vbaac10.chm14063
f1_keywords:
- vbaac10.chm14063
ms.prod: ACCESS
api_name:
- Access.TempVar
ms.assetid: 4a0429e6-bcfa-7a8b-7030-6e88c2f1a71d
---


# TempVar Object (Access)

Represents a variable that be used in Visual Basic for Applications (VBA) code or from a macro. 


## Remarks

A  **TempVar** objects provide a convenient way to exchange data between VBA procedures and macros.

Although a  **TempVar** object can be used to store information for use in VBA procedures, it does not have the same funcitonality as a VBA variable.


- By default, a  **TempVar** object remains in memory until Access is closed. You can use the **[Remove](http://msdn.microsoft.com/library/tempvars-remove-method-access%28Office.15%29.aspx)** method or the[RemoveTempVar](http://msdn.microsoft.com/library/removetempvar-macro-action%28Office.15%29.aspx) macro action to remove a **TempVar** object.
    
- In VBA, a  **TempVar** object is accessible only to the members of the Access **[Application](http://msdn.microsoft.com/library/application-object-access%28Office.15%29.aspx)** object, referenced databases, or add-ins.
    
- A  **TempVar** object can store only text or numeric data. **TempVar** objects cannot store objects.
    
To refer to a  **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:


-  **TempVar** ![name]
    

|**Name**|
|:-----|
|[Name](http://msdn.microsoft.com/library/tempvar-name-property-access%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/tempvar-value-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[TempVar Object Members](http://msdn.microsoft.com/library/tempvar-members-access%28Office.15%29.aspx)
