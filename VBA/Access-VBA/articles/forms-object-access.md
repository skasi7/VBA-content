---
title: Forms Object (Access)
keywords: vbaac10.chm12355
f1_keywords:
- vbaac10.chm12355
ms.prod: ACCESS
api_name:
- Access.Forms
ms.assetid: a41af7be-873c-ef8b-20cd-24b78a25b5ca
---


# Forms Object (Access)

The  **Forms** collection contains all of the currently open forms in a Microsoft Access database.


## Remarks

Use the  **Forms** collection in Visual Basic or in an expression to refer to forms that are currently open. For example, you can enumerate the **Forms** collection to set or return the values of properties of individual forms in the collection.

You can refer to an individual  **Form** object in the **Forms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific form in the **Forms** collection, it's better to refer to the form by name because a form's collection index may change.

The  **Forms** collection is indexed beginning with zero. If you refer to a form by its index, the first form opened is Forms(0), the second form opened is Forms(1), and so on. If you opened Form1 and then opened Form2, Form2 would be referenced in the **Forms** collection by its index as Forms(1). If you then closed Form1, Form2 would be referenced in the **Forms** collection by its index as Forms(0).


 **Note**   To list all forms in the database, whether open or closed, enumerate the **AllForms** collection of the **[CurrentProject](currentproject-object-access.md)** object. You can then use the **Name** property of each individual **[AccessObject](accessobject-object-access.md)** object to return the name of a form.

You can't add or delete a  **Form** object from the **Forms** collection.


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/forms-application-property-access%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/forms-count-property-access%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/forms-item-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/forms-parent-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[Forms Object Members](http://msdn.microsoft.com/library/forms-members-access%28Office.15%29.aspx)
