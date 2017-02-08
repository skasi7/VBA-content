---
title: Module Object (Access)
keywords: vbaac10.chm12268
f1_keywords:
- vbaac10.chm12268
ms.prod: ACCESS
api_name:
- Access.Module
ms.assetid: e04272fa-9c29-2567-bd15-1cea38906894
---


# Module Object (Access)

A  **Module** object refers to a standard module or a class module.


## Remarks

Microsoft Access includes class modules that are not associated with any object, and form modules and report modules, which are associated with a form or report.

To determine whether a  **Module** object represents a standard module or a class module from code, check the **Module** object's **Type** property.

The  **[Modules](http://msdn.microsoft.com/library/modules-object-access%28Office.15%29.aspx)** collection contains all open **Module** objects, regardless of their type. Modules in the **Modules** collection can be compiled or uncompiled.

To return a reference to a particular standard or class  **Module** object in the **Modules** collection, use any of the following syntax forms.



|**Syntax**|**Description**|
|:-----|:-----|
|**Modules** !modulename|The  _modulename_ argument is the name of the **Module** object.|
|**Modules** (" _modulename_")|The  _modulename_ argument is the name of the **Module** object.|
|**Modules** ( _index_)|The  _index_ argument is the numeric position of the object within the collection.|
The following example returns a reference to a standard  **Module** object and assigns it to an object variable:




```
Dim mdl As Module 
Set mdl = Modules![Utility Functions]
```

Note that the brackets enclosing the name of the  **Module** object are necessary only if the name of the **Module** includes spaces.

The next example returns a reference to a form  **Module** object and assigns it to an object variable:




```
Dim mdl As Module 
Set mdl = Modules!Form_Employees
```

To refer to a specific form or report module, you can also use the  **[Form](http://msdn.microsoft.com/library/form-object-access%28Office.15%29.aspx)** or **[Report](report-object-access.md)** object's **Module** property:




```
Forms!formname .Module
```

The following example also returns a reference to the  **Module** object associated with an Employees form and assigns it to an object variable:




```
Dim mdl As Module 
Set mdl = Forms!Employees.Module
```

Once you've returned a reference to a  **Module** object, you can set or read its properties and apply its methods.


## Methods



|**Name**|
|:-----|
|[AddFromFile](http://msdn.microsoft.com/library/module-addfromfile-method-access%28Office.15%29.aspx)|
|[AddFromString](http://msdn.microsoft.com/library/module-addfromstring-method-access%28Office.15%29.aspx)|
|[CreateEventProc](http://msdn.microsoft.com/library/module-createeventproc-method-access%28Office.15%29.aspx)|
|[DeleteLines](http://msdn.microsoft.com/library/module-deletelines-method-access%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/module-find-method-access%28Office.15%29.aspx)|
|[InsertLines](http://msdn.microsoft.com/library/module-insertlines-method-access%28Office.15%29.aspx)|
|[InsertText](http://msdn.microsoft.com/library/module-inserttext-method-access%28Office.15%29.aspx)|
|[ReplaceLine](http://msdn.microsoft.com/library/module-replaceline-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/module-application-property-access%28Office.15%29.aspx)|
|[CountOfDeclarationLines](http://msdn.microsoft.com/library/module-countofdeclarationlines-property-access%28Office.15%29.aspx)|
|[CountOfLines](http://msdn.microsoft.com/library/module-countoflines-property-access%28Office.15%29.aspx)|
|[Lines](http://msdn.microsoft.com/library/module-lines-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/module-name-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/module-parent-property-access%28Office.15%29.aspx)|
|[ProcBodyLine](http://msdn.microsoft.com/library/module-procbodyline-property-access%28Office.15%29.aspx)|
|[ProcCountLines](http://msdn.microsoft.com/library/module-proccountlines-property-access%28Office.15%29.aspx)|
|[ProcOfLine](http://msdn.microsoft.com/library/module-procofline-property-access%28Office.15%29.aspx)|
|[ProcStartLine](http://msdn.microsoft.com/library/module-procstartline-property-access%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/module-type-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[Module Object Members](http://msdn.microsoft.com/library/module-members-access%28Office.15%29.aspx)
