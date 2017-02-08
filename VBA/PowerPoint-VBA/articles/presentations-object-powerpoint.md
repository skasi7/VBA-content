---
title: Presentations Object (PowerPoint)
keywords: vbapp10.chm522000
f1_keywords:
- vbapp10.chm522000
ms.prod: POWERPOINT
ms.assetid: 0b952edc-8628-71ef-e854-3bcefbb3bc61
---


# Presentations Object (PowerPoint)

A collection of all the  **[Presentation](presentation-object-powerpoint.md)** objects in Microsoft PowerPoint. Each **Presentation** object represents a presentation that's currently open in PowerPoint.


## Remarks

The  **Presentations** collection doesn't include open add-ins, which are a special kind of hidden presentation. You can, however, return a single open add-in if you know its file name. For example `Presentations("oscar.ppa")` will return the open add-in named "Oscar.ppa" as a **Presentation** object. However, it is recommended that the **AddIns** collection be used to return open add-ins.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.Presentations.GetEnumerator** (to enumerate the **Presentation** objects.)
    

## Example

Use the [Presentations](http://msdn.microsoft.com/library/application-presentations-property-powerpoint%28Office.15%29.aspx) property to return the **Presentations** collection. Use the[Add](http://msdn.microsoft.com/library/presentations-add-method-powerpoint%28Office.15%29.aspx) method to create a new presentation and add it to the collection. The following example creates a new presentation, adds a slide to the presentation, and then saves the presentation.


```
Set newPres = Presentations.Add(True) 
newPres.Slides.Add 1, 1 
newPres.SaveAs "Sample"
```

Use  **Presentations** (index), where index is the presentation's name or index number, to return a single **Presentation** object. The following example prints presentation one.




```
Presentations(1).PrintOut
```

Use the [Open](http://msdn.microsoft.com/library/presentations-open-method-powerpoint%28Office.15%29.aspx) method to open a presentation and add it to the **Presentations** collection. The following example opens the file Sales.ppt as a read-only presentation.




```
Presentations.Open FileName:="sales.ppt", ReadOnly:=True
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/presentations-add-method-powerpoint%28Office.15%29.aspx)|
|[CanCheckOut](http://msdn.microsoft.com/library/presentations-cancheckout-method-powerpoint%28Office.15%29.aspx)|
|[CheckOut](http://msdn.microsoft.com/library/presentations-checkout-method-powerpoint%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/presentations-item-method-powerpoint%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/presentations-open-method-powerpoint%28Office.15%29.aspx)|
|[Open2007](http://msdn.microsoft.com/library/presentations-open2007-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/presentations-application-property-powerpoint%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/presentations-count-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/presentations-parent-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
