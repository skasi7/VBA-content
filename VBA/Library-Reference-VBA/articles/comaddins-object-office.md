---
title: COMAddIns Object (Office)
keywords: vbaof11.chm220000
f1_keywords:
- vbaof11.chm220000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.COMAddIns
ms.assetid: f6efa1cc-8d30-27d5-8b07-7ddad22f16ef
---


# COMAddIns Object (Office)

A collection of  **COMAddIn** objects that provide information about a COM add-in registered in the Windows registry.


## Example

Use the  **COMAddIns** property of the **Application** object to return the **COMAddIns** collection for a Microsoft Office host application. This collection contains all of the COM add-ins that are available to a given Office host application, and the **Count** property of the **COMAddins** collection returns the number of available COM add-ins, as in the following example.


```
MsgBox Application.COMAddIns.Count
```

Use the  **Update** method of the **COMAddins** collection to refresh the list of COM add-ins from the Windows registry, as in the following example.




```
Application.COMAddIns.Update
```

Use  **COMAddIns.Item(index)**, where _index_ is either an ordinal value that returns the COM add-in at that position in the **COMAddIns** collection, or a **String** value that represents the ProgID of the specified COM add-in. The following example displays a COM add-in's description text and ProgID (" **msodraa9.ShapeSelect** ") in a message box.




```
MsgBox Application.COMAddIns.Item("msodraa9.ShapeSelect").Description
```


## Methods



|**Name**|
|:-----|
|[Item](http://msdn.microsoft.com/library/comaddins-item-method-office%28Office.15%29.aspx)|
|[Update](http://msdn.microsoft.com/library/comaddins-update-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/comaddins-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/comaddins-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/comaddins-creator-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/comaddins-parent-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[COMAddIns Object Members](http://msdn.microsoft.com/library/comaddins-members-office%28Office.15%29.aspx)
