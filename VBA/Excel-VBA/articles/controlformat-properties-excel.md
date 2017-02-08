---
title: ControlFormat Properties (Excel)
ms.prod: EXCEL
ms.assetid: a2f74be4-cde2-4062-8370-831686f467ab
---


# ControlFormat Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](controlformat-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](controlformat-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DropDownLines](controlformat-dropdownlines-property-excel.md)|Returns or sets the number of list lines displayed in the drop-down portion of a combo box. Read/write  **Long** .|
|[Enabled](controlformat-enabled-property-excel.md)| **True** if the object is enabled. Read/write **Boolean** .|
|[LargeChange](controlformat-largechange-property-excel.md)|Returns or sets the amount that the scroll box increments or decrements for a page scroll (when the user clicks in the scroll bar body region). Read/write  **Long** .|
|[LinkedCell](controlformat-linkedcell-property-excel.md)|Returns or sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell. Read/write  **String** .|
|[ListCount](controlformat-listcount-property-excel.md)|Returns the number of entries in a list box or combo box. Returns 0 (zero) if there are no entries in the list. Read-only  **Long** .|
|[ListFillRange](controlformat-listfillrange-property-excel.md)|Returns or sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box. Read/write  **String** .|
|[ListIndex](controlformat-listindex-property-excel.md)|Returns or sets the index number of the currently selected item in a list box or combo box. Read/write  **Long** .|
|[LockedText](controlformat-lockedtext-property-excel.md)| **True** if the text in the specified object will be locked to prevent changes when the workbook is protected. Read/write **Boolean** .|
|[Max](controlformat-max-property-excel.md)|Returns or sets the maximum value of a scroll bar or spinner range. The scroll bar or spinner won't take on values greater than this maximum value. Read/write  **Long** .|
|[Min](controlformat-min-property-excel.md)|Returns or sets the minimum value of a scroll bar or spinner range. The scroll bar or spinner won't take on values less than this minimum value. Read/write  **Long** .|
|[MultiSelect](controlformat-multiselect-property-excel.md)|Returns or sets the selection mode of the specified list box. Can be one of the following constants:  **xlNone** , **xlSimple** , or **xlExtended** . Read/write **Long** .|
|[Parent](controlformat-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PrintObject](controlformat-printobject-property-excel.md)| **True** if the object will be printed when the document is printed. Read/write **Boolean** .|
|[SmallChange](controlformat-smallchange-property-excel.md)|Returns or sets the amount that the scroll bar or spinner is incremented or decremented for a line scroll (when the user clicks an arrow). Read/write  **Long** .|
|[Value](controlformat-value-property-excel.md)|Returns or sets a  **Long** value that represents the name of specified control format.|

