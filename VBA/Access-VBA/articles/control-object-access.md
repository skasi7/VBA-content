---
title: Control Object (Access)
keywords: vbaac10.chm10174
f1_keywords:
- vbaac10.chm10174
ms.prod: ACCESS
api_name:
- Access.Control
ms.assetid: ce2362e5-4390-590e-06c0-6f27e8d988cd
---


# Control Object (Access)

The  **Control** object represents a control on a form, report, or section, within another control, or attached to another control.


## Remarks

All controls on a form or report belong to the  **Controls** collection for that **Form** or **Report** object. Controls within a particular section belong to the **Controls** collection for that section. Controls within a tab control or option group control belong to the **Controls** collection for that control. A label control that is attached to another control belongs to the **Controls** collection for that control.

When you refer to an individual  **Control** object in the **Controls** collection, you can refer to the **Controls** collection either implicitly or explicitly.




```
' Implicitly refer to NewData control in Controls 
' collection. 
Me!NewData
```




```
' Use if control name contains space. 
Me![New Data]
```




```
' Performance slightly slower. 
Me("NewData")
```




```
' Refer to a control by its index in the controls 
' collection. 
Me(0)
```




```
' Refer to a NewData control by using the subform 
' Controls collection. 
Me.ctlSubForm.Controls!NewData
```




```
' Explicitly refer to the NewData control in the 
' Controls collection. 
Me.Controls!NewData
```




```
Me.Controls("NewData")
```




```
Me.Controls(0)
```


 **Note**  You can use the  **Me** keyword to represent a **Form** or **Report** object within code only if you're referring to the form or report from code within the class module. If you're referring to a form or report from a standard module or a different form's or report's module, you must use the full reference to the form or report.

Each  **Control** object is denoted by a particular intrinsic constant. For example, the intrinsic constant **acTextBox** is associated with a text box control, and **acCommandButton** is associated with a command button. The constants for the various Microsoft Access controls are set forth in the control's **ControlType** property.

To determine the type of an existing control, you can use the  **ControlType** property. However, you don't need to know the specific type of a control in order to use it in code. You can simply represent it with a variable of data type **Control**.

If you do know the data type of the control to which you are referring, and the control is a built-in Microsoft Access control, you should represent it with a variable of a specific type. For example, if you know that a particular control is a text box, declare a variable of type  **TextBox** to represent it, as shown in the following code.




```
Dim txt As TextBox 
Set txt = Forms!Employees!LastName 

```


 **Note**  If a control is an ActiveX control, then you must declare a variable of type  **Control** to represent it; you cannot use a specific type. If you're not certain what type of control a variable will point to, declare the variable as type **Control**.

The option group control can contain other controls within its  **Controls** collection, including option button, check box, toggle button, and label controls.

The tab control contains a  **[Pages](http://msdn.microsoft.com/library/pages-object-access%28Office.15%29.aspx)** collection, which is a special type of **Controls** collection. The **Pages** collection contains **[Page](http://msdn.microsoft.com/library/page-object-access%28Office.15%29.aspx)** objects, which are controls. Each **Page** object in turn contains a **Controls** collection, which contains all of the controls on that page.

Other  **Control** objects have a **Controls** collection that can contain an attached label. These controls include the text box, option group, option button, toggle button, check box, combo box, list box, command button, bound object frame, and unbound object frame controls.


## Methods



|**Name**|
|:-----|
|[Dropdown](http://msdn.microsoft.com/library/control-dropdown-method-access%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/control-move-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/control-requery-method-access%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/control-setfocus-method-access%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/control-sizetofit-method-access%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/control-undo-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/control-application-property-access%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/control-bottompadding-property-access%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/control-column-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/control-controls-property-access%28Office.15%29.aspx)|
|[Form](http://msdn.microsoft.com/library/control-form-property-access%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/control-gridlinecolor-property-access%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/control-gridlinestylebottom-property-access%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/control-gridlinestyleleft-property-access%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/control-gridlinestyleright-property-access%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/control-gridlinestyletop-property-access%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/control-gridlinewidthbottom-property-access%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/control-gridlinewidthleft-property-access%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/control-gridlinewidthright-property-access%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/control-gridlinewidthtop-property-access%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/control-horizontalanchor-property-access%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/control-hyperlink-property-access%28Office.15%29.aspx)|
|[ItemData](http://msdn.microsoft.com/library/control-itemdata-property-access%28Office.15%29.aspx)|
|[ItemsSelected](http://msdn.microsoft.com/library/control-itemsselected-property-access%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/control-layout-property-access%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/control-layoutid-property-access%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/control-leftpadding-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/control-name-property-access%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/control-object-property-access%28Office.15%29.aspx)|
|[ObjectVerbs](http://msdn.microsoft.com/library/control-objectverbs-property-access%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/control-oldvalue-property-access%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/control-pages-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/control-parent-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/control-properties-property-access%28Office.15%29.aspx)|
|[Report](http://msdn.microsoft.com/library/control-report-property-access%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/control-rightpadding-property-access%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/control-selected-property-access%28Office.15%29.aspx)|
|[SmartTags](http://msdn.microsoft.com/library/control-smarttags-property-access%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/control-toppadding-property-access%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/control-verticalanchor-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Control Object Members](http://msdn.microsoft.com/library/control-members-access%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
