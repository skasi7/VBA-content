---
title: Screen Object (Access)
keywords: vbaac10.chm12484
f1_keywords:
- vbaac10.chm12484
ms.prod: ACCESS
api_name:
- Access.Screen
ms.assetid: 00743775-071b-9ccd-7687-f3b992e9346e
---


# Screen Object (Access)

The  **Screen** object refers to the particular form, report, or control that currently has the focus.


## Remarks

You can use the  **Screen** object together with its properties to refer to a particular form, report, or control that has the focus.

For example, you can use the  **Screen** object with the **ActiveForm** property to refer to the form in the active window without knowing the form's name. The following example displays the name of the form in the active window:




```
MsgBox Screen.ActiveForm.Name
```

Referring to the  **Screen** object doesn't make a form, report, or control active. To make a form, report, or control active, you must use the **SelectObject** method of the **[DoCmd](http://msdn.microsoft.com/library/docmd-object-access%28Office.15%29.aspx)** object.

If you refer to the  **Screen** object when there's no active form, report, or control, Microsoft Access returns a run-time error. For example, if a standard module is in the active window, the code in the preceding example would return an error.


## Example

The following example uses the  **Screen** object to print the name of the form in the active window and of the active control on that form:


```
Sub ActiveObjects() 
 Dim frm As Form, ctl As Control 
 
 ' Return Form object pointing to active form. 
 Set frm = Screen.ActiveForm 
 MsgBox frm.Name &amp; " is the active form." 
 ' Return Control object pointing to active control. 
 Set ctl = Screen.ActiveControl 
 MsgBox ctl.Name &amp; " is the active control " _ 
 &amp; "on this form." 
End Sub 

```



|**Name**|
|:-----|
|[ActiveControl](http://msdn.microsoft.com/library/screen-activecontrol-property-access%28Office.15%29.aspx)|
|[ActiveDatasheet](http://msdn.microsoft.com/library/screen-activedatasheet-property-access%28Office.15%29.aspx)|
|[ActiveForm](http://msdn.microsoft.com/library/screen-activeform-property-access%28Office.15%29.aspx)|
|[ActiveReport](http://msdn.microsoft.com/library/screen-activereport-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/screen-application-property-access%28Office.15%29.aspx)|
|[MousePointer](http://msdn.microsoft.com/library/screen-mousepointer-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/screen-parent-property-access%28Office.15%29.aspx)|
|[PreviousControl](http://msdn.microsoft.com/library/screen-previouscontrol-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Screen Object Members](http://msdn.microsoft.com/library/screen-members-access%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
