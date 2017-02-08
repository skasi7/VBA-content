---
title: CommandBarComboBox Object (Office)
keywords: vbaof11.chm243000
f1_keywords:
- vbaof11.chm243000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarComboBox
ms.assetid: fcfe6bde-dea0-f1f1-ad30-d0e28f97dd07
---


# CommandBarComboBox Object (Office)

Represents a combo box control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

Use  **Controls(index)**, where _index_ is the index number of the control, to return a **CommandBarComboBox** object. Note that the **Type** property of the control must be **msoControlEdit**, **msoControlDropdown**, **msoControlComboBox**, **msoControlButtonDropdown**, **msoControlSplitDropdown**, **msoControlOCXDropdown**, **msoControlGraphicCombo**, or **msoControlGraphicDropdown**.


## Example

The following example adds two items to the second control on the command bar named  **Custom**, and then it adjusts the size of the control.


```
Set combo = CommandBars("Custom").Controls(2) 
With combo 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListIndex = 0 
End With
```

You can also use the  **FindControl** method to return a **CommandBarComboBox** object. The following example searches all command bars for a visible **CommandBarComboBox** object whose tag is "sheet assignments."




```
Set myControl = CommandBars.FindControl _ 
(Type:=msoControlComboBox, Tag:="sheet assignments", Visible:=True)
```


## Events



|**Name**|
|:-----|
|[Change](http://msdn.microsoft.com/library/commandbarcombobox-change-event-office%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddItem](http://msdn.microsoft.com/library/commandbarcombobox-additem-method-office%28Office.15%29.aspx)|
|[Clear](http://msdn.microsoft.com/library/commandbarcombobox-clear-method-office%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/commandbarcombobox-copy-method-office%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/commandbarcombobox-delete-method-office%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/commandbarcombobox-execute-method-office%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/commandbarcombobox-move-method-office%28Office.15%29.aspx)|
|[RemoveItem](http://msdn.microsoft.com/library/commandbarcombobox-removeitem-method-office%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/commandbarcombobox-reset-method-office%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/commandbarcombobox-setfocus-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/commandbarcombobox-application-property-office%28Office.15%29.aspx)|
|[BeginGroup](http://msdn.microsoft.com/library/commandbarcombobox-begingroup-property-office%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/commandbarcombobox-builtin-property-office%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/commandbarcombobox-caption-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/commandbarcombobox-creator-property-office%28Office.15%29.aspx)|
|[DescriptionText](http://msdn.microsoft.com/library/commandbarcombobox-descriptiontext-property-office%28Office.15%29.aspx)|
|[DropDownLines](http://msdn.microsoft.com/library/commandbarcombobox-dropdownlines-property-office%28Office.15%29.aspx)|
|[DropDownWidth](http://msdn.microsoft.com/library/commandbarcombobox-dropdownwidth-property-office%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/commandbarcombobox-enabled-property-office%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/commandbarcombobox-height-property-office%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/commandbarcombobox-helpcontextid-property-office%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/commandbarcombobox-helpfile-property-office%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/commandbarcombobox-id-property-office%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/commandbarcombobox-index-property-office%28Office.15%29.aspx)|
|[IsPriorityDropped](http://msdn.microsoft.com/library/commandbarcombobox-isprioritydropped-property-office%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/commandbarcombobox-left-property-office%28Office.15%29.aspx)|
|[List](http://msdn.microsoft.com/library/commandbarcombobox-list-property-office%28Office.15%29.aspx)|
|[ListCount](http://msdn.microsoft.com/library/commandbarcombobox-listcount-property-office%28Office.15%29.aspx)|
|[ListHeaderCount](http://msdn.microsoft.com/library/commandbarcombobox-listheadercount-property-office%28Office.15%29.aspx)|
|[ListIndex](http://msdn.microsoft.com/library/commandbarcombobox-listindex-property-office%28Office.15%29.aspx)|
|[OLEUsage](http://msdn.microsoft.com/library/commandbarcombobox-oleusage-property-office%28Office.15%29.aspx)|
|[OnAction](http://msdn.microsoft.com/library/commandbarcombobox-onaction-property-office%28Office.15%29.aspx)|
|[Parameter](http://msdn.microsoft.com/library/commandbarcombobox-parameter-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/commandbarcombobox-parent-property-office%28Office.15%29.aspx)|
|[Priority](http://msdn.microsoft.com/library/commandbarcombobox-priority-property-office%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/commandbarcombobox-style-property-office%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/commandbarcombobox-tag-property-office%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/commandbarcombobox-text-property-office%28Office.15%29.aspx)|
|[TooltipText](http://msdn.microsoft.com/library/commandbarcombobox-tooltiptext-property-office%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/commandbarcombobox-top-property-office%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/commandbarcombobox-type-property-office%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/commandbarcombobox-visible-property-office%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/commandbarcombobox-width-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[CommandBarComboBox Object Members](http://msdn.microsoft.com/library/commandbarcombobox-members-office%28Office.15%29.aspx)
