---
title: CommandBar Object (Office)
keywords: vbaof11.chm3000
f1_keywords:
- vbaof11.chm3000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBar
ms.assetid: 78603954-40aa-64cb-c407-2e0820d65231
---


# CommandBar Object (Office)

Represents a command bar in the container application. The  **CommandBar** object is a member of the **CommandBars** collection.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use  **CommandBars** ( _index_ ), where _index_ is the name or index number of a command bar, to return a single **CommandBar** object. The following example steps through the collection of command bars to find the command bar named "Forms." If it finds this command bar, the example makes it visible and protects its docking state. In this example, the variable **cb** represents a **CommandBar** object.


```
foundFlag = False  
For Each cb In CommandBars 
    If cb.Name = "Forms" Then 
        cb.Protection = msoBarNoChangeDock 
        cb.Visible = True  
        foundFlag = True  
    End If 
Next cb 
If Not foundFlag Then 
    MsgBox "The collection does not contain a Forms command bar." 
End If
```

You can use a name or index number to specify a menu bar or toolbar in the list of available menu bars and toolbars in the container application. However, you must use a name to specify a menu, shortcut menu, or submenu (all of which are represented by  **CommandBar** objects). This example adds a new menu item to the bottom of the **Tools** menu. When clicked, the new menu item runs the procedure named "qtrReport."




```
Set newItem = CommandBars("Tools").Controls.Add(Type:=msoControlButton) 
With newItem 
    .BeginGroup = True  
    .Caption = "Make Report" 
    .FaceID = 0 
    .OnAction = "qtrReport" 
End With
```

If two or more custom menus or submenus have the same name,  **CommandBars(index)** returns the first one. To ensure that you return the correct menu or submenu, locate the pop-up control that displays that menu. Then apply the **CommandBar** property to the pop-up control to return the command bar that represents that menu. Assuming that the third control on the toolbar named "Custom Tools" is a pop-up control, this example adds the **Save** command to the bottom of that menu.




```
Set viewMenu = CommandBars("Custom Tools").Controls(3) 
viewMenu.Controls.Add ID:=3    'ID of Save command is 3
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/commandbar-delete-method-office%28Office.15%29.aspx)|
|[FindControl](http://msdn.microsoft.com/library/commandbar-findcontrol-method-office%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/commandbar-reset-method-office%28Office.15%29.aspx)|
|[ShowPopup](http://msdn.microsoft.com/library/commandbar-showpopup-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AdaptiveMenu](http://msdn.microsoft.com/library/commandbar-adaptivemenu-property-office%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/commandbar-application-property-office%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/commandbar-builtin-property-office%28Office.15%29.aspx)|
|[Context](http://msdn.microsoft.com/library/commandbar-context-property-office%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/commandbar-controls-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/commandbar-creator-property-office%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/commandbar-enabled-property-office%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/commandbar-height-property-office%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/commandbar-index-property-office%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/commandbar-left-property-office%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/commandbar-name-property-office%28Office.15%29.aspx)|
|[NameLocal](http://msdn.microsoft.com/library/commandbar-namelocal-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/commandbar-parent-property-office%28Office.15%29.aspx)|
|[Position](http://msdn.microsoft.com/library/commandbar-position-property-office%28Office.15%29.aspx)|
|[Protection](http://msdn.microsoft.com/library/commandbar-protection-property-office%28Office.15%29.aspx)|
|[RowIndex](http://msdn.microsoft.com/library/commandbar-rowindex-property-office%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/commandbar-top-property-office%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/commandbar-type-property-office%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/commandbar-visible-property-office%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/commandbar-width-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[CommandBar Object Members](http://msdn.microsoft.com/library/commandbar-members-office%28Office.15%29.aspx)
