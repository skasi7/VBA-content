---
title: CommandBars Object (Office)
keywords: vbaof11.chm242000
f1_keywords:
- vbaof11.chm242000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBars
ms.assetid: 0e312e21-14ee-5055-d604-b66e61c53b47
---


# CommandBars Object (Office)

A collection of  **CommandBar** objects that represent the command bars in the container application.


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use the  **CommandBars** property to return the **CommandBars** collection. The following example displays in the **Immediate** window both the name and local name of each menu bar and toolbar, and it displays a value that indicates whether the menu bar or toolbar is visible.


```
For Each cbar in CommandBars 
    Debug.Print cbar.Name, cbar.NameLocal, cbar.Visible 
Next
```

Use the  **Add** method to add a new command bar to the collection. The following example creates a custom toolbar named "Custom1" and displays it as a floating toolbar.




```
Set cbar1 = CommandBars.Add(Name:="Custom1", Position:=msoBarFloating) 
cbar1.Visible = True
```

Use enumName, where  _index_ is the name or index number of a command bar, to return a single **CommandBar** object. The following example docks the toolbar named "Custom1" at the bottom of the application window.




```
CommandBars("Custom1").Position = msoBarBottom
```


 **Note**  You can use the name or index number to specify a menu bar or toolbar in the list of available menu bars and toolbars in the container application. However, you must use the name to specify a menu, shortcut menu, or submenu (all of which are represented by  **CommandBar** objects). If two or more custom menus or submenus have the same name, enumName returns the first one. To ensure that you return the correct menu or submenu, locate the pop-up control that displays that menu. Then apply the **CommandBar** property to the pop-up control to return the command bar that represents that menu.


## Events



|**Name**|
|:-----|
|[OnUpdate](http://msdn.microsoft.com/library/commandbars-onupdate-event-office%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/commandbars-add-method-office%28Office.15%29.aspx)|
|[CommitRenderingTransaction](http://msdn.microsoft.com/library/commandbars-commitrenderingtransaction-method-office%28Office.15%29.aspx)|
|[ExecuteMso](http://msdn.microsoft.com/library/commandbars-executemso-method-office%28Office.15%29.aspx)|
|[FindControl](http://msdn.microsoft.com/library/commandbars-findcontrol-method-office%28Office.15%29.aspx)|
|[FindControls](http://msdn.microsoft.com/library/commandbars-findcontrols-method-office%28Office.15%29.aspx)|
|[GetEnabledMso](http://msdn.microsoft.com/library/commandbars-getenabledmso-method-office%28Office.15%29.aspx)|
|[GetImageMso](http://msdn.microsoft.com/library/commandbars-getimagemso-method-office%28Office.15%29.aspx)|
|[GetLabelMso](http://msdn.microsoft.com/library/commandbars-getlabelmso-method-office%28Office.15%29.aspx)|
|[GetPressedMso](http://msdn.microsoft.com/library/commandbars-getpressedmso-method-office%28Office.15%29.aspx)|
|[GetScreentipMso](http://msdn.microsoft.com/library/commandbars-getscreentipmso-method-office%28Office.15%29.aspx)|
|[GetSupertipMso](http://msdn.microsoft.com/library/commandbars-getsupertipmso-method-office%28Office.15%29.aspx)|
|[GetVisibleMso](http://msdn.microsoft.com/library/commandbars-getvisiblemso-method-office%28Office.15%29.aspx)|
|[ReleaseFocus](http://msdn.microsoft.com/library/commandbars-releasefocus-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActionControl](http://msdn.microsoft.com/library/commandbars-actioncontrol-property-office%28Office.15%29.aspx)|
|[ActiveMenuBar](http://msdn.microsoft.com/library/commandbars-activemenubar-property-office%28Office.15%29.aspx)|
|[AdaptiveMenus](http://msdn.microsoft.com/library/commandbars-adaptivemenus-property-office%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/commandbars-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/commandbars-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/commandbars-creator-property-office%28Office.15%29.aspx)|
|[DisableAskAQuestionDropdown](http://msdn.microsoft.com/library/commandbars-disableaskaquestiondropdown-property-office%28Office.15%29.aspx)|
|[DisableCustomize](http://msdn.microsoft.com/library/commandbars-disablecustomize-property-office%28Office.15%29.aspx)|
|[DisplayFonts](http://msdn.microsoft.com/library/commandbars-displayfonts-property-office%28Office.15%29.aspx)|
|[DisplayKeysInTooltips](http://msdn.microsoft.com/library/commandbars-displaykeysintooltips-property-office%28Office.15%29.aspx)|
|[DisplayTooltips](http://msdn.microsoft.com/library/commandbars-displaytooltips-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/commandbars-item-property-office%28Office.15%29.aspx)|
|[LargeButtons](http://msdn.microsoft.com/library/commandbars-largebuttons-property-office%28Office.15%29.aspx)|
|[MenuAnimationStyle](http://msdn.microsoft.com/library/commandbars-menuanimationstyle-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/commandbars-parent-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[CommandBars Object Members](http://msdn.microsoft.com/library/commandbars-members-office%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
