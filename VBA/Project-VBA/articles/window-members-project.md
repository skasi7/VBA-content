---
title: Window Members (Project)
ms.prod: PROJECTSERVER
ms.assetid: 842c9b2d-9a35-cd6d-3960-606c07f48c77
---


# Window Members (Project)
Represents a window in the application or project. The  **Window** object is a member of the **[Windows](windows-object-project.md)** collection.

Represents a window in the application or project. The  **Window** object is a member of the **[Windows](windows-object-project.md)** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](window-activate-method-project.md)|Activates the window, bringing the window to the front of the z-order.|
|[Close](window-close-method-project.md)|Closes a pane or window.|
|[Refresh](window-refresh-method-project.md)|Refreshes the window.|
|[WebBrowserControlFrame](window-webbrowsercontrolframe-method-project.md)|Returns the DOM object of a specified frame in the Web browser control window hosted inside the active window.|
|[WebBrowserControlWindow](window-webbrowsercontrolwindow-method-project.md)|Returns the DOM object for the Microsoft Internet Explorer window loaded in the Web browser control that is hosted within the specified window in Project.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActivePane](window-activepane-property-project.md)|Gets a  **[Pane](pane-object-project.md)** object that represents the active pane of a window. Read-only **Pane**.|
|[Application](window-application-property-project.md)|Gets the  **[Application](application-object-project.md)** object. Read-only **Application**.|
|[BottomPane](window-bottompane-property-project.md)|Gets a  **[Pane](pane-object-project.md)** object representing the bottom pane of a window. Read-only **Pane**.|
|[Caption](window-caption-property-project.md)|Gets or sets the text in the title bar of a project window. Read/write  **String**.|
|[Height](window-height-property-project.md)|Gets or sets the height of a project window in points. Read/write  **Long**.|
|[Index](window-index-property-project.md)|Gets the index of a  **Window** object in the containing object. Read-only **Long**.|
|[Left](window-left-property-project.md)|Gets or sets the distance of a project window from the left edge of the main window in points. Read/write  **Long**.|
|[Parent](window-parent-property-project.md)|Gets the parent of the  **Window** object. Read-only **Object**.|
|[Top](window-top-property-project.md)|Gets or sets the distance in points of the window below the top edge of the window display area. Read/write  **Long**.|
|[TopPane](window-toppane-property-project.md)|Gets a  **[Pane](pane-object-project.md)** object representing the top pane of the window. Read-only **Pane**.|
|[Visible](window-visible-property-project.md)|**True** if the window is visible. Read/write **Boolean**.|
|[Width](window-width-property-project.md)|Gets or sets the width in points of the window. Read/write  **Long**.|
|[WindowState](window-windowstate-property-project.md)|Gets or sets the state the window, where the state is maximized or normal. Read/write  **PjWindowState**.|

