---
title: Folder Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 788acd42-377a-1803-7713-50e45086e2d1
---


# Folder Members (Outlook)
Represents an Outlook folder.

Represents an Outlook folder.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeFolderMove](folder-beforefoldermove-event-outlook.md)|Occurs when a folder is about to be moved or deleted, either as a result of user action or through program code. |
|[BeforeItemMove](folder-beforeitemmove-event-outlook.md)|Occurs when an item is about to be moved or deleted from a folder, either as a result of user action or through program code. |

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddToPFFavorites](folder-addtopffavorites-method-outlook.md)|Adds a Microsoft Exchange public folder to the public folder's Favorites folder.|
|[CopyTo](folder-copyto-method-outlook.md)|Copies the current folder in its entirety to the destination folder. |
|[Delete](folder-delete-method-outlook.md)|Deletes an object from the collection.|
|[Display](folder-display-method-outlook.md)|Displays a new  **[Explorer](explorer-object-outlook.md)** object for the folder.|
|[GetCalendarExporter](folder-getcalendarexporter-method-outlook.md)|Creates a  **[CalendarSharing](calendarsharing-object-outlook.md)** object for the specified **[Folder](folder-object-outlook.md)** .|
|[GetCustomIcon](folder-getcustomicon-method-outlook.md)|Returns an  **[IPictureDisp](http://msdn.microsoft.com/en-us/library/ms680762%28VS.85%29.aspx)** object that represents the custom icon for the folder.|
|[GetExplorer](folder-getexplorer-method-outlook.md)|Returns an  **[Explorer](explorer-object-outlook.md)** object that represents a new, inactive **Explorer** object initialized with the specified folder as the current folder.|
|[GetStorage](folder-getstorage-method-outlook.md)|Gets a  **[StorageItem](storageitem-object-outlook.md)** object on the parent **[Folder](folder-object-outlook.md)** to store data for an Outlook solution.|
|[GetTable](folder-gettable-method-outlook.md)|Obtains a  **[Table](table-object-outlook.md)** object that contains items filtered by _Filter_ .|
|[MoveTo](folder-moveto-method-outlook.md)|Moves a folder to the specified destination folder.|
|[SetCustomIcon](folder-setcustomicon-method-outlook.md)|Sets a custom icon that is specified by  _Picture_ for the folder.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AddressBookName](folder-addressbookname-property-outlook.md)|Returns or sets a  **String** that indicates the Address Book name for the **[Folder](folder-object-outlook.md)** object representing a Contacts folder. Read/write.|
|[Application](folder-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Class](folder-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[CurrentView](folder-currentview-property-outlook.md)|Returns a  **[View](view-object-outlook.md)** object representing the current view. Read-only.|
|[CustomViewsOnly](folder-customviewsonly-property-outlook.md)|Returns or sets a  **Boolean** that determines which views are displayed on the **View** menu for a given folder. Read/write.|
|[DefaultItemType](folder-defaultitemtype-property-outlook.md)|Returns a constant from the  **[OlItemType](olitemtype-enumeration-outlook.md)** enumeration indicating the default Outlook item type contained in the folder. Read-only.|
|[DefaultMessageClass](folder-defaultmessageclass-property-outlook.md)|Returns a  **String** representing the default message class for items in the folder. Read-only.|
|[Description](folder-description-property-outlook.md)|Returns or sets a  **String** representing the description of the folder. Read/write.|
|[EntryID](folder-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[FolderPath](folder-folderpath-property-outlook.md)|Returns a  **String** that indicates the path of the current folder. Read-only.|
|[Folders](folder-folders-property-outlook.md)|Returns the  **[Folders](folders-object-outlook.md)** collection that represents all the folders contained in the specified **[Folder](folder-object-outlook.md)** . Read-only.|
|[InAppFolderSyncObject](folder-inappfoldersyncobject-property-outlook.md)|Returns or sets a  **Boolean** that determines if the specified folder will be synchronized with the e-mail server. Read/write.|
|[IsSharePointFolder](folder-issharepointfolder-property-outlook.md)|Returns a  **Boolean** that determines if the folder is a Microsoft SharePoint Foundation folder. Read-only.|
|[Items](folder-items-property-outlook.md)|Returns an  **[Items](items-object-outlook.md)** collection object as a collection of Outlook items in the specified folder. Read-only.|
|[Name](folder-name-property-outlook.md)|Returns or sets a  **String** value that represents the display name for the object. Read/write.|
|[Parent](folder-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](folder-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[Folder](folder-object-outlook.md)** object. Read-only.|
|[Session](folder-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[ShowAsOutlookAB](folder-showasoutlookab-property-outlook.md)|Returns or sets a  **Boolean** variable that specifies whether the contact items folder will be displayed as an address list in the Outlook Address Book. Read/write.|
|[ShowItemCount](folder-showitemcount-property-outlook.md)|Sets or returns a constant in the  **[OlShowItemCount](olshowitemcount-enumeration-outlook.md)** enumeration that indicates whether to display the number of unread messages in the folder or the total number of items in the folder in the Navigation Pane. Read/write.|
|[Store](folder-store-property-outlook.md)|Returns a  **[Store](store-object-outlook.md)** object representing the store that contains the **[Folder](folder-object-outlook.md)** object. Read-only.|
|[StoreID](folder-storeid-property-outlook.md)|Returns a  **String** indicating the store ID for the folder. Read-only.|
|[UnReadItemCount](folder-unreaditemcount-property-outlook.md)|Returns a  **Long** indicating the number of unread items in the folder. Read-only.|
|[UserDefinedProperties](folder-userdefinedproperties-property-outlook.md)|Returns a  **[UserDefinedProperties](userdefinedproperties-object-outlook.md)** object that represents the user-defined custom properties for the **[Folder](folder-object-outlook.md)** object. Read-only.|
|[Views](folder-views-property-outlook.md)|Returns the  **[Views](views-object-outlook.md)** collection object of the **[Folder](folder-object-outlook.md)** object. Read-only.|
|[WebViewOn](folder-webviewon-property-outlook.md)|Returns or sets a  **Boolean** indicating the Web view state for a folder. Read/write.|
|[WebViewURL](folder-webviewurl-property-outlook.md)|Returns or sets a  **String** indicating the URL of the Web page that is assigned to a folder. Read/write.|

