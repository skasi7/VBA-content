---
title: Folder Object (Outlook)
keywords: vbaol11.chm3020
f1_keywords:
- vbaol11.chm3020
ms.prod: OUTLOOK
api_name:
- Outlook.Folder
ms.assetid: 3cf6cda8-6d70-666e-2643-9d9c5b9cacfc
---


# Folder Object (Outlook)

Represents an Outlook folder.


## Remarks

A  **Folder** object can contain other **Folder** objects, as well as Outlook items. Use the **Folders** property of a **[NameSpace](namespace-object-outlook.md)** object or another **Folder** object to return the set of folders in a **NameSpace** or under a folder. You can navigate nested folders by starting from a top-level folder, say the Inbox, and using a combination of the **[Folder.Folders](folder-folders-property-outlook.md)** property, which returns the set of folders underneath a **Folder** object in the hierarchy, and the **[Folders.Item](http://msdn.microsoft.com/library/folders-item-method-outlook%28Office.15%29.aspx)** method, which returns a folder within the **[Folders](http://msdn.microsoft.com/library/folders-object-outlook%28Office.15%29.aspx)** collection.

There is a set of folders within an Outlook data store that supports the default functionality of Outlook. Use  **[NameSpace.GetDefaultFolder](http://msdn.microsoft.com/library/namespace-getdefaultfolder-method-outlook%28Office.15%29.aspx)**, specifying an _index_ that is one of the constants in the **[OlDefaultFolders](http://msdn.microsoft.com/library/oldefaultfolders-enumeration-outlook%28Office.15%29.aspx)** enumeration to return one of the default Outlook folders in the Outlook **NameSpace** object.

 While generally it is a good practice to place items that serve the same functionality in the same folder, a folder can contain items of different types. For example, by default, the Calendar folder can contain **[AppointmentItem](appointmentitem-object-outlook.md)** and **[MeetingItem](meetingitem-object-outlook.md)** objects, and the Contacts folder can contain **[ContactItem](contactitem-object-outlook.md)** and **[DistListItem](distlistitem-object-outlook.md)** objects. In general, when enumerating items in a folder, do not assume the type of an item in the folder; check the message class of the item before accessing properties that are applicable to the item.

 Use the **[Folders.Add](http://msdn.microsoft.com/library/folders-add-method-outlook%28Office.15%29.aspx)** method to add a folder to the **Folders** object. The **Add** method has an optional argument that can be used to specify the type of items that can be stored in that folder. By default, folders created inside another folder inherit the type of the parent folder.

 Note that when items of a specific type are saved, they are saved directly into their corresponding default folder. For example, when the **[MeetingItem.GetAssociatedAppointment](http://msdn.microsoft.com/library/meetingitem-getassociatedappointment-method-outlook%28Office.15%29.aspx)** method is applied to a **MeetingItem** in the Inbox folder, the appointment that is returned will be saved to the default Calendar folder.


## Events



|**Name**|
|:-----|
|[BeforeFolderMove](http://msdn.microsoft.com/library/folder-beforefoldermove-event-outlook%28Office.15%29.aspx)|
|[BeforeItemMove](http://msdn.microsoft.com/library/folder-beforeitemmove-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddToPFFavorites](http://msdn.microsoft.com/library/folder-addtopffavorites-method-outlook%28Office.15%29.aspx)|
|[CopyTo](http://msdn.microsoft.com/library/folder-copyto-method-outlook%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/folder-delete-method-outlook%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/folder-display-method-outlook%28Office.15%29.aspx)|
|[GetCalendarExporter](http://msdn.microsoft.com/library/folder-getcalendarexporter-method-outlook%28Office.15%29.aspx)|
|[GetCustomIcon](http://msdn.microsoft.com/library/folder-getcustomicon-method-outlook%28Office.15%29.aspx)|
|[GetExplorer](http://msdn.microsoft.com/library/folder-getexplorer-method-outlook%28Office.15%29.aspx)|
|[GetStorage](http://msdn.microsoft.com/library/folder-getstorage-method-outlook%28Office.15%29.aspx)|
|[GetTable](http://msdn.microsoft.com/library/folder-gettable-method-outlook%28Office.15%29.aspx)|
|[MoveTo](http://msdn.microsoft.com/library/folder-moveto-method-outlook%28Office.15%29.aspx)|
|[SetCustomIcon](http://msdn.microsoft.com/library/folder-setcustomicon-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddressBookName](http://msdn.microsoft.com/library/folder-addressbookname-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/folder-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/folder-class-property-outlook%28Office.15%29.aspx)|
|[CurrentView](http://msdn.microsoft.com/library/folder-currentview-property-outlook%28Office.15%29.aspx)|
|[CustomViewsOnly](http://msdn.microsoft.com/library/folder-customviewsonly-property-outlook%28Office.15%29.aspx)|
|[DefaultItemType](http://msdn.microsoft.com/library/folder-defaultitemtype-property-outlook%28Office.15%29.aspx)|
|[DefaultMessageClass](http://msdn.microsoft.com/library/folder-defaultmessageclass-property-outlook%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/folder-description-property-outlook%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/folder-entryid-property-outlook%28Office.15%29.aspx)|
|[FolderPath](http://msdn.microsoft.com/library/folder-folderpath-property-outlook%28Office.15%29.aspx)|
|[Folders](folder-folders-property-outlook.md)|
|[InAppFolderSyncObject](http://msdn.microsoft.com/library/folder-inappfoldersyncobject-property-outlook%28Office.15%29.aspx)|
|[IsSharePointFolder](http://msdn.microsoft.com/library/folder-issharepointfolder-property-outlook%28Office.15%29.aspx)|
|[Items](http://msdn.microsoft.com/library/folder-items-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/folder-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/folder-parent-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/folder-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/folder-session-property-outlook%28Office.15%29.aspx)|
|[ShowAsOutlookAB](http://msdn.microsoft.com/library/folder-showasoutlookab-property-outlook%28Office.15%29.aspx)|
|[ShowItemCount](http://msdn.microsoft.com/library/folder-showitemcount-property-outlook%28Office.15%29.aspx)|
|[Store](http://msdn.microsoft.com/library/folder-store-property-outlook%28Office.15%29.aspx)|
|[StoreID](http://msdn.microsoft.com/library/folder-storeid-property-outlook%28Office.15%29.aspx)|
|[UnReadItemCount](http://msdn.microsoft.com/library/folder-unreaditemcount-property-outlook%28Office.15%29.aspx)|
|[UserDefinedProperties](http://msdn.microsoft.com/library/folder-userdefinedproperties-property-outlook%28Office.15%29.aspx)|
|[Views](http://msdn.microsoft.com/library/folder-views-property-outlook%28Office.15%29.aspx)|
|[WebViewOn](http://msdn.microsoft.com/library/folder-webviewon-property-outlook%28Office.15%29.aspx)|
|[WebViewURL](http://msdn.microsoft.com/library/folder-webviewurl-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[Folder Object Members](http://msdn.microsoft.com/library/folder-members-outlook%28Office.15%29.aspx)
