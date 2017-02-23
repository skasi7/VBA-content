---
title: Application Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 3519c89c-2353-85ee-7ddc-62e5dd85a8e7
---


# Application Members (Outlook)
Represents the entire Microsoft Outlook application.

Represents the entire Microsoft Outlook application.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AdvancedSearchComplete](application-advancedsearchcomplete-event-outlook.md)|Occurs when the  **[AdvancedSearch](application-advancedsearch-method-outlook.md)** method has completed.|
|[AdvancedSearchStopped](application-advancedsearchstopped-event-outlook.md)|Occurs when a specified  **[Search](search-object-outlook.md)** object's **[Stop](search-stop-method-outlook.md)** method has been executed.|
|[BeforeFolderSharingDialog](application-beforefoldersharingdialog-event-outlook.md)|Occurs before the  **Sharing** dialog box is displayed for a selected **[Folder](folder-object-outlook.md)** object.|
|[ItemLoad](application-itemload-event-outlook.md)|Occurs when an Outlook item is loaded into memory.|
|[ItemSend](application-itemsend-event-outlook.md)|Occurs whenever an Microsoft Outlook item is sent, either by the user through an  **[Inspector](inspector-object-outlook.md)** (before the inspector is closed, but after the user clicks the **Send** button) or when the **[Send](mailitem-send-method-outlook.md)** method for an Outlook item, such as **[MailItem](mailitem-object-outlook.md)** , is used in a program.|
|[MAPILogonComplete](application-mapilogoncomplete-event-outlook.md)|Occurs after the user has logged onto the system.|
|[NewMail](application-newmail-event-outlook.md)|Occurs when one or more new e-mail messages are received in the  **Inbox**. |
|[NewMailEx](application-newmailex-event-outlook.md)|Occurs when a new item is received in the Inbox.|
|[OptionsPagesAdd](application-optionspagesadd-event-outlook.md)|Occurs whenever the user clicks the  **Add-in Options** button on the **Add-ins** tab of the Outlook **Options** dialog box.|
|[Quit](application-quit-event-outlook.md)|Occurs when Microsoft Outlook begins to close. |
|[Reminder](application-reminder-event-outlook.md)|Occurs immediately before a reminder is displayed.|
|[Startup](application-startup-event-outlook.md)|Occurs when Microsoft Outlook is starting, but after all add-in programs have been loaded. |

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ActiveExplorer](application-activeexplorer-method-outlook.md)|Returns the topmost  **[Explorer](explorer-object-outlook.md)** object on the desktop.|
|[ActiveInspector](application-activeinspector-method-outlook.md)|Returns the topmost  **[Inspector](inspector-object-outlook.md)** object on the desktop.|
|[ActiveWindow](application-activewindow-method-outlook.md)|Returns an object representing the current Microsoft Outlook window on the desktop, either an  **[Explorer](explorer-object-outlook.md)** or an **[Inspector](inspector-object-outlook.md)** object.|
|[AdvancedSearch](application-advancedsearch-method-outlook.md)|Performs a search based on a specified DAV Searching and Locating (DASL) search string.|
|[CopyFile](application-copyfile-method-outlook.md)|Copies a file from a specified location into a Microsoft Outlook store.|
|[CreateItem](application-createitem-method-outlook.md)|Creates and returns a new Microsoft Outlook item.|
|[CreateItemFromTemplate](application-createitemfromtemplate-method-outlook.md)|Creates a new Microsoft Outlook item from an Outlook template (.oft) and returns the new item.|
|[CreateObject](application-createobject-method-outlook.md)|Creates an automation object of the specified class.|
|[GetNamespace](application-getnamespace-method-outlook.md)|Returns a  **[NameSpace](namespace-object-outlook.md)** object of the specified type.|
|[GetObjectReference](application-getobjectreference-method-outlook.md)|Creates a strong or weak object reference for a specified Outlook object.|
|[IsSearchSynchronous](application-issearchsynchronous-method-outlook.md)|Returns a  **Boolean** indicating if a search will be synchronous or asynchronous.|
|[Quit](application-quit-method-outlook.md)|Closes all currently open windows. |
|[RefreshFormRegionDefinition](application-refreshformregiondefinition-method-outlook.md)|Refreshes the cache by obtaining the current definition from the Windows registry for one or all of the form regions that are defined for the local machine and the current user.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](application-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Assistance](application-assistance-property-outlook.md)|Returns an  **[IAssistance](iassistance-object-office.md)** object used to invoke help. Read-only.|
|[Class](application-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[COMAddIns](application-comaddins-property-outlook.md)|Returns a  **COMAddIns** collection that represents all the Component Object Model (COM) add-ins currently loaded in Microsoft Outlook.|
|[DefaultProfileName](application-defaultprofilename-property-outlook.md)|Returns a  **String** representing the name of the default profile name. Read-only.|
|[Explorers](application-explorers-property-outlook.md)|Returns an  **[Explorers](explorers-object-outlook.md)** collection object that contains the **[Explorer](explorer-object-outlook.md)** objects representing all open explorers. Read-only.|
|[Inspectors](application-inspectors-property-outlook.md)|Returns an  **[Inspectors](inspectors-object-outlook.md)** collection object that contains the **[Inspector](inspector-object-outlook.md)** objects representing all open inspectors. Read-only.|
|[IsTrusted](application-istrusted-property-outlook.md)|Returns a  **Boolean** to indicate if an add-in or external caller is considered trusted by Outlook. Read-only|
|[LanguageSettings](application-languagesettings-property-outlook.md)|Returns a  **[LanguageSettings](languagesettings-object-office.md)** object for the application that contains the language-specific attributes of Outlook. Read-only.|
|[Name](application-name-property-outlook.md)|Returns a  **String** value that represents the display name for the object. Read-only.|
|[Parent](application-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PickerDialog](application-pickerdialog-property-outlook.md)|Returns a  **[PickerDialog](pickerdialog-object-office.md)** object that provides the functionality to select people or data in a dialog box. Read-only.|
|[ProductCode](application-productcode-property-outlook.md)|Returns a  **String** specifying the Microsoft Outlook globally unique identifier (GUID). Read-only.|
|[Reminders](application-reminders-property-outlook.md)|Returns a  **[Reminders](reminders-object-outlook.md)** collection that represents all current reminders. Read-only.|
|[Session](application-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[TimeZones](application-timezones-property-outlook.md)|Returns a  **[TimeZones](timezones-object-outlook.md)** collection that represents the set of time zones supported by Outlook. Read-only.|
|[Version](application-version-property-outlook.md)|Returns or sets a  **String** indicating the number of the version. Read-only.|

