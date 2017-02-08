---
title: Application Object (Outlook)
keywords: vbaol11.chm2991
f1_keywords:
- vbaol11.chm2991
ms.prod: OUTLOOK
api_name:
- Outlook.Application
ms.assetid: 797003e7-ecd1-eccb-eaaf-32d6ddde8348
---


# Application Object (Outlook)

Represents the entire Microsoft Outlook application.


## Remarks

 This is the only object in the hierarchy that can be returned by using the **[CreateObject](http://msdn.microsoft.com/library/application-createobject-method-outlook%28Office.15%29.aspx)** method or the intrinsic Visual Basic **GetObject** function.

The Outlook  **Application** object has several purposes:


- As the root object, it allows access to other objects in the Outlook hierarchy.
    
- It allows direct access to a new item created by using  **[CreateItem](http://msdn.microsoft.com/library/application-createitem-method-outlook%28Office.15%29.aspx)**, without having to traverse the object hierarchy.
    
- It allows access to the active interface objects (the explorer and the inspector).
    
When you use Automation to control Outlook from another application, you use the  **CreateObject** method to create an Outlook **Application** object.


## Example

The following Visual Basic for Applications (VBA) example starts Outlook (if it's not already running) and opens the default Inbox folder.


```
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myFolder= _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
myFolder.Display
```

The following Visual Basic for Applications (VBA) example uses the  **Application** object to create and open a new contact.




```
Set myItem = Application.CreateItem(olContactItem) 
 
myItem.Display
```


## Events



|**Name**|
|:-----|
|[AdvancedSearchComplete](http://msdn.microsoft.com/library/application-advancedsearchcomplete-event-outlook%28Office.15%29.aspx)|
|[AdvancedSearchStopped](http://msdn.microsoft.com/library/application-advancedsearchstopped-event-outlook%28Office.15%29.aspx)|
|[BeforeFolderSharingDialog](http://msdn.microsoft.com/library/application-beforefoldersharingdialog-event-outlook%28Office.15%29.aspx)|
|[ItemLoad](http://msdn.microsoft.com/library/application-itemload-event-outlook%28Office.15%29.aspx)|
|[ItemSend](http://msdn.microsoft.com/library/application-itemsend-event-outlook%28Office.15%29.aspx)|
|[MAPILogonComplete](http://msdn.microsoft.com/library/application-mapilogoncomplete-event-outlook%28Office.15%29.aspx)|
|[NewMail](http://msdn.microsoft.com/library/application-newmail-event-outlook%28Office.15%29.aspx)|
|[NewMailEx](http://msdn.microsoft.com/library/application-newmailex-event-outlook%28Office.15%29.aspx)|
|[OptionsPagesAdd](http://msdn.microsoft.com/library/application-optionspagesadd-event-outlook%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application-quit-event-outlook%28Office.15%29.aspx)|
|[Reminder](http://msdn.microsoft.com/library/application-reminder-event-outlook%28Office.15%29.aspx)|
|[Startup](http://msdn.microsoft.com/library/application-startup-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[ActiveExplorer](http://msdn.microsoft.com/library/application-activeexplorer-method-outlook%28Office.15%29.aspx)|
|[ActiveInspector](http://msdn.microsoft.com/library/application-activeinspector-method-outlook%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/application-activewindow-method-outlook%28Office.15%29.aspx)|
|[AdvancedSearch](http://msdn.microsoft.com/library/application-advancedsearch-method-outlook%28Office.15%29.aspx)|
|[CopyFile](http://msdn.microsoft.com/library/application-copyfile-method-outlook%28Office.15%29.aspx)|
|[CreateItem](http://msdn.microsoft.com/library/application-createitem-method-outlook%28Office.15%29.aspx)|
|[CreateItemFromTemplate](http://msdn.microsoft.com/library/application-createitemfromtemplate-method-outlook%28Office.15%29.aspx)|
|[CreateObject](http://msdn.microsoft.com/library/application-createobject-method-outlook%28Office.15%29.aspx)|
|[GetNamespace](http://msdn.microsoft.com/library/application-getnamespace-method-outlook%28Office.15%29.aspx)|
|[GetObjectReference](http://msdn.microsoft.com/library/application-getobjectreference-method-outlook%28Office.15%29.aspx)|
|[IsSearchSynchronous](http://msdn.microsoft.com/library/application-issearchsynchronous-method-outlook%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application-quit-method-outlook%28Office.15%29.aspx)|
|[RefreshFormRegionDefinition](http://msdn.microsoft.com/library/application-refreshformregiondefinition-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/application-application-property-outlook%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/application-assistance-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/application-class-property-outlook%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/application-comaddins-property-outlook%28Office.15%29.aspx)|
|[DefaultProfileName](http://msdn.microsoft.com/library/application-defaultprofilename-property-outlook%28Office.15%29.aspx)|
|[Explorers](http://msdn.microsoft.com/library/application-explorers-property-outlook%28Office.15%29.aspx)|
|[Inspectors](http://msdn.microsoft.com/library/application-inspectors-property-outlook%28Office.15%29.aspx)|
|[IsTrusted](http://msdn.microsoft.com/library/application-istrusted-property-outlook%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/application-languagesettings-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/application-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/application-parent-property-outlook%28Office.15%29.aspx)|
|[PickerDialog](http://msdn.microsoft.com/library/application-pickerdialog-property-outlook%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/application-productcode-property-outlook%28Office.15%29.aspx)|
|[Reminders](http://msdn.microsoft.com/library/application-reminders-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/application-session-property-outlook%28Office.15%29.aspx)|
|[TimeZones](http://msdn.microsoft.com/library/application-timezones-property-outlook%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/application-version-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources

<<<<<<< HEAD
=======


>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)

