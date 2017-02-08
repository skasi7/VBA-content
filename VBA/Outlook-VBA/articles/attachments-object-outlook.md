---
title: Attachments Object (Outlook)
keywords: vbaol11.chm169
f1_keywords:
- vbaol11.chm169
ms.prod: OUTLOOK
api_name:
- Outlook.Attachments
ms.assetid: 4cc96a5f-a822-8ad5-6f61-e996bee8ba22
---


# Attachments Object (Outlook)

Contains a set of  **[Attachment](http://msdn.microsoft.com/library/attachment-object-outlook%28Office.15%29.aspx)** objects that represent the attachments in an Outlook item.


## Remarks

Use the  **[Attachments](http://msdn.microsoft.com/library/attachments-item-method-outlook%28Office.15%29.aspx)** property to return the **Attachments** collection for any Outlook item (except notes).

Use the  **[Add](http://msdn.microsoft.com/library/attachments-add-method-outlook%28Office.15%29.aspx)** method to add an attachment to an item.

To ensure consistent results, always save an item before adding or removing objects in the  **Attachments** collection of the item.


## Example

The following Visual Basic for Applications (VBA) example creates a new mail message, attaches a Q496.xls as an attachment (not a link), and gives the attachment a descriptive caption.


```
Set myItem = Application.CreateItem(olMailItem) 
 
myItem.Save 
 
Set myAttachments = myItem.Attachments 
 
myAttachments.Add "C:\My Documents\Q496.xls", _ 
 
 olByValue, 1, "4th Quarter 1996 Results Chart"
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/attachments-add-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/attachments-item-method-outlook%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/attachments-remove-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/attachments-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/attachments-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/attachments-count-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/attachments-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/attachments-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Attach a File to a Mail Item](http://msdn.microsoft.com/library/attach-a-file-to-a-mail-item%28Office.15%29.aspx)
[Attach an Outlook Contact Item to an Email Message](http://msdn.microsoft.com/library/attach-an-outlook-contact-item-to-an-email-message%28Office.15%29.aspx)
[Limit the Size of an Attachment to an Outlook Email Message](http://msdn.microsoft.com/library/limit-the-size-of-an-attachment-to-an-outlook-email-message%28Office.15%29.aspx)
[Modify an Attachment of an Outlook Email Message](http://msdn.microsoft.com/library/modify-an-attachment-of-an-outlook-email-message%28Office.15%29.aspx)
[Attachments Object Members](http://msdn.microsoft.com/library/attachments-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
