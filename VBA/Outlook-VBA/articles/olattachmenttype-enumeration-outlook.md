---
title: OlAttachmentType Enumeration (Outlook)
keywords: vbaol11.chm3052
f1_keywords:
- vbaol11.chm3052
ms.prod: OUTLOOK
api_name:
- Outlook.OlAttachmentType
ms.assetid: b6373ef7-0f30-d6c4-eb52-c6ef1de40b52
---


# OlAttachmentType Enumeration (Outlook)

Specifies the attachment type.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olByReference**|4|This value is no longer supported since Microsoft Outlook 2007. Use  **olByValue** to attach a copy of a file in the file system.|
| **olByValue**|1|The attachment is a copy of the original file and can be accessed even if the original file is removed.|
| **olEmbeddeditem**|5|The attachment is an Outlook message format file (.msg) and is a copy of the original message.|
| **olOLE**|6|The attachment is an OLE document.|

## Remarks

Used as an optional parameter to the [Attachments.Add Method (Outlook)](attachments-add-method-outlook.md) to specify the attachment type.


## See also


#### Other resources


[Attach a File to a Mail Item](http://msdn.microsoft.com/library/attach-a-file-to-a-mail-item%28Office.15%29.aspx)
[Attach an Outlook Contact Item to an Email Message](http://msdn.microsoft.com/library/attach-an-outlook-contact-item-to-an-email-message%28Office.15%29.aspx)
[Limit the Size of an Attachment to an Outlook Email Message](http://msdn.microsoft.com/library/limit-the-size-of-an-attachment-to-an-outlook-email-message%28Office.15%29.aspx)
[Modify an Attachment of an Outlook Email Message](http://msdn.microsoft.com/library/modify-an-attachment-of-an-outlook-email-message%28Office.15%29.aspx)

