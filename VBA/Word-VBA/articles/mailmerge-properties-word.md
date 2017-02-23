---
title: MailMerge Properties (Word)
ms.prod: WORD
ms.assetid: 39df68bf-0ce3-41c6-a3b7-5671cec0b831
---


# MailMerge Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](mailmerge-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](mailmerge-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DataSource](mailmerge-datasource-property-word.md)|Returns a  **[MailMergeDataSource](mailmergedatasource-object-word.md)** object that refers to the data source attached to a mail merge main document. Read-only.|
|[Destination](mailmerge-destination-property-word.md)|Returns or sets the destination of the mail merge results. Read/write  **WdMailMergeDestination** .|
|[Fields](mailmerge-fields-property-word.md)|Returns a read-only  **MailMergeFields** collection that represents all the mail merge fields in the specified document.|
|[HighlightMergeFields](mailmerge-highlightmergefields-property-word.md)| **True** to highlight the merge fields in a document. Read/write **Boolean** .|
|[MailAddressFieldName](mailmerge-mailaddressfieldname-property-word.md)|Returns or sets the name of the field that contains e-mail addresses that are used when the mail merge destination is electronic mail. Read/write  **String** .|
|[MailAsAttachment](mailmerge-mailasattachment-property-word.md)| **True** if the merge documents are sent as attachments when the mail merge destination is an e-mail message or a fax. Read/write **Boolean** .|
|[MailFormat](mailmerge-mailformat-property-word.md)|Returns a  **WdMailMergeMailFormat** constant that represents the format to use when the mail merge destination is an e-mail message. Read/write.|
|[MailSubject](mailmerge-mailsubject-property-word.md)|Returns or sets the subject line used when the mail merge destination is electronic mail. Read/write  **String** .|
|[MainDocumentType](mailmerge-maindocumenttype-property-word.md)|Returns or sets the mail merge main document type. Read/write  **WdMailMergeMainDocType** .|
|[Parent](mailmerge-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **MailMerge** object.|
|[ShowSendToCustom](mailmerge-showsendtocustom-property-word.md)|Returns or sets a  **String** corresponding to the caption on a custom button on the Complete the merge step (step six) of the Mail Merge Wizard. Read/write.|
|[State](mailmerge-state-property-word.md)|Returns the current state of a mail merge operation. Read-only  **WdMailMergeState** .|
|[SuppressBlankLines](mailmerge-suppressblanklines-property-word.md)| **True** if blank lines are suppressed when mail merge fields in a mail merge main document are empty. Read/write **Boolean** .|
|[ViewMailMergeFieldCodes](mailmerge-viewmailmergefieldcodes-property-word.md)| **True** if merge field names are displayed in a mail merge main document. **False** if information from the current record is displayed. Read/write **Long** .|
|[WizardState](mailmerge-wizardstate-property-word.md)|Returns or sets a  **Long** indicating the current Mail Merge Wizard step for a document. The WizardState method returns a number that equates to the current Mail Merge Wizard step; a zero (0) means the Mail Merge Wizard is closed. Read/write.|

