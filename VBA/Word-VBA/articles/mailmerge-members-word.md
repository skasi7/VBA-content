---
title: MailMerge Members (Word)
ms.prod: WORD
ms.assetid: b4db0f00-0f03-4162-7312-b3aa417bea03
---


# MailMerge Members (Word)
Represents the mail merge functionality in Word.

Represents the mail merge functionality in Word.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Check](mailmerge-check-method-word.md)|Simulates the mail merge operation, pausing to report each error as it occurs.|
|[CreateDataSource](mailmerge-createdatasource-method-word.md)|Creates a Microsoft Word document that uses a table to store data for a mail merge.|
|[CreateHeaderSource](mailmerge-createheadersource-method-word.md)|Creates a Microsoft Word document that stores a header record that is used instead of the data source header record in a mail merge.|
|[EditDataSource](mailmerge-editdatasource-method-word.md)|Opens or switches to the mail merge data source.|
|[EditHeaderSource](mailmerge-editheadersource-method-word.md)|Opens the header source attached to a mail merge main document, or activates the header source if it is already open.|
|[EditMainDocument](mailmerge-editmaindocument-method-word.md)|Activates the mail merge main document associated with the specified header source or data source document.|
|[Execute](mailmerge-execute-method-word.md)|Performs the specified mail merge operation.|
|[OpenDataSource](mailmerge-opendatasource-method-word.md)|Attaches a data source to the specified document, which becomes a main document if it is not one already.|
|[OpenHeaderSource](mailmerge-openheadersource-method-word.md)|Attaches a mail merge header source to the specified document.|
|[ShowWizard](mailmerge-showwizard-method-word.md)|Displays the Mail Merge Wizard in a document.|

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

