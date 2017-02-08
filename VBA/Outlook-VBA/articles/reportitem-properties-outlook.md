---
title: ReportItem Properties (Outlook)
ms.prod: OUTLOOK
ms.assetid: 81122f2f-bc98-45ae-8dcb-f9eecb3c7590
---


# ReportItem Properties (Outlook)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](reportitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[Application](reportitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](reportitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoResolvedWinner](reportitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](reportitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](reportitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[Categories](reportitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](reportitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](reportitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](reportitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ConversationID](reportitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[ReportItem](reportitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](reportitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](reportitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](reportitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DownloadState](reportitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[EntryID](reportitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[FormDescription](reportitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](reportitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[Importance](reportitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[IsConflict](reportitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[ItemProperties](reportitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](reportitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](reportitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MessageClass](reportitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](reportitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](reportitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OutlookInternalVersion](reportitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](reportitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](reportitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](reportitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[ReportItem](reportitem-object-outlook.md)** object. Read-only.|
|[RetentionExpirationDate](reportitem-retentionexpirationdate-property-outlook.md)|Returns a  **Date** that specifies the date when the **[ReportItem](reportitem-object-outlook.md)** object expires, after which the Messaging Records Management (MRM) Assistant will delete the item. Read-only.|
|[RetentionPolicyName](reportitem-retentionpolicyname-property-outlook.md)|Returns a  **String** that specifies the name of the retention policy. Read-only.|
|[Saved](reportitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[Sensitivity](reportitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Session](reportitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](reportitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Subject](reportitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[UnRead](reportitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](reportitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

