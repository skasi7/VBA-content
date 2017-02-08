---
title: JournalItem Properties (Outlook)
ms.prod: OUTLOOK
ms.assetid: fc01962f-3c0a-4a97-a633-55e509e34fdb
---


# JournalItem Properties (Outlook)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](journalitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[Application](journalitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](journalitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoResolvedWinner](journalitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](journalitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](journalitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[Categories](journalitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](journalitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](journalitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](journalitem-conflicts-property-outlook.md)|Returns the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ContactNames](journalitem-contactnames-property-outlook.md)|Returns or sets a  **String** representing the contact names associated with the Outlook item. Read/write.|
|[ConversationID](journalitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[JournalItem](journalitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](journalitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](journalitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](journalitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DocPosted](journalitem-docposted-property-outlook.md)|Returns a  **Boolean** value that indicates whether the journalized item was posted as part of the journalized session. Read/write.|
|[DocPrinted](journalitem-docprinted-property-outlook.md)|Returns a  **Boolean** value that indicates wheter the journalized item was printed as part of the journalized session. Read/write.|
|[DocRouted](journalitem-docrouted-property-outlook.md)|Returns a  **Boolean** value that indicates whether the journalized item was routed as part of the journalized session. Read/write.|
|[DocSaved](journalitem-docsaved-property-outlook.md)|Returns a  **Boolean** value that indicates whether the journalized item was saved as part of the journalized session. Read/write.|
|[DownloadState](journalitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[Duration](journalitem-duration-property-outlook.md)|Returns or sets a  **Long** indicating the duration (in minutes) of the **[JournalItem](journalitem-object-outlook.md)** Read/write.|
|[End](journalitem-end-property-outlook.md)|Returns or sets a  **Date** indicating the end date and time of a Journal entry. Read/write.|
|[EntryID](journalitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[FormDescription](journalitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](journalitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[Importance](journalitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[IsConflict](journalitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[ItemProperties](journalitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](journalitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](journalitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MessageClass](journalitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](journalitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](journalitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OutlookInternalVersion](journalitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](journalitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](journalitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](journalitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[JournalItem](journalitem-object-outlook.md)** object. Read-only.|
|[Recipients](journalitem-recipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the recipients for the Outlook item. Read-only.|
|[Saved](journalitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[Sensitivity](journalitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Session](journalitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](journalitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Start](journalitem-start-property-outlook.md)|Returns or sets a  **Date** indicating the starting date and time for the Outlook item. Read/write.|
|[Subject](journalitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[Type](journalitem-type-property-outlook.md)|Returns or sets a  **String** representing a free-form field, usually containing the display name of the journalizing application (for example, "MSWord".) Read/write.|
|[UnRead](journalitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](journalitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

