---
title: SelectNamesDialog Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 0f5546af-f89a-8a8b-ced9-a2d646bf9634
---


# SelectNamesDialog Members (Outlook)
Displays the  **Select Names** dialog box for the user to select entries from one or more address lists, and returns the selected entries in the collection object specified by the property **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)** .

Displays the  **Select Names** dialog box for the user to select entries from one or more address lists, and returns the selected entries in the collection object specified by the property **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)** .


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Display](selectnamesdialog-display-method-outlook.md)|Displays the  **Select Names** dialog box.|
|[SetDefaultDisplayMode](selectnamesdialog-setdefaultdisplaymode-method-outlook.md)|Sets the default display mode for the  **Select Names** dialog box, specifying its caption and button labels.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowMultipleSelection](selectnamesdialog-allowmultipleselection-property-outlook.md)|Returns or sets a  **Boolean** that determines whether more than one address entry can be selected at a time in the **Select Names** dialog. Read/write.|
|[Application](selectnamesdialog-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent application (Outlook) for the **[SelectNamesDialog](selectnamesdialog-object-outlook.md)** object. Read-only.|
|[BccLabel](selectnamesdialog-bcclabel-property-outlook.md)|Returns or sets a  **String** for the text that appears on the **Bcc** command button on the **Select Names** dialog box. Read/write.|
|[Caption](selectnamesdialog-caption-property-outlook.md)|Returns or sets a  **String** value that represents the title for the **Select Names** dialog box. Read/write.|
|[CcLabel](selectnamesdialog-cclabel-property-outlook.md)|Returns or sets a  **String** for the text that appears on the **Cc** command button on the **Select Names** dialog box. Read/write.|
|[Class](selectnamesdialog-class-property-outlook.md)|Returns a constant in the  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** enumeration indicating the class of the **[SelectNamesDialog](selectnamesdialog-object-outlook.md)** object. Read-only.|
|[ForceResolution](selectnamesdialog-forceresolution-property-outlook.md)|Returns or sets a  **Boolean** that determines if Outlook must resolve all recipients in the object specified by **[SelectNamesDialog.Recipients](selectnamesdialog-recipients-property-outlook.md)** before the user can click **OK** to accept the typed or selected recipients in the **Select Names** dialog box. Read/write.|
|[InitialAddressList](selectnamesdialog-initialaddresslist-property-outlook.md)|Returns or sets an  **[AddressList](addresslist-object-outlook.md)** object that determines the initial address list to be displayed in the **Select Names** dialog box. Read/write.|
|[NumberOfRecipientSelectors](selectnamesdialog-numberofrecipientselectors-property-outlook.md)|Returns or sets a  **[OlRecipientSelectors](olrecipientselectors-enumeration-outlook.md)** constant that determines the number of recipient edit boxes (each associated with a command button) displayed in the **Select Names** dialog box. Read/write.|
|[Parent](selectnamesdialog-parent-property-outlook.md)|Returns the parent object of the  **[SelectNamesDialog](selectnamesdialog-object-outlook.md)** object. Read-only.|
|[Recipients](selectnamesdialog-recipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection object that represents the recipients selected in the **Select Names** dialog, or sets a **Recipients** collection object that represents the initial recipients to be displayed in the **Select Names** dialog box. Read/write.|
|[Session](selectnamesdialog-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[ShowOnlyInitialAddressList](selectnamesdialog-showonlyinitialaddresslist-property-outlook.md)|Returns or sets a  **Boolean** that determines if the **[AddressList](addresslist-object-outlook.md)** represented by **[SelectNamesDialog.InitialAddressList](selectnamesdialog-initialaddresslist-property-outlook.md)** is the only **AddressList** available in the drop-down list for **Address Book** in the **Select Names** dialog box. Read/write.|
|[ToLabel](selectnamesdialog-tolabel-property-outlook.md)|Returns or sets a  **String** for the text that appears on the **To** command button on the **Select Names** dialog box. Read/write.|

