---
title: ContactItem Object (Outlook)
keywords: vbaol11.chm2992
f1_keywords:
- vbaol11.chm2992
ms.prod: OUTLOOK
api_name:
- Outlook.ContactItem
ms.assetid: 8e32093c-a678-f1fd-3f35-c2d8994d166f
---


# ContactItem Object (Outlook)

Represents a contact in a Contacts folder.


## Remarks

A contact can represent any person with whom you have any personal or professional contact.

Use the  **[CreateItem](http://msdn.microsoft.com/library/application-createitem-method-outlook%28Office.15%29.aspx)** method to create a **ContactItem** object that represents a new contact.

Use  **[Items](http://msdn.microsoft.com/library/folder-items-property-outlook%28Office.15%29.aspx)** ( _index_ ), where _index_ is the index number of a contact or a value used to match the default property of a contact, to return a single **ContactItem** object from a Contacts folder.


## Example

The following Visual Basic for Applications (VBA) example returns a new contact.


```
Set myItem = Application.CreateItem(olContactItem)
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/contactitem-afterwrite-event-outlook%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/contactitem-attachmentadd-event-outlook%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/contactitem-attachmentread-event-outlook%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/contactitem-attachmentremove-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/contactitem-beforeattachmentadd-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/contactitem-beforeattachmentpreview-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/contactitem-beforeattachmentread-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/contactitem-beforeattachmentsave-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/contactitem-beforeattachmentwritetotempfile-event-outlook%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/contactitem-beforeautosave-event-outlook%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/contactitem-beforechecknames-event-outlook%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/contactitem-beforedelete-event-outlook%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/contactitem-beforeread-event-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/contactitem-close-event-outlook%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/contactitem-customaction-event-outlook%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/contactitem-custompropertychange-event-outlook%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/contactitem-forward-event-outlook%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/contactitem-open-event-outlook%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/contactitem-propertychange-event-outlook%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/contactitem-read-event-outlook%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/contactitem-readcomplete-event-outlook%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/contactitem-reply-event-outlook%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/contactitem-replyall-event-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/contactitem-send-event-outlook%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/contactitem-unload-event-outlook%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/contactitem-write-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddBusinessCardLogoPicture](http://msdn.microsoft.com/library/contactitem-addbusinesscardlogopicture-method-outlook%28Office.15%29.aspx)|
|[AddPicture](http://msdn.microsoft.com/library/contactitem-addpicture-method-outlook%28Office.15%29.aspx)|
|[ClearTaskFlag](http://msdn.microsoft.com/library/contactitem-cleartaskflag-method-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/contactitem-close-method-outlook%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/contactitem-copy-method-outlook%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/contactitem-delete-method-outlook%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/contactitem-display-method-outlook%28Office.15%29.aspx)|
|[ForwardAsBusinessCard](http://msdn.microsoft.com/library/contactitem-forwardasbusinesscard-method-outlook%28Office.15%29.aspx)|
|[ForwardAsVcard](http://msdn.microsoft.com/library/contactitem-forwardasvcard-method-outlook%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/contactitem-getconversation-method-outlook%28Office.15%29.aspx)|
|[MarkAsTask](http://msdn.microsoft.com/library/contactitem-markastask-method-outlook%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/contactitem-move-method-outlook%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/contactitem-printout-method-outlook%28Office.15%29.aspx)|
|[RemovePicture](http://msdn.microsoft.com/library/contactitem-removepicture-method-outlook%28Office.15%29.aspx)|
|[ResetBusinessCard](http://msdn.microsoft.com/library/contactitem-resetbusinesscard-method-outlook%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/contactitem-save-method-outlook%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/contactitem-saveas-method-outlook%28Office.15%29.aspx)|
|[SaveBusinessCardImage](http://msdn.microsoft.com/library/contactitem-savebusinesscardimage-method-outlook%28Office.15%29.aspx)|
|[ShowBusinessCardEditor](http://msdn.microsoft.com/library/contactitem-showbusinesscardeditor-method-outlook%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/contactitem-showcategoriesdialog-method-outlook%28Office.15%29.aspx)|
|[ShowCheckAddressDialog](http://msdn.microsoft.com/library/contactitem-showcheckaddressdialog-method-outlook%28Office.15%29.aspx)|
|[ShowCheckFullNameDialog](http://msdn.microsoft.com/library/contactitem-showcheckfullnamedialog-method-outlook%28Office.15%29.aspx)|
|[ShowCheckPhoneDialog](http://msdn.microsoft.com/library/contactitem-showcheckphonedialog-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Account](http://msdn.microsoft.com/library/contactitem-account-property-outlook%28Office.15%29.aspx)|
|[Actions](http://msdn.microsoft.com/library/contactitem-actions-property-outlook%28Office.15%29.aspx)|
|[Anniversary](http://msdn.microsoft.com/library/contactitem-anniversary-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/contactitem-application-property-outlook%28Office.15%29.aspx)|
|[AssistantName](http://msdn.microsoft.com/library/contactitem-assistantname-property-outlook%28Office.15%29.aspx)|
|[AssistantTelephoneNumber](http://msdn.microsoft.com/library/contactitem-assistanttelephonenumber-property-outlook%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/contactitem-attachments-property-outlook%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/contactitem-autoresolvedwinner-property-outlook%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/contactitem-billinginformation-property-outlook%28Office.15%29.aspx)|
|[Birthday](http://msdn.microsoft.com/library/contactitem-birthday-property-outlook%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/contactitem-body-property-outlook%28Office.15%29.aspx)|
|[Business2TelephoneNumber](http://msdn.microsoft.com/library/contactitem-business2telephonenumber-property-outlook%28Office.15%29.aspx)|
|[BusinessAddress](http://msdn.microsoft.com/library/contactitem-businessaddress-property-outlook%28Office.15%29.aspx)|
|[BusinessAddressCity](http://msdn.microsoft.com/library/contactitem-businessaddresscity-property-outlook%28Office.15%29.aspx)|
|[BusinessAddressCountry](http://msdn.microsoft.com/library/contactitem-businessaddresscountry-property-outlook%28Office.15%29.aspx)|
|[BusinessAddressPostalCode](http://msdn.microsoft.com/library/contactitem-businessaddresspostalcode-property-outlook%28Office.15%29.aspx)|
|[BusinessAddressPostOfficeBox](http://msdn.microsoft.com/library/contactitem-businessaddresspostofficebox-property-outlook%28Office.15%29.aspx)|
|[BusinessAddressState](http://msdn.microsoft.com/library/contactitem-businessaddressstate-property-outlook%28Office.15%29.aspx)|
|[BusinessAddressStreet](http://msdn.microsoft.com/library/contactitem-businessaddressstreet-property-outlook%28Office.15%29.aspx)|
|[BusinessCardLayoutXml](http://msdn.microsoft.com/library/contactitem-businesscardlayoutxml-property-outlook%28Office.15%29.aspx)|
|[BusinessCardType](http://msdn.microsoft.com/library/contactitem-businesscardtype-property-outlook%28Office.15%29.aspx)|
|[BusinessFaxNumber](http://msdn.microsoft.com/library/contactitem-businessfaxnumber-property-outlook%28Office.15%29.aspx)|
|[BusinessHomePage](http://msdn.microsoft.com/library/contactitem-businesshomepage-property-outlook%28Office.15%29.aspx)|
|[BusinessTelephoneNumber](http://msdn.microsoft.com/library/contactitem-businesstelephonenumber-property-outlook%28Office.15%29.aspx)|
|[CallbackTelephoneNumber](http://msdn.microsoft.com/library/contactitem-callbacktelephonenumber-property-outlook%28Office.15%29.aspx)|
|[CarTelephoneNumber](http://msdn.microsoft.com/library/contactitem-cartelephonenumber-property-outlook%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/contactitem-categories-property-outlook%28Office.15%29.aspx)|
|[Children](http://msdn.microsoft.com/library/contactitem-children-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/contactitem-class-property-outlook%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/contactitem-companies-property-outlook%28Office.15%29.aspx)|
|[CompanyAndFullName](http://msdn.microsoft.com/library/contactitem-companyandfullname-property-outlook%28Office.15%29.aspx)|
|[CompanyLastFirstNoSpace](http://msdn.microsoft.com/library/contactitem-companylastfirstnospace-property-outlook%28Office.15%29.aspx)|
|[CompanyLastFirstSpaceOnly](http://msdn.microsoft.com/library/contactitem-companylastfirstspaceonly-property-outlook%28Office.15%29.aspx)|
|[CompanyMainTelephoneNumber](http://msdn.microsoft.com/library/contactitem-companymaintelephonenumber-property-outlook%28Office.15%29.aspx)|
|[CompanyName](http://msdn.microsoft.com/library/contactitem-companyname-property-outlook%28Office.15%29.aspx)|
|[ComputerNetworkName](http://msdn.microsoft.com/library/contactitem-computernetworkname-property-outlook%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/contactitem-conflicts-property-outlook%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/contactitem-conversationid-property-outlook%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/contactitem-conversationindex-property-outlook%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/contactitem-conversationtopic-property-outlook%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/contactitem-creationtime-property-outlook%28Office.15%29.aspx)|
|[CustomerID](http://msdn.microsoft.com/library/contactitem-customerid-property-outlook%28Office.15%29.aspx)|
|[Department](http://msdn.microsoft.com/library/contactitem-department-property-outlook%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/contactitem-downloadstate-property-outlook%28Office.15%29.aspx)|
|[Email1Address](http://msdn.microsoft.com/library/contactitem-email1address-property-outlook%28Office.15%29.aspx)|
|[Email1AddressType](http://msdn.microsoft.com/library/contactitem-email1addresstype-property-outlook%28Office.15%29.aspx)|
|[Email1DisplayName](http://msdn.microsoft.com/library/contactitem-email1displayname-property-outlook%28Office.15%29.aspx)|
|[Email1EntryID](http://msdn.microsoft.com/library/contactitem-email1entryid-property-outlook%28Office.15%29.aspx)|
|[Email2Address](http://msdn.microsoft.com/library/contactitem-email2address-property-outlook%28Office.15%29.aspx)|
|[Email2AddressType](http://msdn.microsoft.com/library/contactitem-email2addresstype-property-outlook%28Office.15%29.aspx)|
|[Email2DisplayName](http://msdn.microsoft.com/library/contactitem-email2displayname-property-outlook%28Office.15%29.aspx)|
|[Email2EntryID](http://msdn.microsoft.com/library/contactitem-email2entryid-property-outlook%28Office.15%29.aspx)|
|[Email3Address](http://msdn.microsoft.com/library/contactitem-email3address-property-outlook%28Office.15%29.aspx)|
|[Email3AddressType](http://msdn.microsoft.com/library/contactitem-email3addresstype-property-outlook%28Office.15%29.aspx)|
|[Email3DisplayName](http://msdn.microsoft.com/library/contactitem-email3displayname-property-outlook%28Office.15%29.aspx)|
|[Email3EntryID](http://msdn.microsoft.com/library/contactitem-email3entryid-property-outlook%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/contactitem-entryid-property-outlook%28Office.15%29.aspx)|
|[FileAs](http://msdn.microsoft.com/library/contactitem-fileas-property-outlook%28Office.15%29.aspx)|
|[FirstName](http://msdn.microsoft.com/library/contactitem-firstname-property-outlook%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/contactitem-formdescription-property-outlook%28Office.15%29.aspx)|
|[FTPSite](http://msdn.microsoft.com/library/contactitem-ftpsite-property-outlook%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/contactitem-fullname-property-outlook%28Office.15%29.aspx)|
|[FullNameAndCompany](http://msdn.microsoft.com/library/contactitem-fullnameandcompany-property-outlook%28Office.15%29.aspx)|
|[Gender](http://msdn.microsoft.com/library/contactitem-gender-property-outlook%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/contactitem-getinspector-property-outlook%28Office.15%29.aspx)|
|[GovernmentIDNumber](http://msdn.microsoft.com/library/contactitem-governmentidnumber-property-outlook%28Office.15%29.aspx)|
|[HasPicture](http://msdn.microsoft.com/library/contactitem-haspicture-property-outlook%28Office.15%29.aspx)|
|[Hobby](http://msdn.microsoft.com/library/contactitem-hobby-property-outlook%28Office.15%29.aspx)|
|[Home2TelephoneNumber](http://msdn.microsoft.com/library/contactitem-home2telephonenumber-property-outlook%28Office.15%29.aspx)|
|[HomeAddress](http://msdn.microsoft.com/library/contactitem-homeaddress-property-outlook%28Office.15%29.aspx)|
|[HomeAddressCity](http://msdn.microsoft.com/library/contactitem-homeaddresscity-property-outlook%28Office.15%29.aspx)|
|[HomeAddressCountry](http://msdn.microsoft.com/library/contactitem-homeaddresscountry-property-outlook%28Office.15%29.aspx)|
|[HomeAddressPostalCode](http://msdn.microsoft.com/library/contactitem-homeaddresspostalcode-property-outlook%28Office.15%29.aspx)|
|[HomeAddressPostOfficeBox](http://msdn.microsoft.com/library/contactitem-homeaddresspostofficebox-property-outlook%28Office.15%29.aspx)|
|[HomeAddressState](http://msdn.microsoft.com/library/contactitem-homeaddressstate-property-outlook%28Office.15%29.aspx)|
|[HomeAddressStreet](http://msdn.microsoft.com/library/contactitem-homeaddressstreet-property-outlook%28Office.15%29.aspx)|
|[HomeFaxNumber](http://msdn.microsoft.com/library/contactitem-homefaxnumber-property-outlook%28Office.15%29.aspx)|
|[HomeTelephoneNumber](http://msdn.microsoft.com/library/contactitem-hometelephonenumber-property-outlook%28Office.15%29.aspx)|
|[IMAddress](http://msdn.microsoft.com/library/contactitem-imaddress-property-outlook%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/contactitem-importance-property-outlook%28Office.15%29.aspx)|
|[Initials](http://msdn.microsoft.com/library/contactitem-initials-property-outlook%28Office.15%29.aspx)|
|[InternetFreeBusyAddress](http://msdn.microsoft.com/library/contactitem-internetfreebusyaddress-property-outlook%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/contactitem-isconflict-property-outlook%28Office.15%29.aspx)|
|[ISDNNumber](http://msdn.microsoft.com/library/contactitem-isdnnumber-property-outlook%28Office.15%29.aspx)|
|[IsMarkedAsTask](http://msdn.microsoft.com/library/contactitem-ismarkedastask-property-outlook%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/contactitem-itemproperties-property-outlook%28Office.15%29.aspx)|
|[JobTitle](http://msdn.microsoft.com/library/contactitem-jobtitle-property-outlook%28Office.15%29.aspx)|
|[Journal](http://msdn.microsoft.com/library/contactitem-journal-property-outlook%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/contactitem-language-property-outlook%28Office.15%29.aspx)|
|[LastFirstAndSuffix](http://msdn.microsoft.com/library/contactitem-lastfirstandsuffix-property-outlook%28Office.15%29.aspx)|
|[LastFirstNoSpace](http://msdn.microsoft.com/library/contactitem-lastfirstnospace-property-outlook%28Office.15%29.aspx)|
|[LastFirstNoSpaceAndSuffix](http://msdn.microsoft.com/library/contactitem-lastfirstnospaceandsuffix-property-outlook%28Office.15%29.aspx)|
|[LastFirstNoSpaceCompany](http://msdn.microsoft.com/library/contactitem-lastfirstnospacecompany-property-outlook%28Office.15%29.aspx)|
|[LastFirstSpaceOnly](http://msdn.microsoft.com/library/contactitem-lastfirstspaceonly-property-outlook%28Office.15%29.aspx)|
|[LastFirstSpaceOnlyCompany](http://msdn.microsoft.com/library/contactitem-lastfirstspaceonlycompany-property-outlook%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/contactitem-lastmodificationtime-property-outlook%28Office.15%29.aspx)|
|[LastName](http://msdn.microsoft.com/library/contactitem-lastname-property-outlook%28Office.15%29.aspx)|
|[LastNameAndFirstName](http://msdn.microsoft.com/library/contactitem-lastnameandfirstname-property-outlook%28Office.15%29.aspx)|
|[MailingAddress](http://msdn.microsoft.com/library/contactitem-mailingaddress-property-outlook%28Office.15%29.aspx)|
|[MailingAddressCity](http://msdn.microsoft.com/library/contactitem-mailingaddresscity-property-outlook%28Office.15%29.aspx)|
|[MailingAddressCountry](http://msdn.microsoft.com/library/contactitem-mailingaddresscountry-property-outlook%28Office.15%29.aspx)|
|[MailingAddressPostalCode](http://msdn.microsoft.com/library/contactitem-mailingaddresspostalcode-property-outlook%28Office.15%29.aspx)|
|[MailingAddressPostOfficeBox](http://msdn.microsoft.com/library/contactitem-mailingaddresspostofficebox-property-outlook%28Office.15%29.aspx)|
|[MailingAddressState](http://msdn.microsoft.com/library/contactitem-mailingaddressstate-property-outlook%28Office.15%29.aspx)|
|[MailingAddressStreet](http://msdn.microsoft.com/library/contactitem-mailingaddressstreet-property-outlook%28Office.15%29.aspx)|
|[ManagerName](http://msdn.microsoft.com/library/contactitem-managername-property-outlook%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/contactitem-markfordownload-property-outlook%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/contactitem-messageclass-property-outlook%28Office.15%29.aspx)|
|[MiddleName](http://msdn.microsoft.com/library/contactitem-middlename-property-outlook%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/contactitem-mileage-property-outlook%28Office.15%29.aspx)|
|[MobileTelephoneNumber](http://msdn.microsoft.com/library/contactitem-mobiletelephonenumber-property-outlook%28Office.15%29.aspx)|
|[NetMeetingAlias](http://msdn.microsoft.com/library/contactitem-netmeetingalias-property-outlook%28Office.15%29.aspx)|
|[NetMeetingServer](http://msdn.microsoft.com/library/contactitem-netmeetingserver-property-outlook%28Office.15%29.aspx)|
|[NickName](http://msdn.microsoft.com/library/contactitem-nickname-property-outlook%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/contactitem-noaging-property-outlook%28Office.15%29.aspx)|
|[OfficeLocation](http://msdn.microsoft.com/library/contactitem-officelocation-property-outlook%28Office.15%29.aspx)|
|[OrganizationalIDNumber](http://msdn.microsoft.com/library/contactitem-organizationalidnumber-property-outlook%28Office.15%29.aspx)|
|[OtherAddress](http://msdn.microsoft.com/library/contactitem-otheraddress-property-outlook%28Office.15%29.aspx)|
|[OtherAddressCity](http://msdn.microsoft.com/library/contactitem-otheraddresscity-property-outlook%28Office.15%29.aspx)|
|[OtherAddressCountry](http://msdn.microsoft.com/library/contactitem-otheraddresscountry-property-outlook%28Office.15%29.aspx)|
|[OtherAddressPostalCode](http://msdn.microsoft.com/library/contactitem-otheraddresspostalcode-property-outlook%28Office.15%29.aspx)|
|[OtherAddressPostOfficeBox](http://msdn.microsoft.com/library/contactitem-otheraddresspostofficebox-property-outlook%28Office.15%29.aspx)|
|[OtherAddressState](http://msdn.microsoft.com/library/contactitem-otheraddressstate-property-outlook%28Office.15%29.aspx)|
|[OtherAddressStreet](http://msdn.microsoft.com/library/contactitem-otheraddressstreet-property-outlook%28Office.15%29.aspx)|
|[OtherFaxNumber](http://msdn.microsoft.com/library/contactitem-otherfaxnumber-property-outlook%28Office.15%29.aspx)|
|[OtherTelephoneNumber](http://msdn.microsoft.com/library/contactitem-othertelephonenumber-property-outlook%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/contactitem-outlookinternalversion-property-outlook%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/contactitem-outlookversion-property-outlook%28Office.15%29.aspx)|
|[PagerNumber](http://msdn.microsoft.com/library/contactitem-pagernumber-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/contactitem-parent-property-outlook%28Office.15%29.aspx)|
|[PersonalHomePage](http://msdn.microsoft.com/library/contactitem-personalhomepage-property-outlook%28Office.15%29.aspx)|
|[PrimaryTelephoneNumber](http://msdn.microsoft.com/library/contactitem-primarytelephonenumber-property-outlook%28Office.15%29.aspx)|
|[Profession](http://msdn.microsoft.com/library/contactitem-profession-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/contactitem-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[RadioTelephoneNumber](http://msdn.microsoft.com/library/contactitem-radiotelephonenumber-property-outlook%28Office.15%29.aspx)|
|[ReferredBy](http://msdn.microsoft.com/library/contactitem-referredby-property-outlook%28Office.15%29.aspx)|
|[ReminderOverrideDefault](http://msdn.microsoft.com/library/contactitem-reminderoverridedefault-property-outlook%28Office.15%29.aspx)|
|[ReminderPlaySound](http://msdn.microsoft.com/library/contactitem-reminderplaysound-property-outlook%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/contactitem-reminderset-property-outlook%28Office.15%29.aspx)|
|[ReminderSoundFile](http://msdn.microsoft.com/library/contactitem-remindersoundfile-property-outlook%28Office.15%29.aspx)|
|[ReminderTime](http://msdn.microsoft.com/library/contactitem-remindertime-property-outlook%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/contactitem-rtfbody-property-outlook%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/contactitem-saved-property-outlook%28Office.15%29.aspx)|
|[SelectedMailingAddress](http://msdn.microsoft.com/library/contactitem-selectedmailingaddress-property-outlook%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/contactitem-sensitivity-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/contactitem-session-property-outlook%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/contactitem-size-property-outlook%28Office.15%29.aspx)|
|[Spouse](http://msdn.microsoft.com/library/contactitem-spouse-property-outlook%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/contactitem-subject-property-outlook%28Office.15%29.aspx)|
|[Suffix](http://msdn.microsoft.com/library/contactitem-suffix-property-outlook%28Office.15%29.aspx)|
|[TaskCompletedDate](http://msdn.microsoft.com/library/contactitem-taskcompleteddate-property-outlook%28Office.15%29.aspx)|
|[TaskDueDate](http://msdn.microsoft.com/library/contactitem-taskduedate-property-outlook%28Office.15%29.aspx)|
|[TaskStartDate](http://msdn.microsoft.com/library/contactitem-taskstartdate-property-outlook%28Office.15%29.aspx)|
|[TaskSubject](http://msdn.microsoft.com/library/contactitem-tasksubject-property-outlook%28Office.15%29.aspx)|
|[TelexNumber](http://msdn.microsoft.com/library/contactitem-telexnumber-property-outlook%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/contactitem-title-property-outlook%28Office.15%29.aspx)|
|[ToDoTaskOrdinal](http://msdn.microsoft.com/library/contactitem-todotaskordinal-property-outlook%28Office.15%29.aspx)|
|[TTYTDDTelephoneNumber](http://msdn.microsoft.com/library/contactitem-ttytddtelephonenumber-property-outlook%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/contactitem-unread-property-outlook%28Office.15%29.aspx)|
|[User1](http://msdn.microsoft.com/library/contactitem-user1-property-outlook%28Office.15%29.aspx)|
|[User2](http://msdn.microsoft.com/library/contactitem-user2-property-outlook%28Office.15%29.aspx)|
|[User3](http://msdn.microsoft.com/library/contactitem-user3-property-outlook%28Office.15%29.aspx)|
|[User4](http://msdn.microsoft.com/library/contactitem-user4-property-outlook%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/contactitem-userproperties-property-outlook%28Office.15%29.aspx)|
|[WebPage](http://msdn.microsoft.com/library/contactitem-webpage-property-outlook%28Office.15%29.aspx)|
|[YomiCompanyName](http://msdn.microsoft.com/library/contactitem-yomicompanyname-property-outlook%28Office.15%29.aspx)|
|[YomiFirstName](http://msdn.microsoft.com/library/contactitem-yomifirstname-property-outlook%28Office.15%29.aspx)|
|[YomiLastName](http://msdn.microsoft.com/library/contactitem-yomilastname-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[ContactItem Object Members](http://msdn.microsoft.com/library/contactitem-members-outlook%28Office.15%29.aspx)
