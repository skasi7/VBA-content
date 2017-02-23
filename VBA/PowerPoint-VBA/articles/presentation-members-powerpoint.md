---
title: Presentation Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: b3538c7e-5fd9-d34d-ab5c-0105dbd516d0
---


# Presentation Members (PowerPoint)
Represents a Microsoft PowerPoint presentation. 

Represents a Microsoft PowerPoint presentation. 


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AcceptAll](presentation-acceptall-method-powerpoint.md)|Accepts all changes.|
|[AddTitleMaster](presentation-addtitlemaster-method-powerpoint.md)|Adds a title master to the specified presentation and returns a  **[Master](master-object-powerpoint.md)** object that represents the title master.|
|[AddToFavorites](presentation-addtofavorites-method-powerpoint.md)|Adds a shortcut that represents the current selection in the specified presentation to the Windows Favorites folder.|
|[ApplyTemplate](presentation-applytemplate-method-powerpoint.md)|Applies a design template to the specified presentation.|
|[ApplyTemplate2](presentation-applytemplate2-method-powerpoint.md)|Applies a design template and theme variant to the presentation.|
|[ApplyTheme](presentation-applytheme-method-powerpoint.md)|Applies a theme or design template to the specified presentation.|
|[CanCheckIn](presentation-cancheckin-method-powerpoint.md)|Returns  **True** if Microsoft PowerPoint can check in a specified presentation to a server.|
|[CheckIn](presentation-checkin-method-powerpoint.md)|Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.|
|[CheckInWithVersion](presentation-checkinwithversion-method-powerpoint.md)|Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.|
|[Close](presentation-close-method-powerpoint.md)|Closes the specified presentation.|
|[Convert2](presentation-convert2-method-powerpoint.md)|Converts a file to a different file type.|
|[CreateVideo](presentation-createvideo-method-powerpoint.md)|Creates a video in a  **Presentation** object.|
|[EndReview](presentation-endreview-method-powerpoint.md)|Ends the review cycle.|
|[EnsureAllMediaUpgraded](presentation-ensureallmediaupgraded-method-powerpoint.md)|Ensures that all media is up to date in a  **Presentation** object.|
|[Export](presentation-export-method-powerpoint.md)|Exports each slide in the presentation, using the specified graphics filter, and saves the exported files in the specified folder.|
|[ExportAsFixedFormat](presentation-exportasfixedformat-method-powerpoint.md)|Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.|
|[ExportAsFixedFormat2](presentation-exportasfixedformat2-method-powerpoint.md)|Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.|
|[FollowHyperlink](presentation-followhyperlink-method-powerpoint.md)|Displays a cached document, if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document and displays it in the appropriate application.|
|[GetWorkflowTasks](presentation-getworkflowtasks-method-powerpoint.md)| Returns the Microsoft Office **WorkflowTasks** collection.|
|[GetWorkflowTemplates](presentation-getworkflowtemplates-method-powerpoint.md)|Returns the Microsoft Office  **WorkflowTemplates** collection.|
|[LockServerFile](presentation-lockserverfile-method-powerpoint.md)|Locks the presentation on the Microsoft Office SharePoint server to prevent its modification.|
|[Merge](presentation-merge-method-powerpoint.md)|Merges the changes in one presentation with another.|
|[MergeWithBaseline](presentation-mergewithbaseline-method-powerpoint.md)|Merges a presentation into another presentation.|
|[NewWindow](presentation-newwindow-method-powerpoint.md)| Opens a new window that contains the specified presentation. Returns a **[DocumentWindow](documentwindow-object-powerpoint.md)** object that represents the new window.|
|[PrintOut](presentation-printout-method-powerpoint.md)|Prints the specified presentation.|
|[PublishSlides](presentation-publishslides-method-powerpoint.md)|Creates a Web presentation (in HTML format) containing slides from any loaded presentation. You can view the published presentation in a Web browser.|
|[RejectAll](presentation-rejectall-method-powerpoint.md)|Rejects all changes.|
|[RemoveDocumentInformation](presentation-removedocumentinformation-method-powerpoint.md)|Removes document information, such as personal information, comments, and document properties, from a Microsoft PowerPoint presentation.|
|[Save](presentation-save-method-powerpoint.md)|Saves the specified presentation.|
|[SaveAs](presentation-saveas-method-powerpoint.md)|Saves a presentation that's never been saved, or saves a previously saved presentation under a different name.|
|[SaveCopyAs](presentation-savecopyas-method-powerpoint.md)|Saves a copy of the specified presentation to a file without modifying the original.|
|[SendFaxOverInternet](presentation-sendfaxoverinternet-method-powerpoint.md)|Sends a presentation as a fax to the specified recipients.|
|[SetPasswordEncryptionOptions](presentation-setpasswordencryptionoptions-method-powerpoint.md)|Sets the options Microsoft PowerPoint uses for encrypting presentations with passwords.|
|[UpdateLinks](presentation-updatelinks-method-powerpoint.md)|Updates linked OLE objects in the specified presentation.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](presentation-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[Broadcast](presentation-broadcast-property-powerpoint.md)|Returns the  **Broadcast** object of the current **Presentation** object. Read-only.|
|[BuiltInDocumentProperties](presentation-builtindocumentproperties-property-powerpoint.md)|Returns a  **DocumentProperties** collection that represents all the built-in document properties for the specified presentation. Read-only.|
|[ChartDataPointTrack](presentation-chartdatapointtrack-property-powerpoint.md)|Returns or sets a  **Boolean** that specifies whether charts use cell-reference data-point tracking. Read-write.|
|[Coauthoring](presentation-coauthoring-property-powerpoint.md)|Returns a  **Coauthoring** object in the current **Presentation** object. Read-only.|
|[ColorSchemes](presentation-colorschemes-property-powerpoint.md)|Returns a  **[ColorSchemes](colorschemes-object-powerpoint.md)** collection that represents the color schemes in the specified presentation. Read-only.|
|[CommandBars](presentation-commandbars-property-powerpoint.md)|Returns a  **CommandBars** collection that represents the merged command bar set from the host container application and Microsoft PowerPoint. This property returns a valid object only when the container is a DocObject server, like Microsoft Binder, and PowerPoint is acting as an OLE server. Read-only.|
|[Container](presentation-container-property-powerpoint.md)|Returns the object that contains the specified embedded presentation. Read-only.|
|[ContentTypeProperties](presentation-contenttypeproperties-property-powerpoint.md)|Returns the Microsoft Office  **MetaProperties** collection that describes the metadata stored in the presentation. Read-only.|
|[CreateVideoStatus](presentation-createvideostatus-property-powerpoint.md)|Returns the status of creating a video in the current  **Presentation**. Read-only.|
|[CustomDocumentProperties](presentation-customdocumentproperties-property-powerpoint.md)|Returns a  **DocumentProperties** collection that represents all the custom document properties for the specified presentation. Read-only.|
|[CustomerData](presentation-customerdata-property-powerpoint.md)|Returns a  **CustomerData** object. Read-only.|
|[CustomXMLParts](presentation-customxmlparts-property-powerpoint.md)|Returns a  **[CustomXMLParts](customxmlparts-object-office.md)** object that represents the collection of custom XML parts associated with the specified **Presentation** object. Read-only.|
|[DefaultLanguageID](presentation-defaultlanguageid-property-powerpoint.md)|Returns or sets the default language of a presentation. Read/write.|
|[DefaultShape](presentation-defaultshape-property-powerpoint.md)|Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the default shape for the presentation. Read-only.|
|[Designs](presentation-designs-property-powerpoint.md)|Returns a  **[Designs](designs-object-powerpoint.md)** object that represents a collection of designs.|
|[DisplayComments](presentation-displaycomments-property-powerpoint.md)|Determines whether comments are displayed in the specified presentation. Read/write.|
|[DocumentInspectors](presentation-documentinspectors-property-powerpoint.md)|Returns the Microsoft Office  **[DocumentInspectors](documentinspectors-object-office.md)** collection. Read-only.|
|[DocumentLibraryVersions](presentation-documentlibraryversions-property-powerpoint.md)|Returns a  **DocumentLibraryVersions** collection that represents the collection of versions of a shared presentation that has versioning enabled and that is stored in a document library on a server.|
|[EncryptionProvider](presentation-encryptionprovider-property-powerpoint.md)|Returns a  **String** that specifies the name of the algorithm encryption provider that PowerPoint uses when encrypting documents. Read/write.|
|[EnvelopeVisible](presentation-envelopevisible-property-powerpoint.md)|Determines whether the e-mail message header is visible in the document window. Read/write.|
|[ExtraColors](presentation-extracolors-property-powerpoint.md)|Returns an  **[ExtraColors](extracolors-object-powerpoint.md)** object that represents the extra colors available in the specified presentation. Read-only.|
|[FarEastLineBreakLanguage](presentation-fareastlinebreaklanguage-property-powerpoint.md)|Returns or sets the language used to determine which line break level is used when the line break control option is turned on. Read/write.|
|[FarEastLineBreakLevel](presentation-fareastlinebreaklevel-property-powerpoint.md)|Returns or sets the line break based upon Asian character level. Read/write.|
|[Final](presentation-final-property-powerpoint.md)|Determines whether the presentation is marked as final (read-only). Read/write.|
|[Fonts](presentation-fonts-property-powerpoint.md)|Returns a  **[Fonts](fonts-object-powerpoint.md)** collection that represents all fonts used in the specified presentation. Read-only.|
|[FullName](presentation-fullname-property-powerpoint.md)|Returns the name of the specified add-in or saved presentation, including the path, the current file system separator, and the file name extension. Read-only  **String**.|
|[GridDistance](presentation-griddistance-property-powerpoint.md)|Sets or returns a  **Single** that represents the distance between gridlines. Read/write.|
|[Guides](presentation-guides-property-powerpoint.md)|Returns the [Guides](923ef616-0670-6ad1-2d9b-f7fe4642185b.md) collection associated with a custom layout. Read-only.|
|[HandoutMaster](presentation-handoutmaster-property-powerpoint.md)|Returns a  **[Master](master-object-powerpoint.md)** object that represents the handout master. Read-only.|
|[HasHandoutMaster](presentation-hashandoutmaster-property-powerpoint.md)|Indicates whether the presentation has media that resides on a handout master. Read-only|
|[HasNotesMaster](presentation-hasnotesmaster-property-powerpoint.md)|Indicates whether the presentation has media that resides on a notes master. Read-only.|
|[HasTitleMaster](presentation-hastitlemaster-property-powerpoint.md)|**MsoTrue** if the specified presentation has a title master. Read-only.|
|[HasVBProject](presentation-hasvbproject-property-powerpoint.md)|Returns whether the active presentation contains a Microsoft Visual Basic for Applications (VBA) project. Read-only.|
|[InMergeMode](presentation-inmergemode-property-powerpoint.md)|Indicates whether the document window is in merge mode. Read-only|
|[LayoutDirection](presentation-layoutdirection-property-powerpoint.md)|Returns or sets the layout direction for the user interface. Read/write.|
|[Name](presentation-name-property-powerpoint.md)|The name of the presentation includes the file name extension (for file types that are registered) but doesn't include its path. You cannot use this property to set the name. Use the  **[SaveAs](presentation-saveas-method-powerpoint.md)** method to save the presentation under a different name if you need to change the name. Read-only.|
|[NoLineBreakAfter](presentation-nolinebreakafter-property-powerpoint.md)|Returns or sets the characters that cannot end a line. Read/write.|
|[NoLineBreakBefore](presentation-nolinebreakbefore-property-powerpoint.md)|Returns or sets the characters that cannot begin a line. Read/write.|
|[NotesMaster](presentation-notesmaster-property-powerpoint.md)|Returns a  **[Master](master-object-powerpoint.md)** object that represents the notes master. Read-only.|
|[PageSetup](presentation-pagesetup-property-powerpoint.md)|Returns a  **[PageSetup](pagesetup-object-powerpoint.md)** object whose properties control slide setup attributes for the specified presentation. Read-only.|
|[Parent](presentation-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[Password](presentation-password-property-powerpoint.md)|Returns or sets the password that must be supplied to open the specified presentation. Read/write.|
|[PasswordEncryptionAlgorithm](presentation-passwordencryptionalgorithm-property-powerpoint.md)|Returns the algorithm Microsoft PowerPoint uses for encrypting documents with passwords. Read-only.|
|[PasswordEncryptionFileProperties](presentation-passwordencryptionfileproperties-property-powerpoint.md)|Returns whether Microsoft PowerPoint encrypts file properties for password-protected documents. Read-only.|
|[PasswordEncryptionKeyLength](presentation-passwordencryptionkeylength-property-powerpoint.md)|Returns the key length of the algorithm Microsoft PowerPoint uses when it encrypts documents with passwords. Read-only.|
|[PasswordEncryptionProvider](presentation-passwordencryptionprovider-property-powerpoint.md)|Returns the name of the algorithm encryption provider that Microsoft PowerPoint uses when it encrypts documents with passwords. Read-only.|
|[Path](presentation-path-property-powerpoint.md)|Returns a  **String** that represents the path to the specified **[Presentation](presentation-object-powerpoint.md)** object. Read-only.|
|[Permission](presentation-permission-property-powerpoint.md)||
|[PrintOptions](presentation-printoptions-property-powerpoint.md)|Returns a  **[PrintOptions](printoptions-object-powerpoint.md)** object that represents print options that are saved with the specified presentation. Read-only.|
|[ReadOnly](presentation-readonly-property-powerpoint.md)|Returns whether the specified presentation is read-only. Read-only.|
|[RemovePersonalInformation](presentation-removepersonalinformation-property-powerpoint.md)|Determines whether Microsoft PowerPoint should remove all user information from comments and revisions upon saving a presentation. Read/write.|
|[Research](presentation-research-property-powerpoint.md)|Returns a  **Research** object that provides access to the research service feature of Microsoft PowerPoint. Read-only.|
|[Saved](presentation-saved-property-powerpoint.md)|Determines whether changes have been made to a presentation since it was last saved. Read/write.|
|[SectionProperties](presentation-sectionproperties-property-powerpoint.md)|Returns a  **SectionProperties** object. Read-only.|
|[ServerPolicy](presentation-serverpolicy-property-powerpoint.md)|Returns a Microsoft Office  **[ServerPolicy](serverpolicy-object-office.md)** object. Read-only.|
|[SharedWorkspace](presentation-sharedworkspace-property-powerpoint.md)|Returns a  **SharedWorkspace** object that represents the Document Workspace in which a specified presentation is located. Read-only.|
|[Signatures](presentation-signatures-property-powerpoint.md)|Returns a  **SignatureSet** object that represents a collection of digital signatures. Read-only.|
|[SlideMaster](presentation-slidemaster-property-powerpoint.md)|Returns a  **[Master](master-object-powerpoint.md)** object that represents the slide master.|
|[Slides](presentation-slides-property-powerpoint.md)|Returns a  **[Slides](slides-object-powerpoint.md)** collection that represents all slides in the specified presentation. Read-only.|
|[SlideShowSettings](presentation-slideshowsettings-property-powerpoint.md)|Returns a  **[SlideShowSettings](slideshowsettings-object-powerpoint.md)** object that represents the slide show settings for the specified presentation. Read-only.|
|[SlideShowWindow](presentation-slideshowwindow-property-powerpoint.md)|Returns a  **[SlideShowWindow](slideshowwindow-object-powerpoint.md)** object that represents the slide show window in which the specified presentation is running. Read-only.|
|[SnapToGrid](presentation-snaptogrid-property-powerpoint.md)|Determines whether to snap shapes to the gridlines in the specified presentation. Read/write.|
|[Sync](presentation-sync-property-powerpoint.md)|Returns a  **Sync** object that enables you to manage the synchronization of the local and server copies of a shared presentation stored in a Microsoft SharePoint Server shared workspace. Read-only.|
|[Tags](presentation-tags-property-powerpoint.md)|Returns a  **[Tags](tags-object-powerpoint.md)** object that represents the tags for the specified object. Read-only.|
|[TemplateName](presentation-templatename-property-powerpoint.md)|Returns the name of the design template associated with the specified presentation. Read-only.|
|[TitleMaster](presentation-titlemaster-property-powerpoint.md)|Returns a  **[Master](master-object-powerpoint.md)** object that represents the title master for the specified presentation.|
|[VBASigned](presentation-vbasigned-property-powerpoint.md)|Determines whether the Visual Basic for Applications (VBA) project for the specified document has been digitally signed. Read-only.|
|[VBProject](presentation-vbproject-property-powerpoint.md)|Returns a  **VBProject** object that represents the individual Visual Basic project for the presentation. Read-only.|
|[Windows](presentation-windows-property-powerpoint.md)|Returns a  **[DocumentWindows](documentwindows-object-powerpoint.md)** collection that represents all document windows associated with the specified presentation. Read-only.|
|[WritePassword](presentation-writepassword-property-powerpoint.md)|Sets or returns the password for saving changes to the specified document. Read/write.|

