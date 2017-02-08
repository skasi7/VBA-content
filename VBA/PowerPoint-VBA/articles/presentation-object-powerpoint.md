---
title: Presentation Object (PowerPoint)
keywords: vbapp10.chm524000
f1_keywords:
- vbapp10.chm524000
ms.prod: POWERPOINT
ms.assetid: ec75cf52-69f8-d35b-0a26-4a8da8a9683f
---


# Presentation Object (PowerPoint)

Represents a Microsoft PowerPoint presentation. 


## Remarks

The  **Presentation** object is a member of the **[Presentations](presentations-object-powerpoint.md)** collection. The **Presentations** collection contains all the **Presentation** objects that represent open presentations in PowerPoint.

The following examples describe how to:


- Return a presentation that you specify by name or index number
    
- Return the presentation in the active window
    
- Return the presentation in any document window or slide show window you specify
    

## Example

Use  **Presentations** (index), where index is the presentation's name or index number, to return a single **Presentation** object. The name of the presentation is the file name, with or without the file name extension, and without the path. The following example adds a slide to the beginning of Sample Presentation.


```
Presentations("Sample Presentation").Slides.Add 1, 1
```

Note that if multiple presentations with the same name are open, the first presentation in the collection with the specified name is returned.

Use the [ActivePresentation](http://msdn.microsoft.com/library/application-activepresentation-property-powerpoint%28Office.15%29.aspx)property to return the presentation in the active window. The following example saves the active presentation.




```
ActivePresentation.Save
```

Use the [Presentation](http://msdn.microsoft.com/library/documentwindow-presentation-property-powerpoint%28Office.15%29.aspx)property to return the presentation that's in the specified document window or slide show window. The following example displays the name of the slide show running in slide show window one.




```
MsgBox SlideShowWindows(1).Presentation.Name
```


## Methods



|**Name**|
|:-----|
|[AcceptAll](http://msdn.microsoft.com/library/presentation-acceptall-method-powerpoint%28Office.15%29.aspx)|
|[AddTitleMaster](http://msdn.microsoft.com/library/presentation-addtitlemaster-method-powerpoint%28Office.15%29.aspx)|
|[AddToFavorites](http://msdn.microsoft.com/library/presentation-addtofavorites-method-powerpoint%28Office.15%29.aspx)|
|[ApplyTemplate](http://msdn.microsoft.com/library/presentation-applytemplate-method-powerpoint%28Office.15%29.aspx)|
|[ApplyTemplate2](http://msdn.microsoft.com/library/presentation-applytemplate2-method-powerpoint%28Office.15%29.aspx)|
|[ApplyTheme](http://msdn.microsoft.com/library/presentation-applytheme-method-powerpoint%28Office.15%29.aspx)|
|[CanCheckIn](http://msdn.microsoft.com/library/presentation-cancheckin-method-powerpoint%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/presentation-checkin-method-powerpoint%28Office.15%29.aspx)|
|[CheckInWithVersion](http://msdn.microsoft.com/library/presentation-checkinwithversion-method-powerpoint%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/presentation-close-method-powerpoint%28Office.15%29.aspx)|
|[Convert2](http://msdn.microsoft.com/library/presentation-convert2-method-powerpoint%28Office.15%29.aspx)|
|[CreateVideo](http://msdn.microsoft.com/library/presentation-createvideo-method-powerpoint%28Office.15%29.aspx)|
|[EndReview](http://msdn.microsoft.com/library/presentation-endreview-method-powerpoint%28Office.15%29.aspx)|
|[EnsureAllMediaUpgraded](http://msdn.microsoft.com/library/presentation-ensureallmediaupgraded-method-powerpoint%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/presentation-export-method-powerpoint%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/presentation-exportasfixedformat-method-powerpoint%28Office.15%29.aspx)|
|[ExportAsFixedFormat2](http://msdn.microsoft.com/library/presentation-exportasfixedformat2-method-powerpoint%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/presentation-followhyperlink-method-powerpoint%28Office.15%29.aspx)|
|[GetWorkflowTasks](http://msdn.microsoft.com/library/presentation-getworkflowtasks-method-powerpoint%28Office.15%29.aspx)|
|[GetWorkflowTemplates](http://msdn.microsoft.com/library/presentation-getworkflowtemplates-method-powerpoint%28Office.15%29.aspx)|
|[LockServerFile](http://msdn.microsoft.com/library/presentation-lockserverfile-method-powerpoint%28Office.15%29.aspx)|
|[Merge](http://msdn.microsoft.com/library/presentation-merge-method-powerpoint%28Office.15%29.aspx)|
|[MergeWithBaseline](http://msdn.microsoft.com/library/presentation-mergewithbaseline-method-powerpoint%28Office.15%29.aspx)|
|[NewWindow](http://msdn.microsoft.com/library/presentation-newwindow-method-powerpoint%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/presentation-printout-method-powerpoint%28Office.15%29.aspx)|
|[PublishSlides](http://msdn.microsoft.com/library/presentation-publishslides-method-powerpoint%28Office.15%29.aspx)|
|[RejectAll](http://msdn.microsoft.com/library/presentation-rejectall-method-powerpoint%28Office.15%29.aspx)|
|[RemoveDocumentInformation](http://msdn.microsoft.com/library/presentation-removedocumentinformation-method-powerpoint%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/presentation-save-method-powerpoint%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/presentation-saveas-method-powerpoint%28Office.15%29.aspx)|
|[SaveCopyAs](http://msdn.microsoft.com/library/presentation-savecopyas-method-powerpoint%28Office.15%29.aspx)|
|[SendFaxOverInternet](http://msdn.microsoft.com/library/presentation-sendfaxoverinternet-method-powerpoint%28Office.15%29.aspx)|
|[SetPasswordEncryptionOptions](http://msdn.microsoft.com/library/presentation-setpasswordencryptionoptions-method-powerpoint%28Office.15%29.aspx)|
|[UpdateLinks](http://msdn.microsoft.com/library/presentation-updatelinks-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/presentation-application-property-powerpoint%28Office.15%29.aspx)|
|[Broadcast](http://msdn.microsoft.com/library/presentation-broadcast-property-powerpoint%28Office.15%29.aspx)|
|[BuiltInDocumentProperties](http://msdn.microsoft.com/library/presentation-builtindocumentproperties-property-powerpoint%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/presentation-chartdatapointtrack-property-powerpoint%28Office.15%29.aspx)|
|[Coauthoring](http://msdn.microsoft.com/library/presentation-coauthoring-property-powerpoint%28Office.15%29.aspx)|
|[ColorSchemes](http://msdn.microsoft.com/library/presentation-colorschemes-property-powerpoint%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/presentation-commandbars-property-powerpoint%28Office.15%29.aspx)|
|[Container](http://msdn.microsoft.com/library/presentation-container-property-powerpoint%28Office.15%29.aspx)|
|[ContentTypeProperties](http://msdn.microsoft.com/library/presentation-contenttypeproperties-property-powerpoint%28Office.15%29.aspx)|
|[CreateVideoStatus](http://msdn.microsoft.com/library/presentation-createvideostatus-property-powerpoint%28Office.15%29.aspx)|
|[CustomDocumentProperties](http://msdn.microsoft.com/library/presentation-customdocumentproperties-property-powerpoint%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/presentation-customerdata-property-powerpoint%28Office.15%29.aspx)|
|[CustomXMLParts](http://msdn.microsoft.com/library/presentation-customxmlparts-property-powerpoint%28Office.15%29.aspx)|
|[DefaultLanguageID](http://msdn.microsoft.com/library/presentation-defaultlanguageid-property-powerpoint%28Office.15%29.aspx)|
|[DefaultShape](http://msdn.microsoft.com/library/presentation-defaultshape-property-powerpoint%28Office.15%29.aspx)|
|[Designs](http://msdn.microsoft.com/library/presentation-designs-property-powerpoint%28Office.15%29.aspx)|
|[DisplayComments](http://msdn.microsoft.com/library/presentation-displaycomments-property-powerpoint%28Office.15%29.aspx)|
|[DocumentInspectors](http://msdn.microsoft.com/library/presentation-documentinspectors-property-powerpoint%28Office.15%29.aspx)|
|[DocumentLibraryVersions](http://msdn.microsoft.com/library/presentation-documentlibraryversions-property-powerpoint%28Office.15%29.aspx)|
|[EncryptionProvider](http://msdn.microsoft.com/library/presentation-encryptionprovider-property-powerpoint%28Office.15%29.aspx)|
|[EnvelopeVisible](http://msdn.microsoft.com/library/presentation-envelopevisible-property-powerpoint%28Office.15%29.aspx)|
|[ExtraColors](http://msdn.microsoft.com/library/presentation-extracolors-property-powerpoint%28Office.15%29.aspx)|
|[FarEastLineBreakLanguage](http://msdn.microsoft.com/library/presentation-fareastlinebreaklanguage-property-powerpoint%28Office.15%29.aspx)|
|[FarEastLineBreakLevel](http://msdn.microsoft.com/library/presentation-fareastlinebreaklevel-property-powerpoint%28Office.15%29.aspx)|
|[Final](http://msdn.microsoft.com/library/presentation-final-property-powerpoint%28Office.15%29.aspx)|
|[Fonts](http://msdn.microsoft.com/library/presentation-fonts-property-powerpoint%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/presentation-fullname-property-powerpoint%28Office.15%29.aspx)|
|[GridDistance](http://msdn.microsoft.com/library/presentation-griddistance-property-powerpoint%28Office.15%29.aspx)|
|[Guides](http://msdn.microsoft.com/library/presentation-guides-property-powerpoint%28Office.15%29.aspx)|
|[HandoutMaster](http://msdn.microsoft.com/library/presentation-handoutmaster-property-powerpoint%28Office.15%29.aspx)|
|[HasHandoutMaster](http://msdn.microsoft.com/library/presentation-hashandoutmaster-property-powerpoint%28Office.15%29.aspx)|
|[HasNotesMaster](http://msdn.microsoft.com/library/presentation-hasnotesmaster-property-powerpoint%28Office.15%29.aspx)|
|[HasTitleMaster](http://msdn.microsoft.com/library/presentation-hastitlemaster-property-powerpoint%28Office.15%29.aspx)|
|[HasVBProject](http://msdn.microsoft.com/library/presentation-hasvbproject-property-powerpoint%28Office.15%29.aspx)|
|[InMergeMode](http://msdn.microsoft.com/library/presentation-inmergemode-property-powerpoint%28Office.15%29.aspx)|
|[LayoutDirection](http://msdn.microsoft.com/library/presentation-layoutdirection-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/presentation-name-property-powerpoint%28Office.15%29.aspx)|
|[NoLineBreakAfter](http://msdn.microsoft.com/library/presentation-nolinebreakafter-property-powerpoint%28Office.15%29.aspx)|
|[NoLineBreakBefore](http://msdn.microsoft.com/library/presentation-nolinebreakbefore-property-powerpoint%28Office.15%29.aspx)|
|[NotesMaster](http://msdn.microsoft.com/library/presentation-notesmaster-property-powerpoint%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/presentation-pagesetup-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/presentation-parent-property-powerpoint%28Office.15%29.aspx)|
|[Password](http://msdn.microsoft.com/library/presentation-password-property-powerpoint%28Office.15%29.aspx)|
|[PasswordEncryptionAlgorithm](http://msdn.microsoft.com/library/presentation-passwordencryptionalgorithm-property-powerpoint%28Office.15%29.aspx)|
|[PasswordEncryptionFileProperties](http://msdn.microsoft.com/library/presentation-passwordencryptionfileproperties-property-powerpoint%28Office.15%29.aspx)|
|[PasswordEncryptionKeyLength](http://msdn.microsoft.com/library/presentation-passwordencryptionkeylength-property-powerpoint%28Office.15%29.aspx)|
|[PasswordEncryptionProvider](http://msdn.microsoft.com/library/presentation-passwordencryptionprovider-property-powerpoint%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/presentation-path-property-powerpoint%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/presentation-permission-property-powerpoint%28Office.15%29.aspx)|
|[PrintOptions](http://msdn.microsoft.com/library/presentation-printoptions-property-powerpoint%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/presentation-readonly-property-powerpoint%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/presentation-removepersonalinformation-property-powerpoint%28Office.15%29.aspx)|
|[Research](http://msdn.microsoft.com/library/presentation-research-property-powerpoint%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/presentation-saved-property-powerpoint%28Office.15%29.aspx)|
|[SectionProperties](http://msdn.microsoft.com/library/presentation-sectionproperties-property-powerpoint%28Office.15%29.aspx)|
|[ServerPolicy](http://msdn.microsoft.com/library/presentation-serverpolicy-property-powerpoint%28Office.15%29.aspx)|
|[SharedWorkspace](http://msdn.microsoft.com/library/presentation-sharedworkspace-property-powerpoint%28Office.15%29.aspx)|
|[Signatures](http://msdn.microsoft.com/library/presentation-signatures-property-powerpoint%28Office.15%29.aspx)|
|[SlideMaster](http://msdn.microsoft.com/library/presentation-slidemaster-property-powerpoint%28Office.15%29.aspx)|
|[Slides](http://msdn.microsoft.com/library/presentation-slides-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowSettings](http://msdn.microsoft.com/library/presentation-slideshowsettings-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowWindow](http://msdn.microsoft.com/library/presentation-slideshowwindow-property-powerpoint%28Office.15%29.aspx)|
|[SnapToGrid](http://msdn.microsoft.com/library/presentation-snaptogrid-property-powerpoint%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/presentation-sync-property-powerpoint%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/presentation-tags-property-powerpoint%28Office.15%29.aspx)|
|[TemplateName](http://msdn.microsoft.com/library/presentation-templatename-property-powerpoint%28Office.15%29.aspx)|
|[TitleMaster](http://msdn.microsoft.com/library/presentation-titlemaster-property-powerpoint%28Office.15%29.aspx)|
|[VBASigned](http://msdn.microsoft.com/library/presentation-vbasigned-property-powerpoint%28Office.15%29.aspx)|
|[VBProject](http://msdn.microsoft.com/library/presentation-vbproject-property-powerpoint%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/presentation-windows-property-powerpoint%28Office.15%29.aspx)|
|[WritePassword](http://msdn.microsoft.com/library/presentation-writepassword-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
