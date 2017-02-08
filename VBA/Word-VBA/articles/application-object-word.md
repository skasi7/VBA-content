---
title: Application Object (Word)
keywords: vbawd10.chm2416
f1_keywords:
- vbawd10.chm2416
ms.prod: WORD
api_name:
- Word.Application
ms.assetid: d1cf6f8f-4e88-bf01-93b4-90a83f79cb44
---


# Application Object (Word)

Represents the Microsoft Word application. The  **Application** object includes properties and methods that return top-level objects. For example, the **[ActiveDocument](http://msdn.microsoft.com/library/application-activedocument-property-word%28Office.15%29.aspx)** property returns a **[Document](document-object-word.md)** object.


## Remarks

Use the  **Application** property to return the **Application** object. The following example displays the user name for Word.


```
MsgBox Application.UserName
```

Many of the properties and methods that return the most common user-interface objects—such as the active document ( **ActiveDocument** property)—can be used without the **Application** object qualifier. For example, instead of writing `Application.ActiveDocument.PrintOut`, you can write  `ActiveDocument.PrintOut`. Properties and methods that can be used without the  **Application** object qualifier are considered "global." To view the global properties and methods in the **Object Browser**, click  **<globals>** at the top of the list in the **Classes** box. (Also see the **[Global](http://msdn.microsoft.com/library/global-object-word%28Office.15%29.aspx)** object.)

Remarks

To use Automation (formerly OLE Automation) to control Word from another application, use the Microsoft Visual Basic  **CreateObject** or **GetObject** function to return a Word **Application** object. The following Microsoft Excel example starts Word (if it is not already running) and opens an existing document.




```
Set wrd = GetObject(, "Word.Application") 
wrd.Visible = True 
wrd.Documents.Open "C:\My Documents\Temp.doc" 
Set wrd = Nothing
```


## Events



|**Name**|
|:-----|
|[DocumentBeforeClose](http://msdn.microsoft.com/library/application-documentbeforeclose-event-word%28Office.15%29.aspx)|
|[DocumentBeforePrint](http://msdn.microsoft.com/library/application-documentbeforeprint-event-word%28Office.15%29.aspx)|
|[DocumentBeforeSave](http://msdn.microsoft.com/library/application-documentbeforesave-event-word%28Office.15%29.aspx)|
|[DocumentChange](http://msdn.microsoft.com/library/application-documentchange-event-word%28Office.15%29.aspx)|
|[DocumentOpen](http://msdn.microsoft.com/library/application-documentopen-event-word%28Office.15%29.aspx)|
|[DocumentSync](http://msdn.microsoft.com/library/application-documentsync-event-word%28Office.15%29.aspx)|
|[EPostageInsert](http://msdn.microsoft.com/library/application-epostageinsert-event-word%28Office.15%29.aspx)|
|[EPostageInsertEx](http://msdn.microsoft.com/library/application-epostageinsertex-event-word%28Office.15%29.aspx)|
|[EPostagePropertyDialog](http://msdn.microsoft.com/library/application-epostagepropertydialog-event-word%28Office.15%29.aspx)|
|[MailMergeAfterMerge](http://msdn.microsoft.com/library/application-mailmergeaftermerge-event-word%28Office.15%29.aspx)|
|[MailMergeAfterRecordMerge](http://msdn.microsoft.com/library/application-mailmergeafterrecordmerge-event-word%28Office.15%29.aspx)|
|[MailMergeBeforeMerge](http://msdn.microsoft.com/library/application-mailmergebeforemerge-event-word%28Office.15%29.aspx)|
|[MailMergeBeforeRecordMerge](http://msdn.microsoft.com/library/application-mailmergebeforerecordmerge-event-word%28Office.15%29.aspx)|
|[MailMergeDataSourceLoad](http://msdn.microsoft.com/library/application-mailmergedatasourceload-event-word%28Office.15%29.aspx)|
|[MailMergeDataSourceValidate](http://msdn.microsoft.com/library/application-mailmergedatasourcevalidate-event-word%28Office.15%29.aspx)|
|[MailMergeDataSourceValidate2](http://msdn.microsoft.com/library/application-mailmergedatasourcevalidate2-event-word%28Office.15%29.aspx)|
|[MailMergeWizardSendToCustom](http://msdn.microsoft.com/library/application-mailmergewizardsendtocustom-event-word%28Office.15%29.aspx)|
|[MailMergeWizardStateChange](http://msdn.microsoft.com/library/application-mailmergewizardstatechange-event-word%28Office.15%29.aspx)|
|[NewDocument](http://msdn.microsoft.com/library/application-newdocument-event-word%28Office.15%29.aspx)|
|[ProtectedViewWindowActivate](http://msdn.microsoft.com/library/application-protectedviewwindowactivate-event-word%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeClose](http://msdn.microsoft.com/library/application-protectedviewwindowbeforeclose-event-word%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeEdit](http://msdn.microsoft.com/library/application-protectedviewwindowbeforeedit-event-word%28Office.15%29.aspx)|
|[ProtectedViewWindowDeactivate](http://msdn.microsoft.com/library/application-protectedviewwindowdeactivate-event-word%28Office.15%29.aspx)|
|[ProtectedViewWindowOpen](http://msdn.microsoft.com/library/application-protectedviewwindowopen-event-word%28Office.15%29.aspx)|
|[ProtectedViewWindowSize](http://msdn.microsoft.com/library/application-protectedviewwindowsize-event-word%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application-quit-event-word%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/application-windowactivate-event-word%28Office.15%29.aspx)|
|[WindowBeforeDoubleClick](http://msdn.microsoft.com/library/application-windowbeforedoubleclick-event-word%28Office.15%29.aspx)|
|[WindowBeforeRightClick](http://msdn.microsoft.com/library/application-windowbeforerightclick-event-word%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/application-windowdeactivate-event-word%28Office.15%29.aspx)|
|[WindowSelectionChange](http://msdn.microsoft.com/library/application-windowselectionchange-event-word%28Office.15%29.aspx)|
|[WindowSize](http://msdn.microsoft.com/library/application-windowsize-event-word%28Office.15%29.aspx)|
|[XMLSelectionChange](http://msdn.microsoft.com/library/application-xmlselectionchange-event-word%28Office.15%29.aspx)|
|[XMLValidationError](http://msdn.microsoft.com/library/application-xmlvalidationerror-event-word%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/application-activate-method-word%28Office.15%29.aspx)|
|[AddAddress](http://msdn.microsoft.com/library/application-addaddress-method-word%28Office.15%29.aspx)|
|[AutomaticChange](http://msdn.microsoft.com/library/application-automaticchange-method-word%28Office.15%29.aspx)|
|[BuildKeyCode](http://msdn.microsoft.com/library/application-buildkeycode-method-word%28Office.15%29.aspx)|
|[CentimetersToPoints](http://msdn.microsoft.com/library/application-centimeterstopoints-method-word%28Office.15%29.aspx)|
|[ChangeFileOpenDirectory](http://msdn.microsoft.com/library/application-changefileopendirectory-method-word%28Office.15%29.aspx)|
|[CheckGrammar](http://msdn.microsoft.com/library/application-checkgrammar-method-word%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/application-checkspelling-method-word%28Office.15%29.aspx)|
|[CleanString](http://msdn.microsoft.com/library/application-cleanstring-method-word%28Office.15%29.aspx)|
|[CompareDocuments](http://msdn.microsoft.com/library/application-comparedocuments-method-word%28Office.15%29.aspx)|
|[DDEExecute](http://msdn.microsoft.com/library/application-ddeexecute-method-word%28Office.15%29.aspx)|
|[DDEInitiate](http://msdn.microsoft.com/library/application-ddeinitiate-method-word%28Office.15%29.aspx)|
|[DDEPoke](http://msdn.microsoft.com/library/application-ddepoke-method-word%28Office.15%29.aspx)|
|[DDERequest](http://msdn.microsoft.com/library/application-dderequest-method-word%28Office.15%29.aspx)|
|[DDETerminate](http://msdn.microsoft.com/library/application-ddeterminate-method-word%28Office.15%29.aspx)|
|[DDETerminateAll](http://msdn.microsoft.com/library/application-ddeterminateall-method-word%28Office.15%29.aspx)|
|[DefaultWebOptions](http://msdn.microsoft.com/library/application-defaultweboptions-method-word%28Office.15%29.aspx)|
|[GetAddress](http://msdn.microsoft.com/library/application-getaddress-method-word%28Office.15%29.aspx)|
|[GetDefaultTheme](http://msdn.microsoft.com/library/application-getdefaulttheme-method-word%28Office.15%29.aspx)|
|[GetSpellingSuggestions](http://msdn.microsoft.com/library/application-getspellingsuggestions-method-word%28Office.15%29.aspx)|
|[GoBack](http://msdn.microsoft.com/library/application-goback-method-word%28Office.15%29.aspx)|
|[GoForward](http://msdn.microsoft.com/library/application-goforward-method-word%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/application-help-method-word%28Office.15%29.aspx)|
|[HelpTool](http://msdn.microsoft.com/library/application-helptool-method-word%28Office.15%29.aspx)|
|[InchesToPoints](http://msdn.microsoft.com/library/application-inchestopoints-method-word%28Office.15%29.aspx)|
|[Keyboard](http://msdn.microsoft.com/library/application-keyboard-method-word%28Office.15%29.aspx)|
|[KeyboardBidi](http://msdn.microsoft.com/library/application-keyboardbidi-method-word%28Office.15%29.aspx)|
|[KeyboardLatin](http://msdn.microsoft.com/library/application-keyboardlatin-method-word%28Office.15%29.aspx)|
|[KeyString](http://msdn.microsoft.com/library/application-keystring-method-word%28Office.15%29.aspx)|
|[LinesToPoints](http://msdn.microsoft.com/library/application-linestopoints-method-word%28Office.15%29.aspx)|
|[ListCommands](http://msdn.microsoft.com/library/application-listcommands-method-word%28Office.15%29.aspx)|
|[LoadMasterList](http://msdn.microsoft.com/library/application-loadmasterlist-method-word%28Office.15%29.aspx)|
|[LookupNameProperties](http://msdn.microsoft.com/library/application-lookupnameproperties-method-word%28Office.15%29.aspx)|
|[MergeDocuments](http://msdn.microsoft.com/library/application-mergedocuments-method-word%28Office.15%29.aspx)|
|[MillimetersToPoints](http://msdn.microsoft.com/library/application-millimeterstopoints-method-word%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/application-move-method-word%28Office.15%29.aspx)|
|[NewWindow](http://msdn.microsoft.com/library/application-newwindow-method-word%28Office.15%29.aspx)|
|[NextLetter](http://msdn.microsoft.com/library/application-nextletter-method-word%28Office.15%29.aspx)|
|[OnTime](http://msdn.microsoft.com/library/application-ontime-method-word%28Office.15%29.aspx)|
|[OrganizerCopy](http://msdn.microsoft.com/library/application-organizercopy-method-word%28Office.15%29.aspx)|
|[OrganizerDelete](http://msdn.microsoft.com/library/application-organizerdelete-method-word%28Office.15%29.aspx)|
|[OrganizerRename](http://msdn.microsoft.com/library/application-organizerrename-method-word%28Office.15%29.aspx)|
|[PicasToPoints](http://msdn.microsoft.com/library/application-picastopoints-method-word%28Office.15%29.aspx)|
|[PixelsToPoints](http://msdn.microsoft.com/library/application-pixelstopoints-method-word%28Office.15%29.aspx)|
|[PointsToCentimeters](http://msdn.microsoft.com/library/application-pointstocentimeters-method-word%28Office.15%29.aspx)|
|[PointsToInches](http://msdn.microsoft.com/library/application-pointstoinches-method-word%28Office.15%29.aspx)|
|[PointsToLines](http://msdn.microsoft.com/library/application-pointstolines-method-word%28Office.15%29.aspx)|
|[PointsToMillimeters](http://msdn.microsoft.com/library/application-pointstomillimeters-method-word%28Office.15%29.aspx)|
|[PointsToPicas](http://msdn.microsoft.com/library/application-pointstopicas-method-word%28Office.15%29.aspx)|
|[PointsToPixels](http://msdn.microsoft.com/library/application-pointstopixels-method-word%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/application-printout-method-word%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/application-productcode-method-word%28Office.15%29.aspx)|
|[PutFocusInMailHeader](http://msdn.microsoft.com/library/application-putfocusinmailheader-method-word%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application-quit-method-word%28Office.15%29.aspx)|
|[Repeat](http://msdn.microsoft.com/library/application-repeat-method-word%28Office.15%29.aspx)|
|[ResetIgnoreAll](http://msdn.microsoft.com/library/application-resetignoreall-method-word%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/application-resize-method-word%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/application-run-method-word%28Office.15%29.aspx)|
|[ScreenRefresh](http://msdn.microsoft.com/library/application-screenrefresh-method-word%28Office.15%29.aspx)|
|[SetDefaultTheme](http://msdn.microsoft.com/library/application-setdefaulttheme-method-word%28Office.15%29.aspx)|
|[ShowClipboard](http://msdn.microsoft.com/library/application-showclipboard-method-word%28Office.15%29.aspx)|
|[ShowMe](http://msdn.microsoft.com/library/application-showme-method-word%28Office.15%29.aspx)|
|[SubstituteFont](http://msdn.microsoft.com/library/application-substitutefont-method-word%28Office.15%29.aspx)|
|[ToggleKeyboard](http://msdn.microsoft.com/library/application-togglekeyboard-method-word%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveDocument](http://msdn.microsoft.com/library/application-activedocument-property-word%28Office.15%29.aspx)|
|[ActiveEncryptionSession](http://msdn.microsoft.com/library/application-activeencryptionsession-property-word%28Office.15%29.aspx)|
|[ActivePrinter](http://msdn.microsoft.com/library/application-activeprinter-property-word%28Office.15%29.aspx)|
|[ActiveProtectedViewWindow](http://msdn.microsoft.com/library/application-activeprotectedviewwindow-property-word%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/application-activewindow-property-word%28Office.15%29.aspx)|
|[AddIns](http://msdn.microsoft.com/library/application-addins-property-word%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/application-application-property-word%28Office.15%29.aspx)|
|[ArbitraryXMLSupportAvailable](http://msdn.microsoft.com/library/application-arbitraryxmlsupportavailable-property-word%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/application-assistance-property-word%28Office.15%29.aspx)|
|[AutoCaptions](http://msdn.microsoft.com/library/application-autocaptions-property-word%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/application-autocorrect-property-word%28Office.15%29.aspx)|
|[AutoCorrectEmail](http://msdn.microsoft.com/library/application-autocorrectemail-property-word%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/application-automationsecurity-property-word%28Office.15%29.aspx)|
|[BackgroundPrintingStatus](http://msdn.microsoft.com/library/application-backgroundprintingstatus-property-word%28Office.15%29.aspx)|
|[BackgroundSavingStatus](http://msdn.microsoft.com/library/application-backgroundsavingstatus-property-word%28Office.15%29.aspx)|
|[Bibliography](http://msdn.microsoft.com/library/application-bibliography-property-word%28Office.15%29.aspx)|
|[BrowseExtraFileTypes](http://msdn.microsoft.com/library/application-browseextrafiletypes-property-word%28Office.15%29.aspx)|
|[Browser](http://msdn.microsoft.com/library/application-browser-property-word%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/application-build-property-word%28Office.15%29.aspx)|
|[CapsLock](http://msdn.microsoft.com/library/application-capslock-property-word%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/application-caption-property-word%28Office.15%29.aspx)|
|[CaptionLabels](http://msdn.microsoft.com/library/application-captionlabels-property-word%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/application-chartdatapointtrack-property-word%28Office.15%29.aspx)|
|[CheckLanguage](http://msdn.microsoft.com/library/application-checklanguage-property-word%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/application-comaddins-property-word%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/application-commandbars-property-word%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/application-creator-property-word%28Office.15%29.aspx)|
|[CustomDictionaries](http://msdn.microsoft.com/library/application-customdictionaries-property-word%28Office.15%29.aspx)|
|[CustomizationContext](http://msdn.microsoft.com/library/application-customizationcontext-property-word%28Office.15%29.aspx)|
|[DefaultLegalBlackline](http://msdn.microsoft.com/library/application-defaultlegalblackline-property-word%28Office.15%29.aspx)|
|[DefaultSaveFormat](http://msdn.microsoft.com/library/application-defaultsaveformat-property-word%28Office.15%29.aspx)|
|[DefaultTableSeparator](http://msdn.microsoft.com/library/application-defaulttableseparator-property-word%28Office.15%29.aspx)|
|[Dialogs](http://msdn.microsoft.com/library/application-dialogs-property-word%28Office.15%29.aspx)|
|[DisplayAlerts](http://msdn.microsoft.com/library/application-displayalerts-property-word%28Office.15%29.aspx)|
|[DisplayAutoCompleteTips](http://msdn.microsoft.com/library/application-displayautocompletetips-property-word%28Office.15%29.aspx)|
|[DisplayDocumentInformationPanel](http://msdn.microsoft.com/library/application-displaydocumentinformationpanel-property-word%28Office.15%29.aspx)|
|[DisplayRecentFiles](http://msdn.microsoft.com/library/application-displayrecentfiles-property-word%28Office.15%29.aspx)|
|[DisplayScreenTips](http://msdn.microsoft.com/library/application-displayscreentips-property-word%28Office.15%29.aspx)|
|[DisplayScrollBars](http://msdn.microsoft.com/library/application-displayscrollbars-property-word%28Office.15%29.aspx)|
|[Documents](http://msdn.microsoft.com/library/application-documents-property-word%28Office.15%29.aspx)|
|[DontResetInsertionPointProperties](http://msdn.microsoft.com/library/application-dontresetinsertionpointproperties-property-word%28Office.15%29.aspx)|
|[EmailOptions](http://msdn.microsoft.com/library/application-emailoptions-property-word%28Office.15%29.aspx)|
|[EmailTemplate](http://msdn.microsoft.com/library/application-emailtemplate-property-word%28Office.15%29.aspx)|
|[EnableCancelKey](http://msdn.microsoft.com/library/application-enablecancelkey-property-word%28Office.15%29.aspx)|
|[FeatureInstall](http://msdn.microsoft.com/library/application-featureinstall-property-word%28Office.15%29.aspx)|
|[FileConverters](http://msdn.microsoft.com/library/application-fileconverters-property-word%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/application-filedialog-property-word%28Office.15%29.aspx)|
|[FileValidation](http://msdn.microsoft.com/library/application-filevalidation-property-word%28Office.15%29.aspx)|
|[FindKey](http://msdn.microsoft.com/library/application-findkey-property-word%28Office.15%29.aspx)|
|[FocusInMailHeader](http://msdn.microsoft.com/library/application-focusinmailheader-property-word%28Office.15%29.aspx)|
|[FontNames](http://msdn.microsoft.com/library/application-fontnames-property-word%28Office.15%29.aspx)|
|[HangulHanjaDictionaries](http://msdn.microsoft.com/library/application-hangulhanjadictionaries-property-word%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/application-height-property-word%28Office.15%29.aspx)|
|[International](http://msdn.microsoft.com/library/application-international-property-word%28Office.15%29.aspx)|
|[IsObjectValid](http://msdn.microsoft.com/library/application-isobjectvalid-property-word%28Office.15%29.aspx)|
|[IsSandboxed](http://msdn.microsoft.com/library/application-issandboxed-property-word%28Office.15%29.aspx)|
|[KeyBindings](http://msdn.microsoft.com/library/application-keybindings-property-word%28Office.15%29.aspx)|
|[KeysBoundTo](http://msdn.microsoft.com/library/application-keysboundto-property-word%28Office.15%29.aspx)|
|[LandscapeFontNames](http://msdn.microsoft.com/library/application-landscapefontnames-property-word%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/application-language-property-word%28Office.15%29.aspx)|
|[Languages](http://msdn.microsoft.com/library/application-languages-property-word%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/application-languagesettings-property-word%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/application-left-property-word%28Office.15%29.aspx)|
|[ListGalleries](http://msdn.microsoft.com/library/application-listgalleries-property-word%28Office.15%29.aspx)|
|[MacroContainer](http://msdn.microsoft.com/library/application-macrocontainer-property-word%28Office.15%29.aspx)|
|[MailingLabel](http://msdn.microsoft.com/library/application-mailinglabel-property-word%28Office.15%29.aspx)|
|[MailMessage](http://msdn.microsoft.com/library/application-mailmessage-property-word%28Office.15%29.aspx)|
|[MailSystem](http://msdn.microsoft.com/library/application-mailsystem-property-word%28Office.15%29.aspx)|
|[MAPIAvailable](http://msdn.microsoft.com/library/application-mapiavailable-property-word%28Office.15%29.aspx)|
|[MathCoprocessorAvailable](http://msdn.microsoft.com/library/application-mathcoprocessoravailable-property-word%28Office.15%29.aspx)|
|[MouseAvailable](http://msdn.microsoft.com/library/application-mouseavailable-property-word%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/application-name-property-word%28Office.15%29.aspx)|
|[NewDocument](http://msdn.microsoft.com/library/application-newdocument-property-word%28Office.15%29.aspx)|
|[NormalTemplate](http://msdn.microsoft.com/library/application-normaltemplate-property-word%28Office.15%29.aspx)|
|[NumLock](http://msdn.microsoft.com/library/application-numlock-property-word%28Office.15%29.aspx)|
|[OMathAutoCorrect](http://msdn.microsoft.com/library/application-omathautocorrect-property-word%28Office.15%29.aspx)|
|[OpenAttachmentsInFullScreen](http://msdn.microsoft.com/library/application-openattachmentsinfullscreen-property-word%28Office.15%29.aspx)|
|[Options](http://msdn.microsoft.com/library/application-options-property-word%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/application-parent-property-word%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/application-path-property-word%28Office.15%29.aspx)|
|[PathSeparator](http://msdn.microsoft.com/library/application-pathseparator-property-word%28Office.15%29.aspx)|
|[PickerDialog](http://msdn.microsoft.com/library/application-pickerdialog-property-word%28Office.15%29.aspx)|
|[PortraitFontNames](http://msdn.microsoft.com/library/application-portraitfontnames-property-word%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/application-printpreview-property-word%28Office.15%29.aspx)|
|[ProtectedViewWindows](http://msdn.microsoft.com/library/application-protectedviewwindows-property-word%28Office.15%29.aspx)|
|[RecentFiles](http://msdn.microsoft.com/library/application-recentfiles-property-word%28Office.15%29.aspx)|
|[RestrictLinkedStyles](http://msdn.microsoft.com/library/application-restrictlinkedstyles-property-word%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/application-screenupdating-property-word%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/application-selection-property-word%28Office.15%29.aspx)|
|[ShowAnimation](http://msdn.microsoft.com/library/application-showanimation-property-word%28Office.15%29.aspx)|
|[ShowStartupDialog](http://msdn.microsoft.com/library/application-showstartupdialog-property-word%28Office.15%29.aspx)|
|[ShowStylePreviews](http://msdn.microsoft.com/library/application-showstylepreviews-property-word%28Office.15%29.aspx)|
|[ShowVisualBasicEditor](http://msdn.microsoft.com/library/application-showvisualbasiceditor-property-word%28Office.15%29.aspx)|
|[SmartArtColors](http://msdn.microsoft.com/library/application-smartartcolors-property-word%28Office.15%29.aspx)|
|[SmartArtLayouts](http://msdn.microsoft.com/library/application-smartartlayouts-property-word%28Office.15%29.aspx)|
|[SmartArtQuickStyles](http://msdn.microsoft.com/library/application-smartartquickstyles-property-word%28Office.15%29.aspx)|
|[SpecialMode](http://msdn.microsoft.com/library/application-specialmode-property-word%28Office.15%29.aspx)|
|[StartupPath](http://msdn.microsoft.com/library/application-startuppath-property-word%28Office.15%29.aspx)|
|[StatusBar](http://msdn.microsoft.com/library/application-statusbar-property-word%28Office.15%29.aspx)|
|[SynonymInfo](http://msdn.microsoft.com/library/application-synonyminfo-property-word%28Office.15%29.aspx)|
|[System](http://msdn.microsoft.com/library/application-system-property-word%28Office.15%29.aspx)|
|[TaskPanes](http://msdn.microsoft.com/library/application-taskpanes-property-word%28Office.15%29.aspx)|
|[Tasks](http://msdn.microsoft.com/library/application-tasks-property-word%28Office.15%29.aspx)|
|[Templates](http://msdn.microsoft.com/library/application-templates-property-word%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/application-top-property-word%28Office.15%29.aspx)|
|[UndoRecord](http://msdn.microsoft.com/library/application-undorecord-property-word%28Office.15%29.aspx)|
|[UsableHeight](http://msdn.microsoft.com/library/application-usableheight-property-word%28Office.15%29.aspx)|
|[UsableWidth](http://msdn.microsoft.com/library/application-usablewidth-property-word%28Office.15%29.aspx)|
|[UserAddress](http://msdn.microsoft.com/library/application-useraddress-property-word%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/application-usercontrol-property-word%28Office.15%29.aspx)|
|[UserInitials](http://msdn.microsoft.com/library/application-userinitials-property-word%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/application-username-property-word%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/application-vbe-property-word%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/application-version-property-word%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/application-visible-property-word%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/application-width-property-word%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/application-windows-property-word%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/application-windowstate-property-word%28Office.15%29.aspx)|
|[WordBasic](http://msdn.microsoft.com/library/application-wordbasic-property-word%28Office.15%29.aspx)|
|[XMLNamespaces](http://msdn.microsoft.com/library/application-xmlnamespaces-property-word%28Office.15%29.aspx)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)
<<<<<<< HEAD
=======

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

