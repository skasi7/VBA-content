---
title: Application Properties (Word)
ms.prod: WORD
ms.assetid: d3029d8f-bc98-40f8-b266-37c1428b9f14
---


# Application Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActiveDocument](application-activedocument-property-word.md)|Returns a  **[Document](document-object-word.md)** object that represents the active document (the document with the focus). If there are no documents open, an error occurs. Read-only.|
|[ActiveEncryptionSession](application-activeencryptionsession-property-word.md)|Returns a  **Long** that represents the encryption session associated with the active document. Read-only.|
|[ActivePrinter](application-activeprinter-property-word.md)|Returns or sets the name of the active printer. Read/write  **String** .|
|[ActiveProtectedViewWindow](application-activeprotectedviewwindow-property-word.md)|Returns a [ProtectedViewWindow](protectedviewwindow-object-word.md) object that represents the active protected view window. Read-only.|
|[ActiveWindow](application-activewindow-property-word.md)|Returns a  **[Window](window-object-word.md)** object that represents the active window (the window with the focus). If there are no windows open, an error occurs. Read-only.|
|[AddIns](application-addins-property-word.md)|Returns an  **[AddIns](addins-object-word.md)** collection that represents all available add-ins, regardless of whether they're currently loaded. Read-only.|
|[Application](application-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[ArbitraryXMLSupportAvailable](application-arbitraryxmlsupportavailable-property-word.md)|Returns a  **Boolean** that represents whether Microsoft Word accepts custom XML schemas. **True** indicates that Word accepts custom XML schemas.|
|[Assistance](application-assistance-property-word.md)|Returns an  **Assistance** object that represents the Microsoft Office Help Viewer. Read-only.|
|[AutoCaptions](application-autocaptions-property-word.md)|Returns an  **[AutoCaptions](autocaptions-object-word.md)** collection that represents the captions that are automatically added when items such as tables and pictures are inserted into a document. Read-only.|
|[AutoCorrect](application-autocorrect-property-word.md)|Returns an  **[AutoCorrect](autocorrect-object-word.md)** object that contains the current AutoCorrect options, entries, and exceptions. Read-only.|
|[AutoCorrectEmail](application-autocorrectemail-property-word.md)|Returns an  **[AutoCorrect](autocorrect-object-word.md)** object that represents automatic corrections made to e-mail messages.|
|[AutomationSecurity](application-automationsecurity-property-word.md)|Returns or sets an  **MsoAutomationSecurity** constant that represents the security setting Microsoft Word uses when programmatically opening files. .|
|[BackgroundPrintingStatus](application-backgroundprintingstatus-property-word.md)|Returns the number of print jobs in the background printing queue. Read-only  **Long** .|
|[BackgroundSavingStatus](application-backgroundsavingstatus-property-word.md)|Returns the number of files queued up to be saved in the background. Read-only  **Long** .|
|[Bibliography](application-bibliography-property-word.md)|Returns a  **[Bibliography](bibliography-object-word.md)** object that represents the bibliography references sources stored in Microsoft Word. Read-only.|
|[BrowseExtraFileTypes](application-browseextrafiletypes-property-word.md)|Set this property to "text/html" to allow hyperlinked HTML files to be opened in Microsoft Word (instead of the default Internet browser). Read/write  **String** .|
|[Browser](application-browser-property-word.md)|Returns a  **[Browser](browser-object-word.md)** object that represents the **Select Browse Object** tool on the vertical scroll bar. Read-only.|
|[Build](application-build-property-word.md)|Returns the version and build number of the Word application. Read-only  **String** .|
|[CapsLock](application-capslock-property-word.md)| **True** if the CAPS LOCK key is turned on. Read-only **Boolean** .|
|[Caption](application-caption-property-word.md)|Returns or sets the text displayed in the Title bar of the application window. Read/write  **String** .|
|[CaptionLabels](application-captionlabels-property-word.md)|Returns a  **[CaptionLabels](captionlabels-object-word.md)** collection that represents all the available caption labels. Read-only.|
|[ChartDataPointTrack](application-chartdatapointtrack-property-word.md)|Returns or sets a  **Boolean** that specifies whether charts use cell-reference data-point tracking. Read-write.|
|[CheckLanguage](application-checklanguage-property-word.md)| **True** if Microsoft Word automatically detects the language you are using as you type. Read/write **Boolean** .|
|[COMAddIns](application-comaddins-property-word.md)|Returns a reference to the  **COMAddIns** collection that represents all the Component Object Model (COM) add-ins currently loaded in Microsoft Word.|
|[CommandBars](application-commandbars-property-word.md)|Returns a  **CommandBars** collection that represents the menu bar and all the toolbars in Microsoft Word.|
|[Creator](application-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[CustomDictionaries](application-customdictionaries-property-word.md)|Returns a  **[Dictionaries](dictionaries-object-word.md)** object that represents the collection of active custom dictionaries. Read-only.|
|[CustomizationContext](application-customizationcontext-property-word.md)|Returns or sets a  **[Template](template-object-word.md)** or **[Document](document-object-word.md)** object that represents the template or document in which changes to menu bars, toolbars, and key bindings are stored. Read/write.|
|[DefaultLegalBlackline](application-defaultlegalblackline-property-word.md)| **True** for Microsoft Word to compare and merge documents using the **Legal blackline** option in the **Compare and Merge Documents** dialog box. Read/write **Boolean** .|
|[DefaultSaveFormat](application-defaultsaveformat-property-word.md)|Returns or sets the default format that will appear in the  **Save as type** box in the **Save As** dialog box. Read/write **String** .|
|[DefaultTableSeparator](application-defaulttableseparator-property-word.md)|Returns or sets the single character used to separate text into cells when text is converted to a table. Read/write  **String** .|
|[Dialogs](application-dialogs-property-word.md)|Returns a  **[Dialogs](dialogs-object-word.md)** collection that represents all the built-in dialog boxes in Word.Read-only.|
|[DisplayAlerts](application-displayalerts-property-word.md)|Returns or sets the way certain alerts and messages are handled while a macro is running. Read/write  **WdAlertLevel** .|
|[DisplayAutoCompleteTips](application-displayautocompletetips-property-word.md)| **True** if Word displays tips that suggest text for completing words, dates, or phrases as you type. Read/write **Boolean** .|
|[DisplayDocumentInformationPanel](application-displaydocumentinformationpanel-property-word.md)|Returns or sets a  **Boolean** that represents whether the document properties panel is displayed. Read/write.|
|[DisplayRecentFiles](application-displayrecentfiles-property-word.md)| **True** if the names of recently used files are displayed on the **File** menu. Read/write **Boolean** .|
|[DisplayScreenTips](application-displayscreentips-property-word.md)| **True** if comments, footnotes, endnotes, and hyperlinks are displayed as tips. Text marked as having comments is highlighted. Read/write **Boolean** .|
|[DisplayScrollBars](application-displayscrollbars-property-word.md)| **True** if Word displays a scroll bar in at least one document window. **False** if there are no scroll bars displayed in any window. Read/write **Boolean** .|
|[Documents](application-documents-property-word.md)|Returns a  **[Documents](documents-object-word.md)** collection that represents all the open documents. Read-only.|
|[DontResetInsertionPointProperties](application-dontresetinsertionpointproperties-property-word.md)|Returns or sets a  **Boolean** that represents whether Microsoft Word maintains the formatting properties of the text at that position of the Insertion Point after running other code. Read/write.|
|[EmailOptions](application-emailoptions-property-word.md)|Returns an  **[EmailOptions](emailoptions-object-word.md)** object that represents the global preferences for e-mail authoring. Read-only.|
|[EmailTemplate](application-emailtemplate-property-word.md)|Returns or sets a  **String** that represents the document template to use when sending e-mail messages. Read/write.|
|[EnableCancelKey](application-enablecancelkey-property-word.md)|Returns or sets the way that Word handles CTRL+BREAK user interruptions. Read/write  **WdEnableCancelKey** .|
|[FeatureInstall](application-featureinstall-property-word.md)|Returns or sets how Microsoft Word handles calls to methods and properties that require features not yet installed. Read/write  **MsoFeatureInstall** .|
|[FileConverters](application-fileconverters-property-word.md)|Returns a  **[FileConverters](fileconverters-object-word.md)** collection that represents all the file converters available to Microsoft Word. Read-only.|
|[FileDialog](application-filedialog-property-word.md)|Returns a  **FileDialog** object which represents a single instance of a file dialog box.|
|[FileValidation](application-filevalidation-property-word.md)|Returns or sets how Word will validate files before opening them. Read/write [MsoFileValidationMode](msofilevalidationmode-enumeration-office.md).|
|[FindKey](application-findkey-property-word.md)|Returns a  **[KeyBinding](keybinding-object-word.md)** object that represents the specified key combination. Read-only.|
|[FocusInMailHeader](application-focusinmailheader-property-word.md)| **True** if the insertion point is in an e-mail header field (the To: field, for example). Read-only **Boolean** .|
|[FontNames](application-fontnames-property-word.md)|Returns a  **[FontNames](fontnames-object-word.md)** object that includes the names of all the available fonts. Read-only.|
|[HangulHanjaDictionaries](application-hangulhanjadictionaries-property-word.md)|Returns a  **[HangulHanjaConversionDictionaries](hangulhanjaconversiondictionaries-object-word.md)** collection that represents all the active custom conversion dictionaries.|
|[Height](application-height-property-word.md)|Returns or sets the height of the active document window in pixels. Read/write  **Long** .|
|[International](application-international-property-word.md)|Returns information about the current country/region and international settings. Read-only  **Variant** .|
|[IsObjectValid](application-isobjectvalid-property-word.md)| **True** if the specified variable that references an object is valid. Read-only **Boolean** .|
|[IsSandboxed](application-issandboxed-property-word.md)| **True** if the application window is a protected view window. Read-only.|
|[KeyBindings](application-keybindings-property-word.md)|Returns a  **[KeyBindings](keybindings-object-word.md)** collection that represents customized key assignments, which include a key code, a key category, and a command.|
|[KeysBoundTo](application-keysboundto-property-word.md)|Returns a  **[KeysBoundTo](keysboundto-object-word.md)** object that represents all the key combinations assigned to the specified item.|
|[LandscapeFontNames](application-landscapefontnames-property-word.md)|Returns a  **[FontNames](fontnames-object-word.md)** object that includes the names of all the available landscape fonts.|
|[Language](application-language-property-word.md)|Returns an  **MsoLanguageID** constant that represents the language selected for the Microsoft Word user interface.|
|[Languages](application-languages-property-word.md)|Returns a  **[Languages](languages-object-word.md)** collection that represents the proofing languages listed in the **Language** dialog box.|
|[LanguageSettings](application-languagesettings-property-word.md)|Returns a  **LanguageSettings** object, which contains information about the language settings in Microsoft Word.|
|[Left](application-left-property-word.md)|Returns or sets a  **Long** that represents the horizontal position of the active document, measured in points. Read/write.|
|[ListGalleries](application-listgalleries-property-word.md)|Returns a  **[ListGalleries](listgalleries-object-word.md)** collection that represents the three list template galleries. .|
|[MacroContainer](application-macrocontainer-property-word.md)|Returns a  **[Template](template-object-word.md)** or **[Document](document-object-word.md)** object that represents the template or document in which the module that contains the running procedure is stored.|
|[MailingLabel](application-mailinglabel-property-word.md)|Returns a  **[MailingLabel](mailinglabel-object-word.md)** object that represents a mailing label.|
|[MailMessage](application-mailmessage-property-word.md)|Returns a  **[MailMessage](mailmessage-object-word.md)** object that represents the active e-mail message.|
|[MailSystem](application-mailsystem-property-word.md)|Returns the mail system (or systems) installed on the host computer. Read-only  **WdMailSystem** .|
|[MAPIAvailable](application-mapiavailable-property-word.md)| **True** if MAPI is installed. Read-only **Boolean** .|
|[MathCoprocessorAvailable](application-mathcoprocessoravailable-property-word.md)| **True** if a math coprocessor is installed and available to Microsoft Word. Read-only **Boolean** .|
|[MouseAvailable](application-mouseavailable-property-word.md)| **True** if there is a mouse available for the system. Read-only **Boolean** .|
|[Name](application-name-property-word.md)|Returns the name of the specified object. Read-only  **String** .|
|[NewDocument](application-newdocument-property-word.md)|Returns a  **NewFile** object that represents a document listed on the **New** tab.|
|[NormalTemplate](application-normaltemplate-property-word.md)|Returns a  **[Template](template-object-word.md)** object that represents the Normal template.|
|[NumLock](application-numlock-property-word.md)|Returns the state of the NUM LOCK key.  **True** if the keys on the numeric keypad insert numbers, **False** if the keys move the insertion point. Read-only **Boolean** .|
|[OMathAutoCorrect](application-omathautocorrect-property-word.md)|Returns an  **OMathAutoCorrect** object that represents the auto correct entries for equations. Read-only.|
|[OpenAttachmentsInFullScreen](application-openattachmentsinfullscreen-property-word.md)|Returns or sets a  **Boolean** that represents whether Microsoft Word opens e-mail attachments in Reading mode. Read/write.|
|[Options](application-options-property-word.md)|Returns an  **[Options](options-object-word.md)** object that represents application settings in Microsoft Word.|
|[Parent](application-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Application** object.|
|[Path](application-path-property-word.md)|Returns the disk or Web path to the specified object. Read-only  **String** .|
|[PathSeparator](application-pathseparator-property-word.md)|Returns the character used to separate folder names. This property returns a backslash (\). Read-only  **String** .|
|[PickerDialog](application-pickerdialog-property-word.md)|Returns a [PickerDialog](pickerdialog-object-office.md) object that provides the functionality to select people or data in a dialog box. Read-only.|
|[PortraitFontNames](application-portraitfontnames-property-word.md)|Returns a  **[FontNames](fontnames-object-word.md)** object that includes the names of all the available portrait fonts.|
|[PrintPreview](application-printpreview-property-word.md)| **True** if print preview is the current view. Read/write **Boolean** .|
|[ProtectedViewWindows](application-protectedviewwindows-property-word.md)|Returns a [ProtectedViewWindows](protectedviewwindows-object-word.md) collection that represents all protected view windows. Read-only.|
|[RecentFiles](application-recentfiles-property-word.md)|Returns a  **[RecentFiles](recentfiles-object-word.md)** collection that represents the most recently accessed files.|
|[RestrictLinkedStyles](application-restrictlinkedstyles-property-word.md)|Returns or sets a  **Boolean** that represents whether Microsoft Word allows linked styles. Read/write.|
|[ScreenUpdating](application-screenupdating-property-word.md)| **True** if screen updating is turned on. Read/write **Boolean** .|
|[Selection](application-selection-property-word.md)|Returns the  **[Selection](selection-object-word.md)** object that represents a selected range or the insertion point. Read-only.|
|[ShowAnimation](application-showanimation-property-word.md)|This object or member is deprecated and is not intended to be used in your code. |
|[ShowStartupDialog](application-showstartupdialog-property-word.md)| **True** to display the **Task Pane** when starting Microsoft Word. Read/write **Boolean** .|
|[ShowStylePreviews](application-showstylepreviews-property-word.md)|Returns or sets a  **Boolean** that represents whether Microsoft Word shows a preview of the formatting for styles in the **Styles** dialog box. Read/write.|
|[ShowVisualBasicEditor](application-showvisualbasiceditor-property-word.md)| **True** if the Visual Basic Editor window is visible. Read/write **Boolean** .|
|[SmartArtColors](application-smartartcolors-property-word.md)|Returns a [SmartArtColors](smartartcolors-object-office.md) object that represents the set of color styles that are currently loaded in the application. Read-only.|
|[SmartArtLayouts](application-smartartlayouts-property-word.md)|Returns a [SmartArtLayouts](smartartlayouts-object-office.md) object that represents the set of SmartArt layouts that are currently loaded in the application. Read-only.|
|[SmartArtQuickStyles](application-smartartquickstyles-property-word.md)|Returns a [SmartArtQuickStyles](smartartquickstyles-object-office.md) object that represents the set of SmartArt styles that are currently loaded in the application. Read-only.|
|[SpecialMode](application-specialmode-property-word.md)| **True** if Microsoft Word is in a special mode (for example, CopyText mode, or MoveText mode). Read-only **Boolean** .|
|[StartupPath](application-startuppath-property-word.md)|Returns or sets the complete path of the startup folder, excluding the final separator. Read/write  **String** .|
|[StatusBar](application-statusbar-property-word.md)|This property is no longer supported in Microsoft Word Visual Basic for Applications.|
|[SynonymInfo](application-synonyminfo-property-word.md)|Returns a  **[SynonymInfo](synonyminfo-object-word.md)** object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the specified word or phrase.|
|[System](application-system-property-word.md)|Returns a  **[System](system-object-word.md)** object, which can be used to return system-related information and perform system-related tasks.|
|[TaskPanes](application-taskpanes-property-word.md)|Returns a  **[TaskPanes](taskpanes-object-word.md)** collection that represents the most commonly performed tasks in Microsoft Word.|
|[Tasks](application-tasks-property-word.md)|Returns a  **[Tasks](tasks-object-word.md)** collection that represents all the applications that are running.|
|[Templates](application-templates-property-word.md)|Returns a  **[Templates](templates-object-word.md)** collection that represents all the available templates—global templates and those attached to open documents.|
|[Top](application-top-property-word.md)|Returns or sets the vertical position of the active document. Read/write  **Long** .|
|[UndoRecord](application-undorecord-property-word.md)|Returns an [UndoRecord](undorecord-object-word.md) object that provides a custom entry point into the undo stack. Read-only.|
|[UsableHeight](application-usableheight-property-word.md)|Returns the maximum height (in points) to which you can set the height of a Microsoft Word document window. Read-only  **Long** .|
|[UsableWidth](application-usablewidth-property-word.md)|Returns the maximum width (in points) to which you can set the width of a Microsoft Word document window. Read-only  **Long** .|
|[UserAddress](application-useraddress-property-word.md)|Returns or sets the user's mailing address. Read/write  **String** .|
|[UserControl](application-usercontrol-property-word.md)| **True** if the document or application was created or opened by the user. Read-only **Boolean** .|
|[UserInitials](application-userinitials-property-word.md)|Returns or sets the user's initials, which Microsoft Word uses to construct comment marks. Read/write  **String** .|
|[UserName](application-username-property-word.md)|Returns or sets the user's name, which is used on envelopes and for the Author document property. Read/write  **String** .|
|[VBE](application-vbe-property-word.md)|Returns a VBE object that represents the Visual Basic Editor.|
|[Version](application-version-property-word.md)|Returns the Microsoft Word version number. Read-only  **String** .|
|[Visible](application-visible-property-word.md)| **True** if the specified object is visible. Read/write **Boolean** .|
|[Width](application-width-property-word.md)|Returns or sets the width of the application window, in points. Read/write  **Long** .|
|[Windows](application-windows-property-word.md)|Returns a  **[Windows](windows-object-word.md)** collection that represents all document windows. Read-only.|
|[WindowState](application-windowstate-property-word.md)|Returns or sets the state of the specified document window or task window. Read/write  **WdWindowState** .|
|[WordBasic](application-wordbasic-property-word.md)|Returns an automation object (Word.Basic) that includes methods for all the WordBasic statements and functions available in Word version 6.0 and Word for Windows 95. Read-only.|
|[XMLNamespaces](application-xmlnamespaces-property-word.md)|Returns an  **** collection that represents the XML schemas in the Schema Library.|

