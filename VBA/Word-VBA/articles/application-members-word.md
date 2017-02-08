---
title: Application Members (Word)
ms.prod: WORD
ms.assetid: 71669f1e-65f1-b0f1-b67d-355dfdbebe50
---


# Application Members (Word)
Represents the Microsoft Word application. The  **Application** object includes properties and methods that return top-level objects. For example, the **[ActiveDocument](application-activedocument-property-word.md)** property returns a **[Document](document-object-word.md)** object.

Represents the Microsoft Word application. The  **Application** object includes properties and methods that return top-level objects. For example, the **[ActiveDocument](application-activedocument-property-word.md)** property returns a **[Document](document-object-word.md)** object.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[DocumentBeforeClose](application-documentbeforeclose-event-word.md)|Occurs immediately before any open document closes.|
|[DocumentBeforePrint](application-documentbeforeprint-event-word.md)|Occurs before any open document is printed.|
|[DocumentBeforeSave](application-documentbeforesave-event-word.md)|Occurs before any open document is saved.|
|[DocumentChange](application-documentchange-event-word.md)|Occurs when a new document is created, when an existing document is opened, or when another document is made the active document.|
|[DocumentOpen](application-documentopen-event-word.md)|Occurs when a document is opened.|
|[DocumentSync](application-documentsync-event-word.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|
|[EPostageInsert](application-epostageinsert-event-word.md)|Occurs when a user inserts electronic postage into a document.|
|[EPostageInsertEx](application-epostageinsertex-event-word.md)|Occurs when a user inserts electronic postage into a document.|
|[EPostagePropertyDialog](application-epostagepropertydialog-event-word.md)|Occurs when a user clicks the  **E-postage Properties** ( **Labels and Envelopes** dialog box) button or **Print Electronic Postage** button.|
|[MailMergeAfterMerge](application-mailmergeaftermerge-event-word.md)|Occurs after all records in a mail merge have merged successfully.|
|[MailMergeAfterRecordMerge](application-mailmergeafterrecordmerge-event-word.md)|Occurs after each record in the data source successfully merges in a mail merge.|
|[MailMergeBeforeMerge](application-mailmergebeforemerge-event-word.md)|Occurs when a merge is executed before any records merge.|
|[MailMergeBeforeRecordMerge](application-mailmergebeforerecordmerge-event-word.md)|Occurs as a merge is executed for the individual records in a merge.|
|[MailMergeDataSourceLoad](application-mailmergedatasourceload-event-word.md)|Occurs when the data source is loaded for a mail merge.|
|[MailMergeDataSourceValidate](application-mailmergedatasourcevalidate-event-word.md)|Occurs when a user validates mail merge recipients by clicking  **Validate** in the **Mail Merge Recipients** dialog box.|
|[MailMergeDataSourceValidate2](application-mailmergedatasourcevalidate2-event-word.md)|Occurs when a user validates mail merge recipients by clicking the  **Validate addresses** link button in the **Mail Merge Recipients** dialog box.|
|[MailMergeWizardSendToCustom](application-mailmergewizardsendtocustom-event-word.md)|Occurs when the custom button is clicked during step six of the Mail Merge Wizard.|
|[MailMergeWizardStateChange](application-mailmergewizardstatechange-event-word.md)|Occurs when a user changes from a specified step to a specified step in the Mail Merge Wizard.|
|[NewDocument](application-newdocument-event-word.md)|Occurs when a new document is created.|
|[ProtectedViewWindowActivate](application-protectedviewwindowactivate-event-word.md)|Occurs when any protected view window is activated.|
|[ProtectedViewWindowBeforeClose](application-protectedviewwindowbeforeclose-event-word.md)|Occurs immediately before a protected view window or a document in a protected view window closes.|
|[ProtectedViewWindowBeforeEdit](application-protectedviewwindowbeforeedit-event-word.md)|Occurs immediately before editing is enabled on the document in the specified protected view window.|
|[ProtectedViewWindowDeactivate](application-protectedviewwindowdeactivate-event-word.md)|Occurs when a protected view window is deactivated.|
|[ProtectedViewWindowOpen](application-protectedviewwindowopen-event-word.md)|Occurs when a protected view window is opened.|
|[ProtectedViewWindowSize](application-protectedviewwindowsize-event-word.md)||
|[Quit](application-quit-event-word.md)|Occurs when the user exits Microsoft Word.|
|[WindowActivate](application-windowactivate-event-word.md)|Occurs when any document window is activated.|
|[WindowBeforeDoubleClick](application-windowbeforedoubleclick-event-word.md)|Occurs when the editing area of a document window is double-clicked, before the default double-click action.|
|[WindowBeforeRightClick](application-windowbeforerightclick-event-word.md)|Occurs when the editing area of a document window is right-clicked, before the default right-click action.|
|[WindowDeactivate](application-windowdeactivate-event-word.md)|Occurs when any document window is deactivated.|
|[WindowSelectionChange](application-windowselectionchange-event-word.md)|Occurs when the selection changes in the active document window.|
|[WindowSize](application-windowsize-event-word.md)|Occurs when the application window is resized or moved.|
|[XMLSelectionChange](application-xmlselectionchange-event-word.md)|Occurs when the parent XML node of the current selection changes.|
|[XMLValidationError](application-xmlvalidationerror-event-word.md)|Occurs when there is a validation error in the document.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](application-activate-method-word.md)|Activates the specified object.|
|[AddAddress](application-addaddress-method-word.md)|Adds an entry to the address book. Each entry has values for one or more tag IDs.|
|[AutomaticChange](application-automaticchange-method-word.md)|Performs an  **AutoFormat** action when there is a change suggested by the Office Assistant. If no AutoFormat action is active, this method generates an error.|
|[BuildKeyCode](application-buildkeycode-method-word.md)|Returns a unique number for the specified key combination.|
|[CentimetersToPoints](application-centimeterstopoints-method-word.md)|Converts a measurement from centimeters to points (1 cm = 28.35 points). Returns the converted measurement as a  **Single** .|
|[ChangeFileOpenDirectory](application-changefileopendirectory-method-word.md)|Sets the folder in which Word searches for documents.|
|[CheckGrammar](application-checkgrammar-method-word.md)|Checks a string for grammatical errors. Returns a  **Boolean** to indicate whether the string contains grammatical errors. **True** if the string contains no errors.|
|[CheckSpelling](application-checkspelling-method-word.md)|Checks a string for spelling errors. Returns a  **Boolean** to indicate whether the string contains spelling errors. **True** if the string has no spelling errors.|
|[CleanString](application-cleanstring-method-word.md)|Removes nonprinting characters (character codes 1 ? 29) and special Word characters from the specified string or changes them to spaces (character code 32). Returns the result as a  **String** .|
|[CompareDocuments](application-comparedocuments-method-word.md)|Compares two documents and returns a  **Document** object that represents the document that contains the differences between the two documents, marked using tracked changes.|
|[DDEExecute](application-ddeexecute-method-word.md)|Sends a command or series of commands to an application through the specified dynamic data exchange (DDE) channel.|
|[DDEInitiate](application-ddeinitiate-method-word.md)|Opens a dynamic data exchange (DDE) channel to another application, and returns the channel number.|
|[DDEPoke](application-ddepoke-method-word.md)|Uses an open dynamic data exchange (DDE) channel to send data to an application.|
|[DDERequest](application-dderequest-method-word.md)|Uses an open dynamic data exchange (DDE) channel to request information from the receiving application, and returns the information as a  **String** .|
|[DDETerminate](application-ddeterminate-method-word.md)|Closes the specified dynamic data exchange (DDE) channel to another application.|
|[DDETerminateAll](application-ddeterminateall-method-word.md)|Closes all dynamic data exchange (DDE) channels opened by Microsoft Word.|
|[DefaultWebOptions](application-defaultweboptions-method-word.md)|Returns the  **[DefaultWebOptions](defaultweboptions-object-word.md)** object that contains global application-level attributes used by Microsoft Word whenever you save a document as a Web page or open a Web page.|
|[GetAddress](application-getaddress-method-word.md)|Returns an address from the default address book.|
|[GetDefaultTheme](application-getdefaulttheme-method-word.md)|Returns a  **String** that represents the name of the default theme plus the theme formatting options Microsoft Word uses for new documents, e-mail messages, or Web pages.|
|[GetSpellingSuggestions](application-getspellingsuggestions-method-word.md)|Returns a  **[SpellingSuggestions](spellingsuggestions-object-word.md)** collection that represents the words suggested as spelling replacements for a given word.|
|[GoBack](application-goback-method-word.md)|Moves the insertion point among the last three locations where editing occurred in the active document (the same as pressing SHIFT+F5).|
|[GoForward](application-goforward-method-word.md)|Moves the insertion point forward among the last three locations where editing occurred in the active document.|
|[Help](application-help-method-word.md)|Displays installed Help information.|
|[HelpTool](application-helptool-method-word.md)||
|[InchesToPoints](application-inchestopoints-method-word.md)|Converts a measurement from inches to points (1 inch = 72 points). Returns the converted measurement as a  **Single** .|
|[Keyboard](application-keyboard-method-word.md)|Returns or sets the keyboard language and layout settings.|
|[KeyboardBidi](application-keyboardbidi-method-word.md)|Sets the keyboard language to a right-to-left language and the text entry direction to right-to-left.|
|[KeyboardLatin](application-keyboardlatin-method-word.md)|Sets the keyboard language to a left-to-right language and the text entry direction to left-to-right.|
|[KeyString](application-keystring-method-word.md)|Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A).|
|[LinesToPoints](application-linestopoints-method-word.md)|Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a  **Single** .|
|[ListCommands](application-listcommands-method-word.md)|Creates a new document and then inserts a table of Word commands along with their associated shortcut keys and menu assignments.|
|[LoadMasterList](application-loadmasterlist-method-word.md)|Loads a bibliography source file.|
|[LookupNameProperties](application-lookupnameproperties-method-word.md)|Looks up a name in the global address book list and displays the  **Properties** dialog box, which includes information about the specified name.|
|[MergeDocuments](application-mergedocuments-method-word.md)|Compares two documents and returns a  **Document** object that represents the document that contains the differences between the two documents, marked using tracked changes.|
|[MillimetersToPoints](application-millimeterstopoints-method-word.md)|Converts a measurement from millimeters to points (1 mm = 2.85 points). Returns the converted measurement as a  **Single** .|
|[Move](application-move-method-word.md)|Positions a task window or the active document window.|
|[NewWindow](application-newwindow-method-word.md)|Opens a new window with the same document as the specified window. Returns a  **Window** object.|
|[NextLetter](application-nextletter-method-word.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about the  **NextLetter** method of the **Application** object, consult the language reference Help included with Microsoft Office Macintosh Edition.|
|[OnTime](application-ontime-method-word.md)|Starts a background timer that runs a macro at a specified time.|
|[OrganizerCopy](application-organizercopy-method-word.md)|Copies the specified AutoText entry, toolbar, style, or macro project item from the source document or template to the destination document or template.|
|[OrganizerDelete](application-organizerdelete-method-word.md)|Deletes the specified style, AutoText entry, toolbar, or macro project item from a document or template.|
|[OrganizerRename](application-organizerrename-method-word.md)|Renames the specified style, AutoText entry, toolbar, or macro project item in a document or template.|
|[PicasToPoints](application-picastopoints-method-word.md)|Converts a measurement from picas to points (1 pica = 12 points). Returns the converted measurement as a  **Single** .|
|[PixelsToPoints](application-pixelstopoints-method-word.md)|Converts a measurement from pixels to points. Returns the converted measurement as a  **Single** .|
|[PointsToCentimeters](application-pointstocentimeters-method-word.md)|Converts a measurement from points to centimeters (1 centimeter = 28.35 points). Returns the converted measurement as a  **Single** .|
|[PointsToInches](application-pointstoinches-method-word.md)|Converts a measurement from points to inches (1 inch = 72 points). Returns the converted measurement as a  **Single** .|
|[PointsToLines](application-pointstolines-method-word.md)|Converts a measurement from points to lines (1 line = 12 points). Returns the converted measurement as a  **Single** .|
|[PointsToMillimeters](application-pointstomillimeters-method-word.md)|Converts a measurement from points to millimeters (1 millimeter = 2.835 points). Returns the converted measurement as a  **Single** .|
|[PointsToPicas](application-pointstopicas-method-word.md)|Converts a measurement from points to picas (1 pica = 12 points). Returns the converted measurement as a  **Single** .|
|[PointsToPixels](application-pointstopixels-method-word.md)|Converts a measurement from points to pixels. Returns the converted measurement as a  **Single** .|
|[PrintOut](application-printout-method-word.md)|Prints all or part of the specified document.|
|[ProductCode](application-productcode-method-word.md)|Returns the Microsoft Word globally unique identifier (GUID) as a  **String** .|
|[PutFocusInMailHeader](application-putfocusinmailheader-method-word.md)|Places the insertion point in the  **To**line of the mail header if the document in the active window is an e-mail document.|
|[Quit](application-quit-method-word.md)|Quits Microsoft Word and optionally saves or routes the open documents.|
|[Repeat](application-repeat-method-word.md)|Repeats the most recent editing action one or more times. Returns  **True** if the commands were repeated successfully.|
|[ResetIgnoreAll](application-resetignoreall-method-word.md)|Clears the list of words that were previously ignored during a spelling check.|
|[Resize](application-resize-method-word.md)|Sizes the Word application window or the specified task window.|
|[Run](application-run-method-word.md)|Runs a Visual Basic macro.|
|[ScreenRefresh](application-screenrefresh-method-word.md)|Updates the display on the monitor with the current information in the video memory buffer.|
|[SetDefaultTheme](application-setdefaulttheme-method-word.md)|Sets a default theme for Microsoft Word to use with new documents, e-mail messages, or Web pages.|
|[ShowClipboard](application-showclipboard-method-word.md)|Displays the  **Clipboard** task pane.|
|[ShowMe](application-showme-method-word.md)||
|[SubstituteFont](application-substitutefont-method-word.md)|Sets font-mapping options.|
|[ToggleKeyboard](application-togglekeyboard-method-word.md)|Switches the keyboard language setting between right-to-left and left-to-right languages.|

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

