---
title: Document Properties (Word)
ms.prod: WORD
ms.assetid: edd1a58d-10a0-47c7-b2df-7e49a41d8a6b
---


# Document Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActiveTheme](document-activetheme-property-word.md)|Returns the name of the active theme plus the theme formatting options for the specified document. Read-only  **String** .|
|[ActiveThemeDisplayName](document-activethemedisplayname-property-word.md)|Returns the display name of the active theme for the specified document. Read-only  **String** .|
|[ActiveWindow](document-activewindow-property-word.md)|Returns a  **[Window](window-object-word.md)** object that represents the active window (the window with the focus). Read-only.|
|[ActiveWritingStyle](document-activewritingstyle-property-word.md)|Returns or sets the writing style for a specified language in the specified document. Read/write  **String** .|
|[Application](document-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AttachedTemplate](document-attachedtemplate-property-word.md)|Returns a  **[Template](template-object-word.md)** object that represents the template attached to the specified document. Read/write **Variant** .|
|[AutoFormatOverride](document-autoformatoverride-property-word.md)|Returns or sets a  **Boolean** that represents whether automatic formatting options override formatting restrictions in a document where formatting restrictions are in effect.|
|[AutoHyphenation](document-autohyphenation-property-word.md)| **True** if automatic hyphenation is turned on for the specified document. Read/write **Boolean** .|
|[Background](document-background-property-word.md)|Returns a  **Shape** object that represents the background image for the specified document. Read-only.|
|[Bibliography](document-bibliography-property-word.md)|Returns a  **[Bibliography](bibliography-object-word.md)** object that represents the bibliography references contained within a document. Read-only.|
|[Bookmarks](document-bookmarks-property-word.md)|Returns a  **[Bookmarks](bookmarks-object-word.md)** collection that represents all the bookmarks in a document. Read-only.|
|[Broadcast](document-broadcast-property-word.md)|Returns a [Broadcast](broadcast-object-word.md) object that represents a broadcast session, in which presenters can present Word documents to remote participants over the web without the participants needing to have rich clients installed.|
|[BuiltInDocumentProperties](document-builtindocumentproperties-property-word.md)|Returns a  **DocumentProperties** collection that represents all the built-in document properties for the specified document.|
|[Characters](document-characters-property-word.md)|Returns a  **[Characters](characters-object-word.md)** collection that represents the characters in a document. Read-only.|
|[ChartDataPointTrack](document-chartdatapointtrack-property-word.md)|Returns or sets a  **Boolean** that specifies whether charts in the active document use cell-reference data-point tracking. Read-write.|
|[ClickAndTypeParagraphStyle](document-clickandtypeparagraphstyle-property-word.md)|Returns or sets the default paragraph style applied to text by the Click and Type feature in the specified document. Read/write  **Variant** .|
|[CoAuthoring](document-coauthoring-property-word.md)|Returns a [CoAuthoring](coauthoring-object-word.md) object that provides the entry point into the co-authoring object model. Read-only.|
|[CodeName](document-codename-property-word.md)|Returns the code name for the specified document. Read-only  **String** .|
|[CommandBars](document-commandbars-property-word.md)|Returns a  **CommandBars** collection that represents the menu bar and all the toolbars in Microsoft Word.|
|[Comments](document-comments-property-word.md)|Returns a  **[Comments](comments-object-word.md)** collection that represents all the comments in the specified document. Read-only.|
|[Compatibility](document-compatibility-property-word.md)| **True** if the compatibility option specified by the Type argument is enabled. Compatibility options affect how a document is displayed in Microsoft Word. Read/write **Boolean** .|
|[CompatibilityMode](document-compatibilitymode-property-word.md)|Returns a  **Long** that specifies the compatibility mode that Word uses when opening the document. Read-only.|
|[ConsecutiveHyphensLimit](document-consecutivehyphenslimit-property-word.md)|Returns or sets the maximum number of consecutive lines that can end with hyphens. Read/write.  **Long** .|
|[Container](document-container-property-word.md)|Returns the object that represents the container application for the specified document. Read-only  **Object** .|
|[Content](document-content-property-word.md)|Returns a  **[Range](range-object-word.md)** object that represents the main document story. Read-only.|
|[ContentControls](document-contentcontrols-property-word.md)|Returns a  **[ContentControls](contentcontrols-object-word.md)** collection that represents all the content controls in a document. Read-only.|
|[ContentTypeProperties](document-contenttypeproperties-property-word.md)|Returns a  **MetaProperties** collection that represents the metadata stored in a document, such as author name, subject, and company. Read-only.|
|[Creator](document-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[CurrentRsid](document-currentrsid-property-word.md)|Returns a  **Long** that represents a random number that Word assigns to changes in a document. Read-only.|
|[CustomDocumentProperties](document-customdocumentproperties-property-word.md)|Returns a  **DocumentProperties** collection that represents all the custom document properties for the specified document.|
|[CustomXMLParts](document-customxmlparts-property-word.md)|Returns a  **[CustomXMLParts](customxmlparts-object-office.md)** collection that represents the custom XML in the XML data store. Read-only.|
|[DefaultTableStyle](document-defaulttablestyle-property-word.md)|Returns a  **Variant** that represents the table style that is applied to all newly created tables in a document. Read-only.|
|[DefaultTabStop](document-defaulttabstop-property-word.md)|Returns or sets the interval (in points) between the default tab stops in the specified document. Read/write  **Single** .|
|[DefaultTargetFrame](document-defaulttargetframe-property-word.md)|Returns or sets a  **String** indicating the browser frame in which to display a Web page reached through a hyperlink. Read/write.|
|[DisableFeatures](document-disablefeatures-property-word.md)| **True** disables all features introduced after the version specified in the **[DisableFeaturesIntroducedAfter](document-disablefeaturesintroducedafter-property-word.md)** property. The default value is **False** . Read/write **Boolean** .|
|[DisableFeaturesIntroducedAfter](document-disablefeaturesintroducedafter-property-word.md)|Disables all features introduced after a specified version of Microsoft Word in the document only. Read/write  **WdDisableFeaturesIntroducedAfter** .|
|[DocumentInspectors](document-documentinspectors-property-word.md)|Returns a  **DocumentInspectors** collection that enables you to locate hidden personal information, such as author name, company name, and revision date. Read-only.|
|[DocumentLibraryVersions](document-documentlibraryversions-property-word.md)|Returns a  **DocumentLibraryVersions** collection that represents the collection of versions of a shared document that has versioning enabled and that is stored in a document library on a server.|
|[DocumentTheme](document-documenttheme-property-word.md)|Returns an  **OfficeTheme** object that represents the Microsoft Office theme applied to a document. Read-only.|
|[DoNotEmbedSystemFonts](document-donotembedsystemfonts-property-word.md)| **True** for Microsoft Word to not embed common system fonts. Read/write **Boolean** .|
|[Email](document-email-property-word.md)|Returns an  **[Email](email-object-word.md)** object that contains all the e-mail-related properties of the current document. Read-only.|
|[EmbedLinguisticData](document-embedlinguisticdata-property-word.md)| **True** for Microsoft Word to embed speech and handwriting so that data can be converted back to speech or handwriting. Read/write **Boolean** .|
|[EmbedTrueTypeFonts](document-embedtruetypefonts-property-word.md)| **True** if Microsoft Word embeds TrueType fonts in a document when it is saved. Read/write **Boolean** .|
|[EncryptionProvider](document-encryptionprovider-property-word.md)|Returns a  **String** specifying the name of the algorithm encryption provider that Microsoft Word uses when encrypting documents. Read/write.|
|[Endnotes](document-endnotes-property-word.md)|Returns an  **[Endnotes](endnotes-object-word.md)** collection that represents all the endnotes in a document. Read-only.|
|[EnforceStyle](document-enforcestyle-property-word.md)|Returns or sets a  **Boolean** that represents whether formatting restrictions are enforced in a protected document.|
|[Envelope](document-envelope-property-word.md)|Returns an  **[Envelope](envelope-object-word.md)** object that represents an envelope and envelope features in a document. Read-only.|
|[FarEastLineBreakLanguage](document-fareastlinebreaklanguage-property-word.md)|Returns or sets a  **[WdFarEastLineBreakLanguageID](wdfareastlinebreaklanguageid-enumeration-word.md)** that represents the East Asian language to use when breaking lines of text in the specified document or template. Read/write.|
|[FarEastLineBreakLevel](document-fareastlinebreaklevel-property-word.md)|Returns or sets a  **WdFarEastLineBreakLevel** that represents the line break control level for the specified document. Read/write.|
|[Fields](document-fields-property-word.md)|Returns a  **[Fields](fields-object-word.md)** collection that represents all the fields in the document. Read-only.|
|[Final](document-final-property-word.md)|Returns or sets a  **Boolean** that indicates whether a document is final. Read/write.|
|[Footnotes](document-footnotes-property-word.md)|Returns a  **[Footnotes](footnotes-object-word.md)** collection that represents all the footnotes in a document. Read-only.|
|[FormattingShowClear](document-formattingshowclear-property-word.md)| **True** for Microsoft Word to show clear formatting in the **Styles and Formatting** task pane. Read/write **Boolean** .|
|[FormattingShowFilter](document-formattingshowfilter-property-word.md)|Sets or returns a  **WdShowFilter** constant that represents the styles and formatting displayed in the **Styles and Formatting** task pane. Read/write **Boolean** .|
|[FormattingShowFont](document-formattingshowfont-property-word.md)| **True** for Microsoft Word to display font formatting in the **Styles and Formatting** task pane. Read/write **Boolean** .|
|[FormattingShowNextLevel](document-formattingshownextlevel-property-word.md)|Returns or sets a  **Boolean** that represents whether Microsoft Word shows the next heading level when the previous heading level is used. Read/write.|
|[FormattingShowNumbering](document-formattingshownumbering-property-word.md)| **True** for Microsoft Word to display number formatting in the **Styles and Formatting** task pane. Read/write **Boolean** .|
|[FormattingShowParagraph](document-formattingshowparagraph-property-word.md)| **True** for Microsoft Word to display paragraph formatting in the **Styles and Formatting** task pane. Read/write **Boolean** .|
|[FormattingShowUserStyleName](document-formattingshowuserstylename-property-word.md)|Returns or sets a  **Boolean** that represents whether to show user-defined styles. Read/write.|
|[FormFields](document-formfields-property-word.md)|Returns a  **[FormFields](formfields-object-word.md)** collection that represents all the form fields in the document. Read-only.|
|[FormsDesign](document-formsdesign-property-word.md)| **True** if the specified document is in form design mode. Read-only **Boolean** .|
|[Frames](document-frames-property-word.md)|Returns a  **[Frames](frames-object-word.md)** collection that represents all the frames in a document. Read-only.|
|[Frameset](document-frameset-property-word.md)|Returns a  **[Frameset](frameset-object-word.md)** object that represents an entire frames page or a single frame on a frames page. Read-only.|
|[FullName](document-fullname-property-word.md)|Returns a  **String** that represents the name of a document, including the path. Read-only.|
|[GrammarChecked](document-grammarchecked-property-word.md)| **True** if a grammar check has been run on the specified range or document. Read/write **Boolean** .|
|[GrammaticalErrors](document-grammaticalerrors-property-word.md)|Returns a  **[ProofreadingErrors](proofreadingerrors-object-word.md)** collection that represents the sentences that failed the grammar check in the specified document. Read-only.|
|[GridDistanceHorizontal](document-griddistancehorizontal-property-word.md)|Returns or sets a  **Single** that represents the amount of horizontal space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize AutoShapes or East Asian characters in the specified document. Read/write.|
|[GridDistanceVertical](document-griddistancevertical-property-word.md)|Returns or sets a  **Single** that represents the amount of vertical space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize AutoShapes or East Asian characters in the specified document. Read/write.|
|[GridOriginFromMargin](document-gridoriginfrommargin-property-word.md)| **True** if Microsoft Word starts the character grid from the upper-left corner of the page. Read/write **Boolean** .|
|[GridOriginHorizontal](document-gridoriginhorizontal-property-word.md)|Returns or sets a  **Single** that represents the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in the specified document. Read/write.|
|[GridOriginVertical](document-gridoriginvertical-property-word.md)|Returns or sets a  **Single** that represents the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in the specified document. Read/write.|
|[GridSpaceBetweenHorizontalLines](document-gridspacebetweenhorizontallines-property-word.md)|Returns or sets the interval at which Microsoft Word displays horizontal character gridlines in print layout view. Read/write  **Long** .|
|[GridSpaceBetweenVerticalLines](document-gridspacebetweenverticallines-property-word.md)|Returns or sets the interval at which Microsoft Word displays vertical character gridlines in print layout view. Read/write  **Long** .|
|[HasPassword](document-haspassword-property-word.md)| **True** if a password is required to open the specified document. Read-only **Boolean** .|
|[HasVBProject](document-hasvbproject-property-word.md)|Returns a  **Boolean** that represents whether a document has an attached Microsoft Visual Basic for Applications project. Read-only.|
|[HTMLDivisions](document-htmldivisions-property-word.md)|Returns an  **[HTMLDivisions](htmldivisions-object-word.md)** collection that represents the HTML DIV elements in a Web document.|
|[Hyperlinks](document-hyperlinks-property-word.md)|Returns a  **[Hyperlinks](hyperlinks-object-word.md)** collection that represents all the hyperlinks in the specified document. Read-only.|
|[HyphenateCaps](document-hyphenatecaps-property-word.md)| **True** if words in all capital letters can be hyphenated. Read/write **Boolean** .|
|[HyphenationZone](document-hyphenationzone-property-word.md)|Returns or sets the width of the hyphenation zone, in points. Read/write  **Long** .|
|[Indexes](document-indexes-property-word.md)|Returns an  **[Indexes](indexes-object-word.md)** collection that represents all the indexes in the specified document. Read-only.|
|[InlineShapes](document-inlineshapes-property-word.md)|Returns an  **[InlineShapes](document-inlineshapes-property-word.md)** collection that represents all the **[InlineShape](inlineshape-object-word.md)** objects in a document. Read-only.|
|[IsInAutosave](document-isinautosave-property-word.md)|Returns  **False** if the most recent firing of the[Application.DocumentBeforeSave Event (Word)](application-documentbeforesave-event-word.md) event was the result of a manual save by the user, and not an automatic save. Read-only.|
|[IsMasterDocument](document-ismasterdocument-property-word.md)| **True** if the specified document is a master document. Read-only **Boolean** .|
|[IsSubdocument](document-issubdocument-property-word.md)| **True** if the specified document is a subdocument of a master document. Read-only **Boolean** .|
|[JustificationMode](document-justificationmode-property-word.md)|Returns or sets the character spacing adjustment for the specified document. Read/write  **[WdJustificationMode](wdjustificationmode-enumeration-word.md)** .|
|[KerningByAlgorithm](document-kerningbyalgorithm-property-word.md)| **True** if Microsoft Word kerns half-width Latin characters and punctuation marks in the specified document. Read/write **Boolean** .|
|[Kind](document-kind-property-word.md)|Returns or sets the format type that Microsoft Word uses when automatically formatting the specified document. Read/write  **[WdDocumentKind](wddocumentkind-enumeration-word.md)** .|
|[LanguageDetected](document-languagedetected-property-word.md)|Returns or sets a value that specifies whether Microsoft Word has detected the language of the specified text. Read/write  **Boolean** .|
|[ListParagraphs](document-listparagraphs-property-word.md)|Returns a  **ListParagraphs** object that represents all the numbered paragraphs in a document. Read-only.|
|[Lists](document-lists-property-word.md)|Returns a  **[Lists](lists-object-word.md)** collection that contains all the formatted lists in the specified document. Read-only.|
|[ListTemplates](document-listtemplates-property-word.md)|Returns a  **ListTemplates** collection that represents all the list formats for the specified document. Read-only.|
|[LockQuickStyleSet](document-lockquickstyleset-property-word.md)|Returns or sets a  **Boolean** that represents whether users can change which set of Quick Styles is being used. Read/write.|
|[LockTheme](document-locktheme-property-word.md)|Returns or sets a  **Boolean** that represents whether a user can change a document theme. Read/write.|
|[MailEnvelope](document-mailenvelope-property-word.md)|Returns an  **MsoEnvelope** object that represents an e-mail header for a document.|
|[MailMerge](document-mailmerge-property-word.md)|Returns a  **[MailMerge](mailmerge-object-word.md)** object that represents the mail merge functionality for the specified document. Read-only.|
|[Name](document-name-property-word.md)|Returns the name of the specified object. Read-only  **String** .|
|[NoLineBreakAfter](document-nolinebreakafter-property-word.md)|Returns or sets the kinsoku characters after which Microsoft Word will not break a line. Read/write  **String** .|
|[NoLineBreakBefore](document-nolinebreakbefore-property-word.md)|Returns or sets the kinsoku characters before which Microsoft Word will not break a line. Read/write  **String** .|
|[OMathBreakBin](document-omathbreakbin-property-word.md)|Returns or sets a  **[WdOMathBreakBin](wdomathbreakbin-enumeration-word.md)** constant that represents where Microsoft Word places binary operators when equations span two or more lines. Read/write.|
|[OMathBreakSub](document-omathbreaksub-property-word.md)|Returns or sets a  **[WdOMathBreakSub](wdomathbreaksub-enumeration-word.md)** constant that represents how Microsoft Word handles a subtraction operator that falls before a line break. Read/write.|
|[OMathFontName](document-omathfontname-property-word.md)|Returns or sets a  **String** that represents the name of the font used in a document to display equations. Read/write.|
|[OMathIntSubSupLim](document-omathintsubsuplim-property-word.md)|Returns or sets a  **Boolean** that represents the default location of limits for integrals. Read/write.|
|[OMathJc](document-omathjc-property-word.md)|Returns or sets a  **[WdOMathJc](wdomathjc-enumeration-word.md)** constant that represents the default justification—left, right, centered, or centered as a group—of a group of equations. Read/write.|
|[OMathLeftMargin](document-omathleftmargin-property-word.md)|Returns or sets a  **Single** that represents the left margin for equations. Read/write.|
|[OMathNarySupSubLim](document-omathnarysupsublim-property-word.md)|Returns or sets a  **Boolean** that represents the default location of limits for n-ary objects other than integrals. Read/write.|
|[OMathRightMargin](document-omathrightmargin-property-word.md)|Returns or sets a  **Single** that represents the right margin for equations. Read/write.|
|[OMaths](document-omaths-property-word.md)|Returns an  **[OMaths](omaths-object-word.md)** collection that represents the **[OMath](omath-object-word.md)** objects within the specified range. Read-only.|
|[OMathSmallFrac](document-omathsmallfrac-property-word.md)|Returns or sets a  **Boolean** that represents whether to use small fractions in equations contained within the document. Read/write.|
|[OMathWrap](document-omathwrap-property-word.md)|Returns or sets a  **Single** that represents the placement of the second line of an equation that wraps to a new line. Read/write.|
|[OpenEncoding](document-openencoding-property-word.md)|Returns the encoding used to open the specified document. Read-only  **MsoEncoding** .|
|[OptimizeForWord97](document-optimizeforword97-property-word.md)| **True** if Microsoft Word optimizes the current document for viewing in Microsoft Word 97 by disabling any incompatible formatting. Read/write **Boolean** .|
|[OriginalDocumentTitle](document-originaldocumenttitle-property-word.md)|Returns a  **String** that represents the document title for the original document after running a legal-blackline document compare function. Read-only.|
|[PageSetup](document-pagesetup-property-word.md)|Returns a  **PageSetup** object that is associated with the specified document.|
|[Paragraphs](document-paragraphs-property-word.md)|Returns a  **Paragraphs** collection that represents all the paragraphs in the specified document. Read-only.|
|[Parent](document-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Document** object.|
|[Password](document-password-property-word.md)|Sets a password that must be supplied to open the specified document. Write-only  **String** .|
|[PasswordEncryptionAlgorithm](document-passwordencryptionalgorithm-property-word.md)|Returns a  **String** indicating the algorithm Microsoft Word uses for encrypting documents with passwords. Read-only.|
|[PasswordEncryptionFileProperties](document-passwordencryptionfileproperties-property-word.md)| **True** if Microsoft Word encrypts file properties for password-protected documents. Read-only **Boolean** .|
|[PasswordEncryptionKeyLength](document-passwordencryptionkeylength-property-word.md)|Returns a  **Long** indicating the key length of the algorithm Microsoft Word uses when encrypting documents with passwords. Read-only.|
|[PasswordEncryptionProvider](document-passwordencryptionprovider-property-word.md)|Returns a  **String** specifying the name of the algorithm encryption provider that Microsoft Word uses when encrypting documents with passwords. Read-only.|
|[Path](document-path-property-word.md)|Returns the disk or Web path to the document. Read-only  **String** .|
|[Permission](document-permission-property-word.md)|Returns a  **Permission** object that represents the permission settings in the specified document.|
|[PrintFormsData](document-printformsdata-property-word.md)| **True** if Microsoft Word prints onto a preprinted form only the data entered in the corresponding online form. Read/write **Boolean** .|
|[PrintPostScriptOverText](document-printpostscriptovertext-property-word.md)| **True** if PRINT field instructions (such as PostScript commands) in a document are to be printed on top of text and graphics when a PostScript printer is used. Read/write **Boolean** .|
|[PrintRevisions](document-printrevisions-property-word.md)| **True** if revision marks are printed with the document. **False** if revision marks aren't printed (that is, tracked changes are printed as if they'd been accepted). Read/write **Boolean** .|
|[ProtectionType](document-protectiontype-property-word.md)|Returns the protection type for the specified document. Can be one of the following  **WdProtectionType** constants: **wdAllowOnlyComments** , **wdAllowOnlyFormFields** , **wdAllowOnlyReading** , **wdAllowOnlyRevisions** , or **wdNoProtection** .|
|[ReadabilityStatistics](document-readabilitystatistics-property-word.md)|Returns a  **ReadabilityStatistics** collection that represents the readability statistics for the specified document or range. Read-only.|
|[ReadingLayoutSizeX](document-readinglayoutsizex-property-word.md)|Sets or returns a  **Long** that represents the width of pages in a document when it is displayed in reading layout view and is frozen for entering handwritten markup.|
|[ReadingLayoutSizeY](document-readinglayoutsizey-property-word.md)|Sets or returns a  **Long** that represents the height of pages in a document when it is displayed in reading layout view and is frozen for entering handwritten markup.|
|[ReadingModeLayoutFrozen](document-readingmodelayoutfrozen-property-word.md)|Sets or returns a  **Boolean** that represents whether pages displayed in reading layout view are frozen to a specified size for inserting handwritten markup into a document.|
|[ReadOnly](document-readonly-property-word.md)| **True** if changes to the document cannot be saved to the original document. Read-only **Boolean** .|
|[ReadOnlyRecommended](document-readonlyrecommended-property-word.md)| **True** if Microsoft Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only. Read/write **Boolean** .|
|[RemoveDateAndTime](document-removedateandtime-property-word.md)|Sets or returns a  **Boolean** indicating whether a document stores the date and time metadata for tracked changes. .|
|[RemovePersonalInformation](document-removepersonalinformation-property-word.md)| **True** if Microsoft Word removes all user information from comments, revisions, and the Properties dialog box upon saving a document. Read/write **Boolean** .|
|[Research](document-research-property-word.md)|Returns a  **Research** object that represents the research service for a document. Read-only.|
|[RevisedDocumentTitle](document-reviseddocumenttitle-property-word.md)|Returns a  **String** that represents the document title for a revised document after running a legal-blackline document compare function. Read-only.|
|[Revisions](document-revisions-property-word.md)|Returns a  **Revisions** collection that represents the tracked changes in the document or range. Read-only.|
|[Saved](document-saved-property-word.md)| **True** if the specified document or template has not changed since it was last saved. **False** if Microsoft Word displays a prompt to save changes when the document is closed. Read/write **Boolean** .|
|[SaveEncoding](document-saveencoding-property-word.md)|Returns or sets the encoding to use when saving a document. Read/write  **MsoEncoding** .|
|[SaveFormat](document-saveformat-property-word.md)|Returns the file format of the specified document or file converter. Read-only  **Long** .|
|[SaveFormsData](document-saveformsdata-property-word.md)| **True** if Microsoft Word saves the data entered in a form as a tab-delimited record for use in a database. Read/write **Boolean** .|
|[SaveSubsetFonts](document-savesubsetfonts-property-word.md)| **True** if Microsoft Word saves a subset of the embedded TrueType fonts with the document. Read/write **Boolean** .|
|[Scripts](document-scripts-property-word.md)|Returns a  **Scripts** collection that represents the collection of HTML scripts in the specified object.|
|[Sections](document-sections-property-word.md)|Returns a  **[Section](section-object-word.md)** collection that represents the sections in the specified document. Read-only.|
|[Sentences](document-sentences-property-word.md)|Returns a  **[Sentences](sentences-object-word.md)** collection that represents all the sentences in the document. Read-only.|
|[ServerPolicy](document-serverpolicy-property-word.md)|Returns a  **ServerPolicy** object that represents a policy specified for a document stored on a server running Microsoft Office SharePoint Server 2007. Read-only.|
|[Shapes](document-shapes-property-word.md)|Returns a  **[Shapes](shapes-object-word.md)** collection that represents all the **Shape** objects in the specified document. Read-only.|
|[ShowGrammaticalErrors](document-showgrammaticalerrors-property-word.md)| **True** if grammatical errors are marked by a wavy green line in the specified document. Read/write **Boolean** .|
|[ShowSpellingErrors](document-showspellingerrors-property-word.md)| **True** if Microsoft Word underlines spelling errors in the document. Read/write **Boolean** .|
|[Signatures](document-signatures-property-word.md)|Returns a  **SignatureSet** collection that represents the digital signatures for a document.|
|[SmartDocument](document-smartdocument-property-word.md)|Returns a  **SmartDocument** object that represents the settings for a smart document solution.|
|[SnapToGrid](document-snaptogrid-property-word.md)| **True** if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in the specified document. Read/write **Boolean** .|
|[SnapToShapes](document-snaptoshapes-property-word.md)| **True** if Microsoft Word automatically aligns AutoShapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes or East Asian characters in the specified document. Read/write **Boolean** .|
|[SpellingChecked](document-spellingchecked-property-word.md)| **True** if spelling has been checked throughout the specified range or document. **False** if all or some of the range or document has not been checked for spelling. Read/write **Boolean** .|
|[SpellingErrors](document-spellingerrors-property-word.md)|Returns a  **[ProofreadingErrors](proofreadingerrors-object-word.md)** collection that represents the words identified as spelling errors in the specified document or range. Read-only.|
|[StoryRanges](document-storyranges-property-word.md)|Returns a  **[StoryRanges](storyranges-object-word.md)** collection that represents all the stories in the specified document. Read-only.|
|[Styles](document-styles-property-word.md)|Returns a  **[Styles](styles-object-word.md)** collection for the specified document. Read-only.|
|[StyleSheets](document-stylesheets-property-word.md)|Returns a  **[StyleSheets](stylesheets-object-word.md)** collection that represents the Web style sheets attached to a document.|
|[StyleSortMethod](document-stylesortmethod-property-word.md)|Returns or sets a **[WdStyleSort](wdstylesort-enumeration-word.md)** constant that represents the sort method to use when sorting styles in the **Styles** task pane. Read/write.|
|[Subdocuments](document-subdocuments-property-word.md)|Returns a  **[Subdocuments](subdocuments-object-word.md)** collection that represents all the subdocuments in the specified document. Read-only.|
|[Sync](document-sync-property-word.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|
|[Tables](document-tables-property-word.md)|Returns a  **[Table](table-object-word.md)** collection that represents all the tables in the specified document. Read-only.|
|[TablesOfAuthorities](document-tablesofauthorities-property-word.md)|Returns a  **[TableOfAuthorities](tableofauthorities-object-word.md)** collection that represents the tables of authorities in the specified document. Read-only.|
|[TablesOfAuthoritiesCategories](document-tablesofauthoritiescategories-property-word.md)|Returns a  **[TablesOfAuthoritiesCategories](tablesofauthoritiescategories-object-word.md)** collection that represents the available table of authorities categories for the specified document. Read-only.|
|[TablesOfContents](document-tablesofcontents-property-word.md)|Returns a  **[TablesOfContents](tablesofcontents-object-word.md)** collection that represents the tables of contents in the specified document. Read-only.|
|[TablesOfFigures](document-tablesoffigures-property-word.md)|Returns a  **[TablesOfFigures](document-tablesoffigures-property-word.md)** collection that represents the tables of figures in the specified document. Read-only.|
|[TextEncoding](document-textencoding-property-word.md)|Returns or sets the code page, or character set, that Microsoft Word uses for a document saved as an encoded text file. Read/write  **MsoEncoding** .|
|[TextLineEnding](document-textlineending-property-word.md)|Returns or sets a  **WdLineEndingType** constant indicating how Microsoft Word marks the line and paragraph breaks in documents saved as text files. Read/write.|
|[TrackFormatting](document-trackformatting-property-word.md)|Returns or sets a ** Boolean** that represents whether to track formatting changes when change tracking is turned on. Read/write.|
|[TrackMoves](document-trackmoves-property-word.md)|Returns or sets a ** Boolean** that represents whether to mark moved text when Track Changes is turned on. Read/write.|
|[TrackRevisions](document-trackrevisions-property-word.md)| **True** if changes are tracked in the specified document. Read/write **Boolean** .|
|[Type](document-type-property-word.md)|Returns the document type (template or document). Read-only  **[WdDocumentType](wddocumenttype-enumeration-word.md)** .|
|[UpdateStylesOnOpen](document-updatestylesonopen-property-word.md)| **True** if the styles in the specified document are updated to match the styles in the attached template each time the document is opened. Read/write **Boolean** .|
|[UseMathDefaults](document-usemathdefaults-property-word.md)|Returns or sets a  **Boolean** that represents whether to use the default math settings when creating new equations. Read/write.|
|[UserControl](document-usercontrol-property-word.md)| **True** if the document was created or opened by the user. Read/write **Boolean** .|
|[Variables](document-variables-property-word.md)|Returns a  **[Variables](variables-object-word.md)** collection that represents the variables stored in the specified document. Read-only.|
|[VBASigned](document-vbasigned-property-word.md)| **True** if the Microsoft Visual Basic for Applications (VBA) project for the specified document has been digitally signed. Read-only **Boolean** .|
|[VBProject](document-vbproject-property-word.md)|Returns the  **VBProject** object for the specified template or document.|
|[WebOptions](document-weboptions-property-word.md)|Returns the  **[WebOptions](weboptions-object-word.md)** object, which contains document-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page. Read-only.|
|[Windows](document-windows-property-word.md)|Returns a  **[Windows](windows-object-word.md)** collection that represents all windows for the specified document. Read-only.|
|[WordOpenXML](document-wordopenxml-property-word.md)|Returns a  **String** that represents the flat XML format for the Word Open XML contents of the document. Read-only.|
|[Words](document-words-property-word.md)|Returns a  **[Words](words-object-word.md)** collection that represents all the words in a document. Read-only.|
|[WritePassword](document-writepassword-property-word.md)|Sets a password for saving changes to the specified document. Write-only  **String** .|
|[WriteReserved](document-writereserved-property-word.md)| **True** if the specified document is protected with a write password. Read-only **Boolean** .|
|[XMLSaveThroughXSLT](document-xmlsavethroughxslt-property-word.md)|Sets or returns a  **String** that specifies the path and file name for the Extensible Stylesheet Language Transformation (XSLT) to apply when a user saves a document.|
|[XMLSchemaReferences](document-xmlschemareferences-property-word.md)|Returns an XMLSchemaReferences collection that represents the schemas attached to a document.|
|[XMLShowAdvancedErrors](document-xmlshowadvancederrors-property-word.md)|Returns or sets a  **Boolean** that represents whether error message text is generated from the built-in Microsoft Word error messages or from the Microsoft XML Core Services (MSXML) 5.0 component included with Office.|
|[XMLUseXSLTWhenSaving](document-xmlusexsltwhensaving-property-word.md)|Returns a  **Boolean** that represents whether to save a document through an Extensible Stylesheet Language Transformation (XSLT). **True** saves a document through an XSLT.|

