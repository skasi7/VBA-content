---
title: Global Members (Word)
ms.prod: WORD
ms.assetid: 35050f7b-bc46-4795-ec17-f68e263c8af0
---


# Global Members (Word)
Contains top-level properties and methods that do not need to be preceded by the  **Application** property.

Contains top-level properties and methods that do not need to be preceded by the  **Application** property.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[BuildKeyCode](global-buildkeycode-method-word.md)|Returns a unique number for the specified key combination.|
|[CentimetersToPoints](global-centimeterstopoints-method-word.md)|Converts a measurement from centimeters to points (1 cm = 28.35 points). Returns the converted measurement as a  **Single** .|
|[ChangeFileOpenDirectory](global-changefileopendirectory-method-word.md)|Sets the folder in which Word searches for documents. .|
|[CheckSpelling](global-checkspelling-method-word.md)|Checks a string for spelling errors. Returns a  **Boolean** to indicate whether the string contains spelling errors. **True** if the string has no spelling errors.|
|[CleanString](global-cleanstring-method-word.md)|Removes nonprinting characters (character codes 1 ? 29) and special Word characters from the specified string or changes them to spaces (character code 32), as described in the "Remarks" section. Returns the result as a  **String** .|
|[DDEExecute](global-ddeexecute-method-word.md)|Sends a command or series of commands to an application through the specified dynamic data exchange (DDE) channel.|
|[DDEInitiate](global-ddeinitiate-method-word.md)|Opens a dynamic data exchange (DDE) channel to another application, and returns the channel number.|
|[DDEPoke](global-ddepoke-method-word.md)|Uses an open dynamic data exchange (DDE) channel to send data to an application.|
|[DDERequest](global-dderequest-method-word.md)|Uses an open dynamic data exchange (DDE) channel to request information from the receiving application, and returns the information as a string.|
|[DDETerminate](global-ddeterminate-method-word.md)|Closes the specified dynamic data exchange (DDE) channel to another application.|
|[DDETerminateAll](global-ddeterminateall-method-word.md)|Closes all dynamic data exchange (DDE) channels opened by Microsoft Word. .|
|[GetSpellingSuggestions](global-getspellingsuggestions-method-word.md)|Returns a  **[SpellingSuggestions](spellingsuggestions-object-word.md)** collection that represents the words suggested as spelling replacements for a given word.|
|[Help](global-help-method-word.md)|Displays on-line Help information.|
|[InchesToPoints](global-inchestopoints-method-word.md)|Converts a measurement from inches to points (1 inch = 72 points). Returns the converted measurement as a  **Single** .|
|[KeyString](global-keystring-method-word.md)|Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A).|
|[LinesToPoints](global-linestopoints-method-word.md)|Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a  **Single** .|
|[MillimetersToPoints](global-millimeterstopoints-method-word.md)|Converts a measurement from millimeters to points (1 mm = 2.85 points). Returns the converted measurement as a  **Single** .|
|[NewWindow](global-newwindow-method-word.md)|Opens a new window with the same document as the specified window. Returns a  **Window** object.|
|[PicasToPoints](global-picastopoints-method-word.md)|Converts a measurement from picas to points (1 pica = 12 points). Returns the converted measurement as a  **Single** .|
|[PixelsToPoints](global-pixelstopoints-method-word.md)|Converts a measurement from pixels to points. Returns the converted measurement as a  **Single** .|
|[PointsToCentimeters](global-pointstocentimeters-method-word.md)|Converts a measurement from points to centimeters (1 centimeter = 28.35 points). Returns the converted measurement as a  **Single** .|
|[PointsToInches](global-pointstoinches-method-word.md)|Converts a measurement from points to inches (1 inch = 72 points). Returns the converted measurement as a  **Single** .|
|[PointsToLines](global-pointstolines-method-word.md)|Converts a measurement from points to lines (1 line = 12 points). Returns the converted measurement as a  **Single** .|
|[PointsToMillimeters](global-pointstomillimeters-method-word.md)|Converts a measurement from points to millimeters (1 millimeter = 2.835 points). Returns the converted measurement as a  **Single** .|
|[PointsToPicas](global-pointstopicas-method-word.md)|Converts a measurement from points to picas (1 pica = 12 points). Returns the converted measurement as a  **Single** .|
|[PointsToPixels](global-pointstopixels-method-word.md)|Converts a measurement from points to pixels. Returns the converted measurement as a  **Single** .|
|[Repeat](global-repeat-method-word.md)|Repeats the most recent editing action one or more times. Returns  **True** if the commands were repeated successfully.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActiveDocument](global-activedocument-property-word.md)|Returns a  **[Document](document-object-word.md)** object that represents the active document (the document with the focus). Read-only.|
|[ActivePrinter](global-activeprinter-property-word.md)|Returns or sets the name of the active printer. Read/write  **String** .|
|[ActiveProtectedViewWindow](global-activeprotectedviewwindow-property-word.md)|Returns a [ProtectedViewWindow](protectedviewwindow-object-word.md) object that represents the active protected view window (the protected view window with the focus). Read-only.|
|[ActiveWindow](global-activewindow-property-word.md)|Returns a  **[Window](window-object-word.md)** object that represents the active window (the window with the focus). Read-only.|
|[AddIns](global-addins-property-word.md)|Returns an  **[AddIns](addins-object-word.md)** collection that represents all available add-ins, regardless of whether they're currently loaded. Read-only.|
|[Application](global-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoCaptions](global-autocaptions-property-word.md)|Returns an  **[AutoCaptions](autocaptions-object-word.md)** collection that represents the captions that are automatically added when items such as tables and pictures are inserted into a document. Read-only.|
|[AutoCorrect](global-autocorrect-property-word.md)|Returns an  **[AutoCorrect](autocorrect-object-word.md)** object that contains the current AutoCorrect options, entries, and exceptions. Read-only.|
|[AutoCorrectEmail](global-autocorrectemail-property-word.md)|Returns an  **[AutoCorrect](autocorrect-object-word.md)** object that represents automatic corrections made to e-mail messages.|
|[CaptionLabels](global-captionlabels-property-word.md)|Returns a  **[CaptionLabels](captionlabels-object-word.md)** collection that represents all the available caption labels. Read-only.|
|[CommandBars](global-commandbars-property-word.md)|Returns a  **CommandBars** collection that represents the menu bar and all the toolbars in Microsoft Word.|
|[Creator](global-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[CustomDictionaries](global-customdictionaries-property-word.md)|Returns a  **[Dictionaries](dictionaries-object-word.md)** object that represents the collection of active custom dictionaries. Read-only.|
|[CustomizationContext](global-customizationcontext-property-word.md)|Returns or sets a  **Template** or **[Document](document-object-word.md)** object that represents the template or document in which changes to menu bars, toolbars, and key bindings are stored. Read/write. .|
|[Dialogs](global-dialogs-property-word.md)|Returns a  **[Dialogs](dialogs-object-word.md)** collection that represents all the built-in dialog boxes in Word. Read-only.|
|[Documents](global-documents-property-word.md)|Returns a  **[Documents](documents-object-word.md)** collection that represents all the open documents. Read-only.|
|[FileConverters](global-fileconverters-property-word.md)|Returns a  **[FileConverters](fileconverters-object-word.md)** collection that represents all the file converters available to Microsoft Word. Read-only.|
|[FindKey](global-findkey-property-word.md)|Returns a  **[KeyBinding](keybinding-object-word.md)** object that represents the specified key combination. Read-only.|
|[FontNames](global-fontnames-property-word.md)|Returns a  **[FontNames](fontnames-object-word.md)** object that includes the names of all the available fonts. Read-only.|
|[HangulHanjaDictionaries](global-hangulhanjadictionaries-property-word.md)|Returns a  **[HangulHanjaConversionDictionaries](hangulhanjaconversiondictionaries-object-word.md)** collection that represents all the active custom conversion dictionaries.|
|[IsObjectValid](global-isobjectvalid-property-word.md)| **True** if the specified variable that references an object is valid. Read-only **Boolean** .|
|[IsSandboxed](global-issandboxed-property-word.md)| **True** if the application window is a protected view window. Read-only.|
|[KeyBindings](global-keybindings-property-word.md)|Returns a  **KeyBindings** collection that represents customized key assignments, which include a key code, a key category, and a command.|
|[KeysBoundTo](global-keysboundto-property-word.md)|Returns a  **KeysBoundTo** object that represents all the key combinations assigned to the specified item.|
|[LandscapeFontNames](global-landscapefontnames-property-word.md)|Returns a  **FontNames** object that includes the names of all the available landscape fonts.|
|[Languages](global-languages-property-word.md)|Returns a  **Languages** collection that represents the proofing languages listed in the **Language** dialog box.|
|[LanguageSettings](global-languagesettings-property-word.md)|Returns a  **LanguageSettings** object, which contains information about the language settings in Microsoft Word.|
|[ListGalleries](global-listgalleries-property-word.md)|Returns a  **ListGalleries** collection that represents the three list template galleries ( **Bulleted**,  **Numbered**, and  **Outline Numbered**).|
|[MacroContainer](global-macrocontainer-property-word.md)|Returns a  **Template** or **Document** object that represents the template or document in which the module that contains the running procedure is stored.|
|[Name](global-name-property-word.md)|Returns name of the specified object. Read-only  **String** .|
|[NormalTemplate](global-normaltemplate-property-word.md)|Returns a  **Template** object that represents the Normal template.|
|[Options](global-options-property-word.md)|Returns an  **Options** object that represents application settings in Microsoft Word.|
|[Parent](global-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Global** object.|
|[PortraitFontNames](global-portraitfontnames-property-word.md)|Returns a  **FontNames** object that includes the names of all the available portrait fonts.|
|[PrintPreview](global-printpreview-property-word.md)| **True** if print preview is the current view. Read/write **Boolean** .|
|[ProtectedViewWindows](global-protectedviewwindows-property-word.md)|Returns a [ProtectedViewWindows](protectedviewwindows-object-word.md) object that represents the open protected view windows. Read-only.|
|[RecentFiles](global-recentfiles-property-word.md)|Returns a  **RecentFiles** collection that represents the most recently accessed files.|
|[Selection](global-selection-property-word.md)|Returns a  **Selection** object that represents a selected range or the insertion point. Read-only.|
|[ShowVisualBasicEditor](global-showvisualbasiceditor-property-word.md)| **True** if the Visual Basic Editor window is visible. Read/write **Boolean** .|
|[StatusBar](global-statusbar-property-word.md)|This property is no longer supported in Microsoft Word Visual Basic for Applications.|
|[SynonymInfo](global-synonyminfo-property-word.md)|Returns a  **SynonymInfo** object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the specified word or phrase.|
|[System](global-system-property-word.md)|Returns a  **System** object, which can be used to return system-related information and perform system-related tasks.|
|[Tasks](global-tasks-property-word.md)|Returns a  **Tasks** collection that represents all the applications that are running.|
|[Templates](global-templates-property-word.md)|Returns a  **Templates** collection that represents all the available templates—global templates and those attached to open documents.|
|[VBE](global-vbe-property-word.md)|Returns a  **VBE** object that represents the Visual Basic Editor.|
|[Windows](global-windows-property-word.md)|Returns a  **Windows** collection that represents all open document windows. Read-only.|
|[WordBasic](global-wordbasic-property-word.md)|Returns an Automation object (Word.Basic) that includes methods for all the WordBasic statements and functions available in Word version 6.0 and Word for Windows 95. Read-only.|

