---
title: Application Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 7a9042da-ef77-ebba-c872-f736bf486674
---


# Application Members (PowerPoint)
Represents the entire Microsoft PowerPoint application. 

Represents the entire Microsoft PowerPoint application. 


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterDragDropOnSlide](application-afterdragdroponslide-event-powerpoint.md)|Occurs after a shape has been dropped onto a slide in an open presentation.|
|[AfterNewPresentation](application-afternewpresentation-event-powerpoint.md)|Occurs after a presentation is created.|
|[AfterPresentationOpen](application-afterpresentationopen-event-powerpoint.md)|Occurs after an existing presentation is opened.|
|[AfterShapeSizeChange](application-aftershapesizechange-event-powerpoint.md)|Occurs after an object (shape, picture, text box, chart, SmartArt as examples) has been resized on the slide.|
|[ColorSchemeChanged](application-colorschemechanged-event-powerpoint.md)|Occurs after a color scheme is changed.|
|[NewPresentation](application-newpresentation-event-powerpoint.md)|Occurs after a presentation is created, as it is added to the  **[Presentations](presentations-object-powerpoint.md)** collection.|
|[PresentationBeforeClose](application-presentationbeforeclose-event-powerpoint.md)|Represents a  **Presentation** object before it closes.|
|[PresentationBeforeSave](application-presentationbeforesave-event-powerpoint.md)|Occurs before a presentation is saved.|
|[PresentationClose](application-presentationclose-event-powerpoint.md)|Occurs immediately before any open presentation closes, as it is removed from the  **[Presentations](presentations-object-powerpoint.md)** collection.|
|[PresentationCloseFinal](application-presentationclosefinal-event-powerpoint.md)|Represents closing the final  **Presentation** object.|
|[PresentationNewSlide](application-presentationnewslide-event-powerpoint.md)|Occurs when a new slide is created in any open presentation, as the slide is added to the  **[Slides](slides-object-powerpoint.md)** collection.|
|[PresentationOpen](application-presentationopen-event-powerpoint.md)|Occurs after an existing presentation is opened, as it is added to the  **[Presentations](presentations-object-powerpoint.md)** collection.|
|[PresentationPrint](application-presentationprint-event-powerpoint.md)|Occurs before a presentation is printed.|
|[PresentationSave](application-presentationsave-event-powerpoint.md)|Occurs before any open presentation is saved.|
|[PresentationSync](application-presentationsync-event-powerpoint.md)|Occurs when the local copy of a presentation that is part of a Document Workspace is synchronized with the copy on the server. Provides important status information regarding the success or failure of the synchronization of the presentation.|
|[ProtectedViewWindowActivate](application-protectedviewwindowactivate-event-powerpoint.md)|Occurs when any protected view window is activated.|
|[ProtectedViewWindowBeforeClose](application-protectedviewwindowbeforeclose-event-powerpoint.md)|Occurs immediately before a protected view window or a document in a protected view window closes.|
|[ProtectedViewWindowBeforeEdit](application-protectedviewwindowbeforeedit-event-powerpoint.md)|Occurs immediately before editing is enabled on the document in the specified protected view window.|
|[ProtectedViewWindowDeactivate](application-protectedviewwindowdeactivate-event-powerpoint.md)|Occurs when a protected view window is deactivated.|
|[ProtectedViewWindowOpen](application-protectedviewwindowopen-event-powerpoint.md)|Occurs when a protected view window is opened.|
|[SlideSelectionChanged](application-slideselectionchanged-event-powerpoint.md)|Occurs at different times depending on the current view.|
|[SlideShowBegin](application-slideshowbegin-event-powerpoint.md)|Occurs when you start a slide show.|
|[SlideShowEnd](application-slideshowend-event-powerpoint.md)|Occurs after a slide show ends, immediately after the last  **[SlideShowNextSlide](application-slideshownextslide-event-powerpoint.md)** event occurs.|
|[SlideShowNextBuild](application-slideshownextbuild-event-powerpoint.md)|Occurs upon mouse-click or timing animation, but before the animated object becomes visible. .|
|[SlideShowNextClick](application-slideshownextclick-event-powerpoint.md)|Occurs on the next click of the slide.|
|[SlideShowNextSlide](application-slideshownextslide-event-powerpoint.md)|Occurs immediately before the transition to the next slide. For the first slide, occurs immediately after the  **[SlideShowBegin](application-slideshowbegin-event-powerpoint.md)** event.|
|[SlideShowOnNext](application-slideshowonnext-event-powerpoint.md)|Occurs when the user clicks  **Next** to move within the current slide.|
|[SlideShowOnPrevious](application-slideshowonprevious-event-powerpoint.md)|Occurs when the user clicks  **Previous** to move within the current slide.|
|[WindowActivate](application-windowactivate-event-powerpoint.md)|Occurs when the application window or any document window is activated.|
|[WindowBeforeDoubleClick](application-windowbeforedoubleclick-event-powerpoint.md)|Occurs when you double-click the items in the views listed in the following table.|
|[WindowBeforeRightClick](application-windowbeforerightclick-event-powerpoint.md)|Occurs when you right-click a shape, a slide, a notes page, or some text. This event is triggered by the  **MouseUp** event.|
|[WindowDeactivate](application-windowdeactivate-event-powerpoint.md)|Occurs when the application window or any document window is deactivated.|
|[WindowSelectionChange](application-windowselectionchange-event-powerpoint.md)|Occurs when the selection of text, a shape, or a slide in the active document window changes, whether in the user interface or in code.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](application-activate-method-powerpoint.md)|Activates the specified object.|
|[Help](application-help-method-powerpoint.md)|Displays a Help topic.|
|[OpenThemeFile](application-openthemefile-method-powerpoint.md)|Opens the specified theme file (*thmx).|
|[Quit](application-quit-method-powerpoint.md)|Quits Microsoft PowerPoint. This is equivalent to clicking the  **Office** button and then clicking **Exit PowerPoint**.|
|[Run](application-run-method-powerpoint.md)|Runs a Visual Basic procedure.|
|[StartNewUndoEntry](application-startnewundoentry-method-powerpoint.md)|Starts a new undo entry in the  **Application** object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Active](application-active-property-powerpoint.md)|Returns whether the specified pane or window is active. Read-only.|
|[ActiveEncryptionSession](application-activeencryptionsession-property-powerpoint.md)|Represents the encryption session associated with the active presentation. Read-only.|
|[ActivePresentation](application-activepresentation-property-powerpoint.md)|Returns a  **[Presentation](presentation-object-powerpoint.md)** object that represents the presentation open in the active window. Read-only.|
|[ActivePrinter](application-activeprinter-property-powerpoint.md)|Returns the name of the active printer. Read-only.|
|[ActiveProtectedViewWindow](application-activeprotectedviewwindow-property-powerpoint.md)|Returns a  **[ProtectedViewWindow](protectedviewwindow-object-powerpoint.md)** object that represents the active **Protected View** window (the window on top). Read-only.|
|[ActiveWindow](application-activewindow-property-powerpoint.md)|Returns a  **[DocumentWindow](documentwindow-object-powerpoint.md)** object that represents the active document window. Read-only.|
|[AddIns](application-addins-property-powerpoint.md)|Returns the program-specific  **AddIns** collection that represents all the add-ins listed in the **Add-Ins** dialog box (click the **Office** button, click **PowerPoint Options**, click  **Add-Ins**, click  **PowerPoint Add-Ins** on the **Manage** list). Read-only.|
|[Assistance](application-assistance-property-powerpoint.md)|Gets a reference to the Microsoft Office  **[IAssistance](iassistance-object-office.md)** object, which provides a means for developers to create a customized help experience for users within Microsoft Office. Read-only.|
|[AutoCorrect](application-autocorrect-property-powerpoint.md)|Returns an  **[AutoCorrect](autocorrect-object-powerpoint.md)** object that represents the AutoCorrect functionality in Microsoft PowerPoint.|
|[AutomationSecurity](application-automationsecurity-property-powerpoint.md)|Represents the security mode that Microsoft PowerPoint uses when it opens files programmatically. Read/write.|
|[Build](application-build-property-powerpoint.md)|Returns the build number for the current instance of Microsoft PowerPoint. Read-only.|
|[Caption](application-caption-property-powerpoint.md)|Returns the text that appears in the title bar of the application window. Read/write.|
|[ChartDataPointTrack](application-chartdatapointtrack-property-powerpoint.md)|Returns or sets a  **Boolean** that specifies whether charts use cell-reference data-point tracking. Read-write.|
|[COMAddIns](application-comaddins-property-powerpoint.md)|Returns a reference to the Component Object Model (COM) add-ins currently loaded in Microsoft PowerPoint. These add-ins are listed on the  **Add-Ins** tab in the **PowerPoint Options** dialog box. Read-only.|
|[CommandBars](application-commandbars-property-powerpoint.md)|Returns a  **CommandBars** collection that represents all the command bars in Microsoft PowerPoint. Read-only.|
|[Creator](application-creator-property-powerpoint.md)|Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.|
|[DisplayAlerts](application-displayalerts-property-powerpoint.md)|Sets or returns whether Microsoft PowerPoint displays alerts while running a macro. Read/write.|
|[DisplayDocumentInformationPanel](application-displaydocumentinformationpanel-property-powerpoint.md)|Returns or sets whether the  **Document Properties** panel is displayed in the Microsoft PowerPoint user interface. Read/write.|
|[DisplayGridLines](application-displaygridlines-property-powerpoint.md)|Determines whether to display gridlines in Microsoft PowerPoint. Read/write.|
|[DisplayGuides](application-displayguides-property-powerpoint.md)|Gets or sets whether drawing guides are displayed in the application. |
|[FeatureInstall](application-featureinstall-property-powerpoint.md)|Returns or sets how Microsoft PowerPoint handles calls to methods and properties that require features not yet installed. Read/write.|
|[FileConverters](application-fileconverters-property-powerpoint.md)|Returns information about installed file converters. Returns  **null** if there are no converters installed. Read-only **Variant**.|
|[FileDialog](application-filedialog-property-powerpoint.md)|Returns a  **FileDialog** object that represents a single instance of a file dialog box. Read-only.|
|[FileValidation](application-filevalidation-property-powerpoint.md)|Returns or sets a value that indicates how PowerPoint will validate files before opening them. Read/write|
|[Height](application-height-property-powerpoint.md)|Returns or sets the height of the specified object, in points. Read/write.|
|[IsSandboxed](application-issandboxed-property-powerpoint.md)|Returns  **True** if the specified presentation is open in a **Protected View** window. Read-only.|
|[LanguageSettings](application-languagesettings-property-powerpoint.md)|Returns a  **LanguageSettings** object that contains information about the language settings in Microsoft PowerPoint. Read-only.|
|[Left](application-left-property-powerpoint.md)|Returns or sets a  **Single** that represents the distance in points from the left edge of the document, application, and slide show windows to the left edge of the application window's client area. Setting this property to a very large positive or negative value may position the window completely off the desktop. Read/write.|
|[Name](application-name-property-powerpoint.md)|Returns the string "Microsoft PowerPoint." Read-only.|
|[NewPresentation](application-newpresentation-property-powerpoint.md)|Returns a  **NewFile** object that represents a presentation listed on the **New Presentation** task pane. Read-only.|
|[OperatingSystem](application-operatingsystem-property-powerpoint.md)|Returns the name of the operating system. Read-only.|
|[Options](application-options-property-powerpoint.md)|Returns an  **[Options](options-object-powerpoint.md)** object that represents application options in Microsoft PowerPoint.|
|[Path](application-path-property-powerpoint.md)|Returns a  **String** that represents the path to the specified **[Application](application-object-powerpoint.md)** object. Read-only.|
|[Presentations](application-presentations-property-powerpoint.md)|Returns a  **[Presentations](presentations-object-powerpoint.md)** collection that represents all open presentations. Read-only.|
|[ProductCode](application-productcode-property-powerpoint.md)|Returns the Microsoft PowerPoint globally unique identifier (GUID). Read-only.|
|[ProtectedViewWindows](application-protectedviewwindows-property-powerpoint.md)|Returns a  **[ProtectedViewWindows](protectedviewwindows-object-powerpoint.md)** collection that represents all the **Protected View** windows that are open in the application. Read-only|
|[ShowStartupDialog](application-showstartupdialog-property-powerpoint.md)|Determines whether to display the  **New Presentation** task pane when Microsoft PowerPoint is started. Read/write.|
|[ShowWindowsInTaskbar](application-showwindowsintaskbar-property-powerpoint.md)|Determines whether there is a separate Windows taskbar button for each open presentation. Read/write.|
|[SlideShowWindows](application-slideshowwindows-property-powerpoint.md)|Returns a  **[SlideShowWindows](slideshowwindows-object-powerpoint.md)** collection that represents all open slide show windows. Read-only.|
|[SmartArtColors](application-smartartcolors-property-powerpoint.md)|Returns the SmartArt colors of the current  **Application** object. Read-only.|
|[SmartArtLayouts](application-smartartlayouts-property-powerpoint.md)|Returns the SmartArt layout of the current  **Application** object. Read-only.|
|[SmartArtQuickStyles](application-smartartquickstyles-property-powerpoint.md)|Returns the quick styles of the SmartArt diagram in the current  **Application** object. Read-only.|
|[Top](application-top-property-powerpoint.md)|Returns or sets a  **Single** that represents the distance in points from the top edge of the document, application, and slide show window to the top edge of the application window's client area. Read/write.|
|[VBE](application-vbe-property-powerpoint.md)|Returns a  **VBE** object that represents the Visual Basic Editor. Read-only.|
|[Version](application-version-property-powerpoint.md)|Returns the Microsoft PowerPoint version number. Read-only.|
|[Visible](application-visible-property-powerpoint.md)|Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.|
|[Width](application-width-property-powerpoint.md)|Returns or sets the width of the specified object, in points. Read/write.|
|[Windows](application-windows-property-powerpoint.md)|Returns a  **[DocumentWindows](documentwindows-object-powerpoint.md)** collection that represents all open document windows. Read-only.|
|[WindowState](application-windowstate-property-powerpoint.md)|Returns or sets the state of the specified window. Read/write.|

