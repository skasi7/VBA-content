---
title: Application Object (PowerPoint)
keywords: vbapp10.chm504000
f1_keywords:
- vbapp10.chm504000
ms.prod: POWERPOINT
ms.assetid: 978c2b99-4271-b953-4283-73b5f3d96f41
---


# Application Object (PowerPoint)

Represents the entire Microsoft PowerPoint application. 


## Remarks

The  **Application** object contains:


- Application-wide settings and options (the name of the active printer, for example).
    
- Properties that return top-level objects, such as  **ActivePresentation**, and **Windows**.
    


When you are writing code that will run from PowerPoint, you can use the following properties of the  **Application** object without the object qualifier: **ActivePresentation**, **ActiveWindow**, **AddIns**, **Presentations**, **SlideShowWindows**, **Windows**.

For example, instead of writing  `Application.ActiveWindow.Height = 200`, you can write  `ActiveWindow.Height = 200`.


## Example

Use the  **Application** property to return the **Application** object. The following example returns the path to the program file.


```
Dim MyPath As String

MyPath = Application.Path
```

The following example creates a PowerPoint  **Application** object in another application, starts PowerPoint (if it is not already running), and opens an existing presentation named "Ex_a2a.ppt".




```
Set ppt = New Powerpoint.Application

ppt.Visible = True

ppt.Presentations.Open "c:\My Documents\ex_a2a.ppt"
```


## Events



|**Name**|
|:-----|
|[AfterDragDropOnSlide](http://msdn.microsoft.com/library/application-afterdragdroponslide-event-powerpoint%28Office.15%29.aspx)|
|[AfterNewPresentation](http://msdn.microsoft.com/library/application-afternewpresentation-event-powerpoint%28Office.15%29.aspx)|
|[AfterPresentationOpen](http://msdn.microsoft.com/library/application-afterpresentationopen-event-powerpoint%28Office.15%29.aspx)|
|[AfterShapeSizeChange](http://msdn.microsoft.com/library/application-aftershapesizechange-event-powerpoint%28Office.15%29.aspx)|
|[ColorSchemeChanged](http://msdn.microsoft.com/library/application-colorschemechanged-event-powerpoint%28Office.15%29.aspx)|
|[NewPresentation](http://msdn.microsoft.com/library/application-newpresentation-event-powerpoint%28Office.15%29.aspx)|
|[PresentationBeforeClose](http://msdn.microsoft.com/library/application-presentationbeforeclose-event-powerpoint%28Office.15%29.aspx)|
|[PresentationBeforeSave](http://msdn.microsoft.com/library/application-presentationbeforesave-event-powerpoint%28Office.15%29.aspx)|
|[PresentationClose](http://msdn.microsoft.com/library/application-presentationclose-event-powerpoint%28Office.15%29.aspx)|
|[PresentationCloseFinal](http://msdn.microsoft.com/library/application-presentationclosefinal-event-powerpoint%28Office.15%29.aspx)|
|[PresentationNewSlide](http://msdn.microsoft.com/library/application-presentationnewslide-event-powerpoint%28Office.15%29.aspx)|
|[PresentationOpen](http://msdn.microsoft.com/library/application-presentationopen-event-powerpoint%28Office.15%29.aspx)|
|[PresentationPrint](http://msdn.microsoft.com/library/application-presentationprint-event-powerpoint%28Office.15%29.aspx)|
|[PresentationSave](http://msdn.microsoft.com/library/application-presentationsave-event-powerpoint%28Office.15%29.aspx)|
|[PresentationSync](http://msdn.microsoft.com/library/application-presentationsync-event-powerpoint%28Office.15%29.aspx)|
|[ProtectedViewWindowActivate](http://msdn.microsoft.com/library/application-protectedviewwindowactivate-event-powerpoint%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeClose](http://msdn.microsoft.com/library/application-protectedviewwindowbeforeclose-event-powerpoint%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeEdit](http://msdn.microsoft.com/library/application-protectedviewwindowbeforeedit-event-powerpoint%28Office.15%29.aspx)|
|[ProtectedViewWindowDeactivate](http://msdn.microsoft.com/library/application-protectedviewwindowdeactivate-event-powerpoint%28Office.15%29.aspx)|
|[ProtectedViewWindowOpen](http://msdn.microsoft.com/library/application-protectedviewwindowopen-event-powerpoint%28Office.15%29.aspx)|
|[SlideSelectionChanged](http://msdn.microsoft.com/library/application-slideselectionchanged-event-powerpoint%28Office.15%29.aspx)|
|[SlideShowBegin](http://msdn.microsoft.com/library/application-slideshowbegin-event-powerpoint%28Office.15%29.aspx)|
|[SlideShowEnd](http://msdn.microsoft.com/library/application-slideshowend-event-powerpoint%28Office.15%29.aspx)|
|[SlideShowNextBuild](http://msdn.microsoft.com/library/application-slideshownextbuild-event-powerpoint%28Office.15%29.aspx)|
|[SlideShowNextClick](http://msdn.microsoft.com/library/application-slideshownextclick-event-powerpoint%28Office.15%29.aspx)|
|[SlideShowNextSlide](http://msdn.microsoft.com/library/application-slideshownextslide-event-powerpoint%28Office.15%29.aspx)|
|[SlideShowOnNext](http://msdn.microsoft.com/library/application-slideshowonnext-event-powerpoint%28Office.15%29.aspx)|
|[SlideShowOnPrevious](http://msdn.microsoft.com/library/application-slideshowonprevious-event-powerpoint%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/application-windowactivate-event-powerpoint%28Office.15%29.aspx)|
|[WindowBeforeDoubleClick](http://msdn.microsoft.com/library/application-windowbeforedoubleclick-event-powerpoint%28Office.15%29.aspx)|
|[WindowBeforeRightClick](http://msdn.microsoft.com/library/application-windowbeforerightclick-event-powerpoint%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/application-windowdeactivate-event-powerpoint%28Office.15%29.aspx)|
|[WindowSelectionChange](http://msdn.microsoft.com/library/application-windowselectionchange-event-powerpoint%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/application-activate-method-powerpoint%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/application-help-method-powerpoint%28Office.15%29.aspx)|
|[OpenThemeFile](http://msdn.microsoft.com/library/application-openthemefile-method-powerpoint%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application-quit-method-powerpoint%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/application-run-method-powerpoint%28Office.15%29.aspx)|
|[StartNewUndoEntry](http://msdn.microsoft.com/library/application-startnewundoentry-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Active](http://msdn.microsoft.com/library/application-active-property-powerpoint%28Office.15%29.aspx)|
|[ActiveEncryptionSession](http://msdn.microsoft.com/library/application-activeencryptionsession-property-powerpoint%28Office.15%29.aspx)|
|[ActivePresentation](http://msdn.microsoft.com/library/application-activepresentation-property-powerpoint%28Office.15%29.aspx)|
|[ActivePrinter](http://msdn.microsoft.com/library/application-activeprinter-property-powerpoint%28Office.15%29.aspx)|
|[ActiveProtectedViewWindow](http://msdn.microsoft.com/library/application-activeprotectedviewwindow-property-powerpoint%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/application-activewindow-property-powerpoint%28Office.15%29.aspx)|
|[AddIns](http://msdn.microsoft.com/library/application-addins-property-powerpoint%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/application-assistance-property-powerpoint%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/application-autocorrect-property-powerpoint%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/application-automationsecurity-property-powerpoint%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/application-build-property-powerpoint%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/application-caption-property-powerpoint%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/application-chartdatapointtrack-property-powerpoint%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/application-comaddins-property-powerpoint%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/application-commandbars-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/application-creator-property-powerpoint%28Office.15%29.aspx)|
|[DisplayAlerts](http://msdn.microsoft.com/library/application-displayalerts-property-powerpoint%28Office.15%29.aspx)|
|[DisplayDocumentInformationPanel](http://msdn.microsoft.com/library/application-displaydocumentinformationpanel-property-powerpoint%28Office.15%29.aspx)|
|[DisplayGridLines](http://msdn.microsoft.com/library/application-displaygridlines-property-powerpoint%28Office.15%29.aspx)|
|[DisplayGuides](http://msdn.microsoft.com/library/application-displayguides-property-powerpoint%28Office.15%29.aspx)|
|[FeatureInstall](http://msdn.microsoft.com/library/application-featureinstall-property-powerpoint%28Office.15%29.aspx)|
|[FileConverters](http://msdn.microsoft.com/library/application-fileconverters-property-powerpoint%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/application-filedialog-property-powerpoint%28Office.15%29.aspx)|
|[FileValidation](http://msdn.microsoft.com/library/application-filevalidation-property-powerpoint%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/application-height-property-powerpoint%28Office.15%29.aspx)|
|[IsSandboxed](http://msdn.microsoft.com/library/application-issandboxed-property-powerpoint%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/application-languagesettings-property-powerpoint%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/application-left-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/application-name-property-powerpoint%28Office.15%29.aspx)|
|[NewPresentation](http://msdn.microsoft.com/library/application-newpresentation-property-powerpoint%28Office.15%29.aspx)|
|[OperatingSystem](http://msdn.microsoft.com/library/application-operatingsystem-property-powerpoint%28Office.15%29.aspx)|
|[Options](http://msdn.microsoft.com/library/application-options-property-powerpoint%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/application-path-property-powerpoint%28Office.15%29.aspx)|
|[Presentations](http://msdn.microsoft.com/library/application-presentations-property-powerpoint%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/application-productcode-property-powerpoint%28Office.15%29.aspx)|
|[ProtectedViewWindows](http://msdn.microsoft.com/library/application-protectedviewwindows-property-powerpoint%28Office.15%29.aspx)|
|[ShowStartupDialog](http://msdn.microsoft.com/library/application-showstartupdialog-property-powerpoint%28Office.15%29.aspx)|
|[ShowWindowsInTaskbar](http://msdn.microsoft.com/library/application-showwindowsintaskbar-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowWindows](http://msdn.microsoft.com/library/application-slideshowwindows-property-powerpoint%28Office.15%29.aspx)|
|[SmartArtColors](http://msdn.microsoft.com/library/application-smartartcolors-property-powerpoint%28Office.15%29.aspx)|
|[SmartArtLayouts](http://msdn.microsoft.com/library/application-smartartlayouts-property-powerpoint%28Office.15%29.aspx)|
|[SmartArtQuickStyles](http://msdn.microsoft.com/library/application-smartartquickstyles-property-powerpoint%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/application-top-property-powerpoint%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/application-vbe-property-powerpoint%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/application-version-property-powerpoint%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/application-visible-property-powerpoint%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/application-width-property-powerpoint%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/application-windows-property-powerpoint%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/application-windowstate-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
