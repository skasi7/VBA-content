---
title: ApplicationSettings Members (Visio)
ms.prod: VISIO
ms.assetid: 6d3ef36b-8a8f-4ba0-ceca-cb501bb9840a
---


# ApplicationSettings Members (Visio)
Represents various application settings for Microsoft Visio.

Represents various application settings for Microsoft Visio.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[GetRasterExportResolution](applicationsettings-getrasterexportresolution-method-visio.md)|Returns the raster export resolution settings.|
|[GetRasterExportSize](applicationsettings-getrasterexportsize-method-visio.md)|Gets the raster export size.|
|[SetRasterExportResolution](applicationsettings-setrasterexportresolution-method-visio.md)|Specifies the raster export resolution settings.|
|[SetRasterExportSize](applicationsettings-setrasterexportsize-method-visio.md)|Sets the raster export size.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](applicationsettings-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[ApplyBackgroundToDocument](applicationsettings-applybackgroundtodocument-property-visio.md)|Determines whether the selected background or border is applied to all pages in the document ( **True** ) or only to the current page ( **False** ). Read/write.|
|[ApplyThemesOnShapeAdd](applicationsettings-applythemesonshapeadd-property-visio.md)|Gets or sets whether to apply themes to new shapes when they are added to the drawing page. Read/write.|
|[AsianTextUI](applicationsettings-asiantextui-property-visio.md)|Gets whether Asian text is displayed in the Microsoft Visio user interface. Read-only.|
|[BIDITextUI](applicationsettings-biditextui-property-visio.md)|Gets the current setting for display of right-to-left languages. Read-only.|
|[CenterSelectionOnZoom](applicationsettings-centerselectiononzoom-property-visio.md)|Determines whether when the user zooms in, the selection appears in the center of the window. Read/write.|
|[ComplexTextUI](applicationsettings-complextextui-property-visio.md)|Gets whether complex text is displayed in the Microsoft Visio user interface. Read-only.|
|[ConnectorSplittingEnabled](applicationsettings-connectorsplittingenabled-property-visio.md)|Determines whether connector splitting is enabled in Microsoft Visio. Read/write.|
|[DefaultSaveFormat](applicationsettings-defaultsaveformat-property-visio.md)|Determines the default format for saving Microsoft Visio files. Read/write.|
|[DeleteConnectorsEnabled](applicationsettings-deleteconnectorsenabled-property-visio.md)|Determines whether connectors are deleted when a shape to which they are connected is deleted. Read/write.|
|[DeveloperMode](applicationsettings-developermode-property-visio.md)|Determines if certain user interface functions for the development environment in Microsoft Visio are enabled. Read/write.|
|[DrawingAids](applicationsettings-drawingaids-property-visio.md)|Determines whether drawing aids are currently active in Microsoft Visio. Read/write.|
|[DrawingBackgroundColor](applicationsettings-drawingbackgroundcolor-property-visio.md)|Determines the background color of the Microsoft Visio drawing window for the current session. Read/write.|
|[DrawingBackgroundColorGradient](applicationsettings-drawingbackgroundcolorgradient-property-visio.md)|Determines the background gradient color of the Microsoft Visio drawing window for the current session. Read/write. |
|[DrawingPageColor](applicationsettings-drawingpagecolor-property-visio.md)|Determines the page color of the Microsoft Visio drawing window for the current session. Read/write. |
|[EnableAutoConnect](applicationsettings-enableautoconnect-property-visio.md)|Determines whether the  **AutoConnect** feature is enabled in the Microsoft Visio user interface (UI). Read/write.|
|[EnableFormulaAutoComplete](applicationsettings-enableformulaautocomplete-property-visio.md)|Indicates whether ShapeSheet formula AutoComplete is enabled. Read/write.|
|[EnterCommitsText](applicationsettings-entercommitstext-property-visio.md)|Returns or sets a  **Boolean** that determines whether pressing **Enter** commits shape text ( **True**) or writes a new line ( **False**, the default). Read/write.|
|[FreeformDrawingPrecision](applicationsettings-freeformdrawingprecision-property-visio.md)|Determines the margin of error allowed when the  **Freeform** tool is drawing a straight line before it switches to drawing a spline. Read/write.|
|[FreeformDrawingSmoothing](applicationsettings-freeformdrawingsmoothing-property-visio.md)|Determines how precisely mouse movements are smoothed when drawing a spline. Read/write.|
|[KanaFindAndReplace](applicationsettings-kanafindandreplace-property-visio.md)|Gets whether additional options specific to Japanese in the  **Find** and **Replace** dialog boxes are available. Read-only.|
|[KashidaTextUI](applicationsettings-kashidatextui-property-visio.md)|Gets the current setting for display of Kashida text-justification in certain cursive languages. Read-only.|
|[ObjectType](applicationsettings-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[RasterExportBackgroundColor](applicationsettings-rasterexportbackgroundcolor-property-visio.md)|Determines the background color that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.|
|[RasterExportColorFormat](applicationsettings-rasterexportcolorformat-property-visio.md)|Determines the color format that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, JPG, PNG, or TIFF file. Read/write.|
|[RasterExportColorReduction](applicationsettings-rasterexportcolorreduction-property-visio.md)|Determines the color reduction that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, GIF, PNG, or TIFF file. Read/write.|
|[RasterExportDataCompression](applicationsettings-rasterexportdatacompression-property-visio.md)|Determines the data compression algorithm that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP or TIFF file. Read/write.|
|[RasterExportDataFormat](applicationsettings-rasterexportdataformat-property-visio.md)|Determines whether the exported raster image is interlaced or non-interlaced when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a GIF or PNG file. Read/write.|
|[RasterExportFlip](applicationsettings-rasterexportflip-property-visio.md)|Determines the flip that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.|
|[RasterExportOperation](applicationsettings-rasterexportoperation-property-visio.md)|Determines the export operation that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a JPG file. Read/write.|
|[RasterExportQuality](applicationsettings-rasterexportquality-property-visio.md)|Determines the export quality that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a JPG file. Read/write.|
|[RasterExportRotation](applicationsettings-rasterexportrotation-property-visio.md)|Determines the rotation that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.|
|[RasterExportTransparencyColor](applicationsettings-rasterexporttransparencycolor-property-visio.md)|Determines the transparency color that is applied to the exported image when you call the  **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a GIF or PNG file. Read/write.|
|[RasterExportUseTransparencyColor](applicationsettings-rasterexportusetransparencycolor-property-visio.md)|Determines whether Microsoft Visio applies, to the exported image, the transparency color that is specified in the  **RasterExportTransparencyColor** property when you call the **Export** method of the **[Master](master-object-visio.md)** , **[Page](page-object-visio.md)** , **[Selection](selection-object-visio.md)** , or **[Shape](shape-object-visio.md)** object to export the specified object to a GIF or PNG file. Read/write.|
|[RecentFilesListSize](applicationsettings-recentfileslistsize-property-visio.md)|Determines the number of entries in the  **Recent Documents** list in the Microsoft Visio user interface. Read/write.|
|||
|[SATextUI](applicationsettings-satextui-property-visio.md)|Gets the current setting for display of South Asian languages. Read-only.|
|[ShowChooseDrawingTypePane](applicationsettings-showchoosedrawingtypepane-property-visio.md)|Determines if the  **New** tab appears when the user opens Microsoft Visio. Read/write.|
|[ShowFileOpenWarnings](applicationsettings-showfileopenwarnings-property-visio.md)|Determines if warning messages appear when the user attempts to open files in XML format that contain errors, such as invalid XML code. Read/write.|
|[ShowFileSaveWarnings](applicationsettings-showfilesavewarnings-property-visio.md)|Determines if warning messages appear when the user attempts to save in XML format drawings that contain errors, such as invalid XML code. Read/write.|
|[ShowMoreShapeHandlesOnHover](applicationsettings-showmoreshapehandlesonhover-property-visio.md)|Gets or sets whether to show additional shape handles when the mouse is paused over a shape. Read/write.|
|[ShowShapeSearchPane](applicationsettings-showshapesearchpane-property-visio.md)|Gets or sets whether the  **Shape Search** pane is visible in the Microsoft Visio user interface (UI). Read/write.|
|[ShowSmartTags](applicationsettings-showsmarttags-property-visio.md)|Determines whether display of smart tags in Microsoft Visio is enabled. Read/write.|
|[SnapStrengthExtensionsX](applicationsettings-snapstrengthextensionsx-property-visio.md)|Specifies the distance in pixels along the  _x_ -axis that shape extension lines pull when snapping is enabled. Read/Write.|
|[SnapStrengthExtensionsY](applicationsettings-snapstrengthextensionsy-property-visio.md)|Specifies the distance in pixels along the  _y-_ axis that shape extension lines pull when snapping is enabled. Read/write.|
|[SnapStrengthGeometryX](applicationsettings-snapstrengthgeometryx-property-visio.md)|Specifies the distance in pixels along the  _x_ -axis that shape geometry pulls when snapping is enabled. Read/write.|
|[SnapStrengthGeometryY](applicationsettings-snapstrengthgeometryy-property-visio.md)|Specifies the distance in pixels along the  _y_ -axis that shape geometry pulls when snapping is enabled. Read/write.|
|[SnapStrengthGridX](applicationsettings-snapstrengthgridx-property-visio.md)|Specifies the distance in pixels along the x-axis that gridlines pull when snapping is enabled. Read/write.|
|[SnapStrengthGridY](applicationsettings-snapstrengthgridy-property-visio.md)|Specifies the distance in pixels along the  _y_-axis that gridlines pull when snapping is enabled. Read/write.|
|[SnapStrengthGuidesX](applicationsettings-snapstrengthguidesx-property-visio.md)|Specifies the distance in pixels along the x-axis that guides pull when snapping is enabled. Read/write.|
|[SnapStrengthGuidesY](applicationsettings-snapstrengthguidesy-property-visio.md)|Specifies the distance in pixels along the y-axis that guides pull when snapping is enabled. Read/write.|
|[SnapStrengthPointsX](applicationsettings-snapstrengthpointsx-property-visio.md)|Specifies the distance in pixels along the x-axis that points pull when snapping is enabled. Read/write.|
|[SnapStrengthPointsY](applicationsettings-snapstrengthpointsy-property-visio.md)|Specifies the distance in pixels along the y-axis that points pull when snapping is enabled. Read/write.|
|[SnapStrengthRulerX](applicationsettings-snapstrengthrulerx-property-visio.md)|Specifies the distance in pixels along the x-axis that rulers pull when snapping is enabled. Read/write.|
|[SnapStrengthRulerY](applicationsettings-snapstrengthrulery-property-visio.md)|Specifies the distance in pixels along the y-axis that rulers pull when snapping is enabled. Read/write.|
|[Stat](applicationsettings-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[StencilBackgroundColor](applicationsettings-stencilbackgroundcolor-property-visio.md)|Determines the background color of the Microsoft Visio stencil window for the current session. Read/write.|
|[StencilBackgroundColorGradient](applicationsettings-stencilbackgroundcolorgradient-property-visio.md)|Determines the background gradient color of the Microsoft Visio stencil window for the current session. Read/write. |
|[StencilCharactersPerLine](applicationsettings-stencilcharactersperline-property-visio.md)|For shapes on stencils, determines approximately how many characters of each shape's name appear on each line before the text wraps to the next line. Read/write.|
|[StencilLinesPerMaster](applicationsettings-stencillinespermaster-property-visio.md)|For shapes on stencils in Microsoft Visio, determines how many lines of text of each shape's name can appear below the shape before the text is truncated and "..." is appended. Read/write.|
|[StencilTextColor](applicationsettings-stenciltextcolor-property-visio.md)|Determines the color of text in stencil windows in Microsoft Visio for the current session. Read/write.|
|[SVGExportFormat](applicationsettings-svgexportformat-property-visio.md)|Returns or sets a [VisSVGExportFormat](vissvgexportformat-enumeration-visio.md) constant that specifies the options for saving a Microsoft Visio file in SVG format. Read/write.|
|[TransitionsEnabled](applicationsettings-transitionsenabled-property-visio.md)|Determines whether Microsoft Visio uses an animated transition to show certain shape movements, such as re-layout of shapes. Read/write.|
|[UndoLevels](applicationsettings-undolevels-property-visio.md)|Determines the number of consecutive actions the user can undo in Microsoft Visio. Read/write.|
|||
|[UserInitials](applicationsettings-userinitials-property-visio.md)|Determines the user initials associated with the Microsoft Visio file. Read/write.|
|[UserName](applicationsettings-username-property-visio.md)|Gets or sets the user name of an  **Application** object. Read/write.|
|[ZoomOnRoll](applicationsettings-zoomonroll-property-visio.md)|Determines whether zooming in to and out from a Microsoft Visio drawing by rolling the wheel of the mouse is enabled. Read/write.|

