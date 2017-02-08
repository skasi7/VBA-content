---
title: Image Object (Access)
keywords: vbaac10.chm10436
f1_keywords:
- vbaac10.chm10436
ms.prod: ACCESS
api_name:
- Access.Image
ms.assetid: 1bcc8552-94e2-b799-6903-392205cb4341
---


# Image Object (Access)

This object corresponds to an image control. The image control can add a picture to a form or report. For example, you could include an image control for a logo on an Invoice report.

 **Note**: The functionality for the Image object's **image.click** and **image.doubleclick** events have been deprecated. If you want an image with click/double click events, instead use a Button control and associate an image with that control as that provides better accessibility. Button controls are part of the Tab Order loop but Image controls are not. Existing applications will not be affected by this change.

## Remarks


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|
![Image control](/images/t-imgctl_ZA06053959.gif)

|
![Image tool](/images/imagefrm_ZA06044465.gif)

|
You can use the image control or an [Unbound object frame](http://msdn.microsoft.com/library/unbound-object-frame-control%28Office.15%29.aspx)for unbound pictures. The advantage of using the image control is that it's faster to display. The advantage of using the unbound object frame is that you can edit the object directly from the form or report.


## Events



|**Name**|
|:-----|
|[Click](http://msdn.microsoft.com/library/image-click-event-access%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/image-dblclick-event-access%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/image-mousedown-event-access%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/image-mousemove-event-access%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/image-mouseup-event-access%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/image-move-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/image-requery-method-access%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/image-setfocus-method-access%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/image-sizetofit-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/image-application-property-access%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/image-backcolor-property-access%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/image-backshade-property-access%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/image-backstyle-property-access%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/image-backthemecolorindex-property-access%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/image-backtint-property-access%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/image-bordercolor-property-access%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/image-bordershade-property-access%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/image-borderstyle-property-access%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/image-borderthemecolorindex-property-access%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/image-bordertint-property-access%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/image-borderwidth-property-access%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/image-bottompadding-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/image-controls-property-access%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/image-controltiptext-property-access%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/image-controltype-property-access%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/image-displaywhen-property-access%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/image-eventprocprefix-property-access%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/image-gridlinecolor-property-access%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/image-gridlineshade-property-access%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/image-gridlinestylebottom-property-access%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/image-gridlinestyleleft-property-access%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/image-gridlinestyleright-property-access%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/image-gridlinestyletop-property-access%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/image-gridlinethemecolorindex-property-access%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/image-gridlinetint-property-access%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/image-gridlinewidthbottom-property-access%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/image-gridlinewidthleft-property-access%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/image-gridlinewidthright-property-access%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/image-gridlinewidthtop-property-access%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/image-height-property-access%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/image-helpcontextid-property-access%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/image-horizontalanchor-property-access%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/image-hyperlink-property-access%28Office.15%29.aspx)|
|[HyperlinkAddress](http://msdn.microsoft.com/library/image-hyperlinkaddress-property-access%28Office.15%29.aspx)|
|[HyperlinkSubAddress](http://msdn.microsoft.com/library/image-hyperlinksubaddress-property-access%28Office.15%29.aspx)|
|[ImageHeight](http://msdn.microsoft.com/library/image-imageheight-property-access%28Office.15%29.aspx)|
|[ImageWidth](http://msdn.microsoft.com/library/image-imagewidth-property-access%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/image-inselection-property-access%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/image-isvisible-property-access%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/image-layout-property-access%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/image-layoutid-property-access%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/image-left-property-access%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/image-leftpadding-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/image-name-property-access%28Office.15%29.aspx)|
|[ObjectPalette](http://msdn.microsoft.com/library/image-objectpalette-property-access%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/image-oldborderstyle-property-access%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/image-oldvalue-property-access%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/image-onclick-property-access%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/image-ondblclick-property-access%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/image-onmousedown-property-access%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/image-onmousemove-property-access%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/image-onmouseup-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/image-parent-property-access%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/image-picture-property-access%28Office.15%29.aspx)|
|[PictureAlignment](http://msdn.microsoft.com/library/image-picturealignment-property-access%28Office.15%29.aspx)|
|[PictureData](http://msdn.microsoft.com/library/image-picturedata-property-access%28Office.15%29.aspx)|
|[PictureTiling](http://msdn.microsoft.com/library/image-picturetiling-property-access%28Office.15%29.aspx)|
|[PictureType](http://msdn.microsoft.com/library/image-picturetype-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/image-properties-property-access%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/image-rightpadding-property-access%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/image-section-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/image-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[SizeMode](http://msdn.microsoft.com/library/image-sizemode-property-access%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/image-specialeffect-property-access%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/image-tag-property-access%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/image-top-property-access%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/image-toppadding-property-access%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/image-verticalanchor-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/image-visible-property-access%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/image-width-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[Image Object Members](http://msdn.microsoft.com/library/image-members-access%28Office.15%29.aspx)
