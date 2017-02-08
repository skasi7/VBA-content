---
title: Attachment Object (Access)
keywords: vbaac10.chm14036
f1_keywords:
- vbaac10.chm14036
ms.prod: ACCESS
api_name:
- Access.Attachment
ms.assetid: b0756145-9012-f9b9-7df9-e168defed3bf
---


# Attachment Object (Access)

This object corresponds to an attachment control. Use an attachment control when you want to manipulate the contents fields of the attachment data type.


## Remarks


 **Note**  You can attach files only to databases that you create in Office Access 2007 and that use the new .accdb file format. You cannot share attachments between a Office Access 2007 (.accdb) database and a database in the earlier (.mdb) file format.

You can attach a maximum of two gigabytes of data (the maximum size for an Access database). Individual files cannot exceed 256 megabytes in size.


### Supported image file formats

Office Access 2007 supports the following graphic file formats natively, meaning the attachment control renders them without the need for additional software.


- BMP (Windows Bitmap)
    
- RLE (Run Length Encoded Bitmap)
    
- DIB (Device Independent Bitmap)
    
- GIF (Graphics Interchange Format)
    
- JPEG, JPG, JPE (Joint Photographic Experts Group)
    
- EXIF (Exchangeable File Format)
    
- PNG (Portable Network Graphics)
    
- TIFF, TIF (Tagged Image File Format)
    
- ICON, ICO (Icon)
    
- WMF (Windows Metafile)
    
- EMF (Enhanced Metafile)
    

### Supported formats for documents and other files

As a rule, you can attach any file that was created with one of the 2007 Microsoft Office system programs. You can also attach log files (.log), text files (.text, .txt), and compressed .zip files.


### File-naming conventions

The names of your attached files can contain any Unicode character supported by the NTFS file system used in Microsoft Windows NT (NTFS). In addition, file names must conform to these guidelines:


- Names must not exceed 255 characters, including the file name extensions.
    
- Names cannot contain the following characters: question marks (?), quotation marks ("), forward or backward slashes (/ \), opening or closing brackets (< >), asterisks (*), vertical bars or pipes (|), colons (:), or paragraph marks (?).
    

### Types of files that Access compresses

Access will compress your attached files unless those files are compressed natively. For example, JPEG files are compressed by the graphics program that created them, so Access does not compress them. the following table lists some supported file types and whether or not Access compresses them.



|**File Extension**|**Compressed?**|**Reason**|
|:-----|:-----|:-----|
|.jpg, .jpeg|No|Already compressed|
|.gif|No|Already compressed|
|.png|No|Already compressed|
|.tif, .tiff|Yes||
|.exif|Yes||
| .bmp|Yes||
|.emf|Yes||
|.wmf|Yes||
|.ico|Yes||
|.zip|No|Already compressed|
|.cab|No|Already compressed|
|.docx|No|Already compressed|
|.xlsx|No|Already compressed|
|.xlsb|No|Already compressed|
|.pptx|No|Already compressed|

### Blocked file formats

Office Access 2007 blocks the following types of attached files. At this time, you cannot unblock any of the file types listed here.


|||||
|:-----|:-----|:-----|:-----|
|.ade|.ins|.mda|.scr|
|.adp|.isp|.mdb|.sct|
|.app|.its|.mde|.shb|
|.asp|.js|.mdt|.shs|
|.bas|.jse|.mdw|.tmp|
|.bat|.ksh|.mdz|.url|
|.cer|.lnk|.msc|.vb|
|.chm|.mad|.msi|.vbe|
|.cmd|.maf|.msp|.vbs|
|.com|.mag|.mst|.vsmacros|
|.cpl|.mam|.ops|.vss|
|.crt|.maq|.pcd|.vst|
|.csh|.mar|.pif|.vsw|
|.exe|.mas|.prf|.ws|
|.fxp|.mat|.prg|.wsc|
|.hlp|.mau|.pst|.wsf|
|.hta|.mav|.reg|.wsh|
|.inf|.maw|.scf||

## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/attachment-afterupdate-event-access%28Office.15%29.aspx)|
|[AttachmentCurrent](http://msdn.microsoft.com/library/attachment-attachmentcurrent-event-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/attachment-beforeupdate-event-access%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/attachment-change-event-access%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/attachment-click-event-access%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/attachment-dblclick-event-access%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/attachment-dirty-event-access%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/attachment-enter-event-access%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/attachment-exit-event-access%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/attachment-gotfocus-event-access%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/attachment-keydown-event-access%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/attachment-keypress-event-access%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/attachment-keyup-event-access%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/attachment-lostfocus-event-access%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/attachment-mousedown-event-access%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/attachment-mousemove-event-access%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/attachment-mouseup-event-access%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Back](http://msdn.microsoft.com/library/attachment-back-method-access%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/attachment-forward-method-access%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/attachment-move-method-access%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/attachment-requery-method-access%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/attachment-setfocus-method-access%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/attachment-sizetofit-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/attachment-addcolon-property-access%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/attachment-afterupdate-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/attachment-application-property-access%28Office.15%29.aspx)|
|[AttachmentCount](http://msdn.microsoft.com/library/attachment-attachmentcount-property-access%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/attachment-autolabel-property-access%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/attachment-backcolor-property-access%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/attachment-backshade-property-access%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/attachment-backstyle-property-access%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/attachment-backthemecolorindex-property-access%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/attachment-backtint-property-access%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/attachment-beforeupdate-property-access%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/attachment-bordercolor-property-access%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/attachment-bordershade-property-access%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/attachment-borderstyle-property-access%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/attachment-borderthemecolorindex-property-access%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/attachment-bordertint-property-access%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/attachment-borderwidth-property-access%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/attachment-bottompadding-property-access%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/attachment-columnhidden-property-access%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/attachment-columnorder-property-access%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/attachment-columnwidth-property-access%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/attachment-controls-property-access%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/attachment-controlsource-property-access%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/attachment-controltiptext-property-access%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/attachment-controltype-property-access%28Office.15%29.aspx)|
|[CurrentAttachment](http://msdn.microsoft.com/library/attachment-currentattachment-property-access%28Office.15%29.aspx)|
|[DefaultPicture](http://msdn.microsoft.com/library/attachment-defaultpicture-property-access%28Office.15%29.aspx)|
|[DefaultPictureType](http://msdn.microsoft.com/library/attachment-defaultpicturetype-property-access%28Office.15%29.aspx)|
|[DisplayAs](http://msdn.microsoft.com/library/attachment-displayas-property-access%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/attachment-displaywhen-property-access%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/attachment-enabled-property-access%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/attachment-eventprocprefix-property-access%28Office.15%29.aspx)|
|[FileName](http://msdn.microsoft.com/library/attachment-filename-property-access%28Office.15%29.aspx)|
|[FileType](http://msdn.microsoft.com/library/attachment-filetype-property-access%28Office.15%29.aspx)|
|[FileURL](http://msdn.microsoft.com/library/attachment-fileurl-property-access%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/attachment-gridlinecolor-property-access%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/attachment-gridlineshade-property-access%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/attachment-gridlinestylebottom-property-access%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/attachment-gridlinestyleleft-property-access%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/attachment-gridlinestyleright-property-access%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/attachment-gridlinestyletop-property-access%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/attachment-gridlinethemecolorindex-property-access%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/attachment-gridlinetint-property-access%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/attachment-gridlinewidthbottom-property-access%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/attachment-gridlinewidthleft-property-access%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/attachment-gridlinewidthright-property-access%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/attachment-gridlinewidthtop-property-access%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/attachment-height-property-access%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/attachment-helpcontextid-property-access%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/attachment-horizontalanchor-property-access%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/attachment-inselection-property-access%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/attachment-isvisible-property-access%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/attachment-labelalign-property-access%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/attachment-labelx-property-access%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/attachment-labely-property-access%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/attachment-layout-property-access%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/attachment-layoutid-property-access%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/attachment-left-property-access%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/attachment-leftpadding-property-access%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/attachment-locked-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/attachment-name-property-access%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/attachment-oldborderstyle-property-access%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/attachment-oldvalue-property-access%28Office.15%29.aspx)|
|[OnAttachmentCurrent](http://msdn.microsoft.com/library/attachment-onattachmentcurrent-property-access%28Office.15%29.aspx)|
|[OnChange](http://msdn.microsoft.com/library/attachment-onchange-property-access%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/attachment-onclick-property-access%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/attachment-ondblclick-property-access%28Office.15%29.aspx)|
|[OnDirty](http://msdn.microsoft.com/library/attachment-ondirty-property-access%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/attachment-onenter-property-access%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/attachment-onexit-property-access%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/attachment-ongotfocus-property-access%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/attachment-onkeydown-property-access%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/attachment-onkeypress-property-access%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/attachment-onkeyup-property-access%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/attachment-onlostfocus-property-access%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/attachment-onmousedown-property-access%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/attachment-onmousemove-property-access%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/attachment-onmouseup-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/attachment-parent-property-access%28Office.15%29.aspx)|
|[PictureAlignment](http://msdn.microsoft.com/library/attachment-picturealignment-property-access%28Office.15%29.aspx)|
|[PictureSizeMode](http://msdn.microsoft.com/library/attachment-picturesizemode-property-access%28Office.15%29.aspx)|
|[PictureTiling](http://msdn.microsoft.com/library/attachment-picturetiling-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/attachment-properties-property-access%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/attachment-rightpadding-property-access%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/attachment-section-property-access%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/attachment-shortcutmenubar-property-access%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/attachment-specialeffect-property-access%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/attachment-statusbartext-property-access%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/attachment-tabindex-property-access%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/attachment-tabstop-property-access%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/attachment-tag-property-access%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/attachment-top-property-access%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/attachment-toppadding-property-access%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/attachment-verticalanchor-property-access%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/attachment-visible-property-access%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/attachment-width-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Attachment Object Members](http://msdn.microsoft.com/library/attachment-members-access%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
