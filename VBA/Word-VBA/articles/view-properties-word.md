---
title: View Properties (Word)
ms.prod: WORD
ms.assetid: 950bc476-c1bd-4011-80d6-c3d30a2522e6
---


# View Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](view-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[ColumnWidth](view-columnwidth-property-word.md)|Returns or gets a constant that determines the column width in reading mode. Read-write [WdColumnWidth](wdcolumnwidth-enumeration-word.md)|
|[ConflictMode](view-conflictmode-property-word.md)| **True** if the document is in conflict mode view. Read/write.|
|[Creator](view-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DisplayBackgrounds](view-displaybackgrounds-property-word.md)|Returns or sets a  **Boolean** that represents whether background colors and images are shown when a document is displayed in print layout view. .|
|[DisplayPageBoundaries](view-displaypageboundaries-property-word.md)| **True** to display the top and bottom margins (white space) and the gray area (gray space) between pages in a document. Read/write **Boolean** .|
|[Draft](view-draft-property-word.md)| **True** if all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display. Read/write **Boolean** .|
|[FieldShading](view-fieldshading-property-word.md)|Returns or sets on-screen shading for fields. Read/write  **WdFieldShading** .|
|[FullScreen](view-fullscreen-property-word.md)| **True** if the window is in full-screen view. Read/write **Boolean** .|
|[Magnifier](view-magnifier-property-word.md)| **True** if the pointer is displayed as a magnifying glass in print preview, indicating that the user can click to zoom in on a particular area of the page or zoom out to see an entire page or spread of pages. Read/write **Boolean** .|
|[MailMergeDataView](view-mailmergedataview-property-word.md)| **True** if mail merge data is displayed instead of mail merge fields in the specified window. Read/write **Boolean** .|
|[MarkupMode](view-markupmode-property-word.md)|Returns or sets a  **[WdRevisionsMode](wdrevisionsmode-enumeration-word.md)** constant that represents the display mode for tracked changes. Read/write.|
|[PageColor](view-pagecolor-property-word.md)|Returns and sets the page color in Reading mode. Read-write [WdPageColor](wdpagecolor-enumeration-word.md).|
|[Panning](view-panning-property-word.md)|Returns or sets a  **Boolean** that represents whether Microsoft Word is in Panning mode. Read/write.|
|[Parent](view-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **View** object.|
|[ReadingLayout](view-readinglayout-property-word.md)|Sets or returns a  **Boolean** that represents whether a document is being viewed in reading layout view. .|
|[ReadingLayoutActualView](view-readinglayoutactualview-property-word.md)|Sets or returns a  **Boolean** that represents whether pages displayed in reading layout view are displayed using the same layout as printed pages.|
|[ReadingLayoutTruncateMargins](view-readinglayouttruncatemargins-property-word.md)|Returns or sets a  **[WdReadingLayoutMargin](wdreadinglayoutmargin-enumeration-word.md)** constant that represents whether margins are visible or hidden when a document is viewed in Full Screen Reading view. Read/write.|
|[RevisionsBalloonShowConnectingLines](view-revisionsballoonshowconnectinglines-property-word.md)| **True** for Microsoft Word to display connecting lines from the text to the revision and comment balloons. Read/write **Boolean** .|
|[RevisionsBalloonSide](view-revisionsballoonside-property-word.md)|Sets or returns a  **WdRevisionsBalloonMargin** constant that specifies whether Word displays revision balloons in the left or right margin in a document.|
|[RevisionsBalloonWidth](view-revisionsballoonwidth-property-word.md)|Sets or returns a  **Single** representing the global setting in Microsoft Word that specifies the width of the revision balloons. Read/write.|
|[RevisionsBalloonWidthType](view-revisionsballoonwidthtype-property-word.md)|Sets or returns a  **WdRevisionsBalloonWidthType** constant representing the global setting that specifies how Microsoft Word measures the width of revision balloons. Read/write.|
|[RevisionsFilter](view-revisionsfilter-property-word.md)|Returns an instance of a  **RevisionsFilter** object. Read-only.|
|[SeekView](view-seekview-property-word.md)|Returns or sets the document element displayed in print layout view. Read/write  **WdSeekView** .|
|[ShadeEditableRanges](view-shadeeditableranges-property-word.md)|Returns or sets a  **Long** that represents whether shading is applied to the ranges in a document for which users have permission to modify. .|
|[ShowAll](view-showall-property-word.md)| **True** if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed. Read/write **Boolean** .|
|[ShowBookmarks](view-showbookmarks-property-word.md)| **True** if square brackets are displayed at the beginning and end of each bookmark. Read/write **Boolean** .|
|[ShowComments](view-showcomments-property-word.md)| **True** for Microsoft Word to display the comments in a document. Read/write **Boolean** .|
|[ShowCropMarks](view-showcropmarks-property-word.md)|Returns or sets a  **Boolean** that represents whether to show crop marks in the corners of pages to indicate where margins are located. Read/write.|
|[ShowDrawings](view-showdrawings-property-word.md)| **True** if objects created with the drawing tools are displayed in print layout view. Read/write **Boolean** .|
|[ShowFieldCodes](view-showfieldcodes-property-word.md)| **True** if field codes are displayed. Read/write **Boolean** .|
|[ShowFirstLineOnly](view-showfirstlineonly-property-word.md)| **True** if only the first line of body text is shown in outline view. Read/write **Boolean** .|
|[ShowFormat](view-showformat-property-word.md)| **True** if character formatting is visible in outline view. Read/write **Boolean** .|
|[ShowFormatChanges](view-showformatchanges-property-word.md)| **True** for Microsoft Word to display formatting changes made to a document with Track Changes enabled. Read/write **Boolean** .|
|[ShowHiddenText](view-showhiddentext-property-word.md)| **True** if text formatted as hidden text is displayed. Read/write **Boolean** .|
|[ShowHighlight](view-showhighlight-property-word.md)| **True** if highlight formatting is displayed and printed with a document. Read/write **Boolean** .|
|[ShowHyphens](view-showhyphens-property-word.md)| **True** if optional hyphens are displayed. An optional hyphen indicates where to break a word when it falls at the end of a line. Read/write **Boolean** .|
|[ShowInkAnnotations](view-showinkannotations-property-word.md)|Returns or sets  **Boolean** that shows or hides handwritten ink annotations. **True** displays ink annotations. **False** hides ink annotations.|
|[ShowInsertionsAndDeletions](view-showinsertionsanddeletions-property-word.md)| **True** for Microsoft Word to display insertions and deletions that were made to a document with Track Changes enabled. Read/write **Boolean** .|
|[ShowMainTextLayer](view-showmaintextlayer-property-word.md)| **True** if the text in the specified document is visible when the header and footer areas are displayed. This property is equivalent to the **Show/Hide Document Text** button on the **Header and Footer** toolbar. Read/write **Boolean** .|
|[ShowMarkupAreaHighlight](view-showmarkupareahighlight-property-word.md)|Returns or sets a  **Boolean** that represents whether the markup area that shows revision and comment balloons is shaded. Read/write.|
|[ShowObjectAnchors](view-showobjectanchors-property-word.md)| **True** if object anchors are displayed next to items that can be positioned in print layout view. Read/write **Boolean** .|
|[ShowOptionalBreaks](view-showoptionalbreaks-property-word.md)| **True** if Microsoft Word displays optional line breaks. Read/write **Boolean** .|
|[ShowOtherAuthors](view-showotherauthors-property-word.md)| **True** if other authors' presence should be visible in the document. Read/write **Boolean** .|
|[ShowParagraphs](view-showparagraphs-property-word.md)| **True** if paragraph marks are displayed. Read/write **Boolean** .|
|[ShowPicturePlaceHolders](view-showpictureplaceholders-property-word.md)| **True** if blank boxes are displayed as placeholders for pictures. Read/write **Boolean** .|
|[ShowRevisionsAndComments](view-showrevisionsandcomments-property-word.md)| **True** for Microsoft Word to display revisions and comments that were made to a document with Track Changes enabled. Read/write **Boolean** .|
|[ShowSpaces](view-showspaces-property-word.md)| **True** if space characters are displayed. Read/write **Boolean** .|
|[ShowTabs](view-showtabs-property-word.md)| **True** if tab characters are displayed. Read/write **Boolean** .|
|[ShowTextBoundaries](view-showtextboundaries-property-word.md)| **True** if dotted lines are displayed around page margins, text columns, objects, and frames in print layout view. Read/write **Boolean** .|
|[ShowXMLMarkup](view-showxmlmarkup-property-word.md)|Returns a  **Long** that represents whether XML tags are visible in a document.|
|[SplitSpecial](view-splitspecial-property-word.md)|Returns or sets the active window pane. Read/write  **WdSpecialPane** .|
|[TableGridlines](view-tablegridlines-property-word.md)| **True** if table gridlines are displayed. Read/write **Boolean** .|
|[Type](view-type-property-word.md)|Returns or sets the view type. Read/write  **[WdViewType](wdviewtype-enumeration-word.md)** .|
|[WrapToWindow](view-wraptowindow-property-word.md)| **True** if lines wrap at the right edge of the document window rather than at the right margin or the right column boundary. Read/write **Boolean** .|
|[Zoom](view-zoom-property-word.md)|Returns a  **[Zoom](zoom-object-word.md)** object that represents the magnification for the specified view.|

