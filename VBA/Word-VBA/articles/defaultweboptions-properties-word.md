---
title: DefaultWebOptions Properties (Word)
ms.prod: WORD
ms.assetid: 73e69735-17c1-4d6c-bd56-e5e4c2417a29
---


# DefaultWebOptions Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowPNG](defaultweboptions-allowpng-property-word.md)| **False** if PNG (Portable Network Graphics) is not allowed as an output format. Read/write **Boolean** .|
|[AlwaysSaveInDefaultEncoding](defaultweboptions-alwayssaveindefaultencoding-property-word.md)| **True** if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened. Read/write **Boolean** .|
|[Application](defaultweboptions-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[BrowserLevel](defaultweboptions-browserlevel-property-word.md)|Returns or sets a  **WdBrowserLevel** constant that represents the level of the Web browser for which you want to target new Web pages created in Microsoft Word. Read/write.|
|[CheckIfOfficeIsHTMLEditor](defaultweboptions-checkifofficeishtmleditor-property-word.md)| **True** if Microsoft Word checks to see whether an Office application is the default HTML editor when you start Word. Read/write **Boolean** .|
|[CheckIfWordIsDefaultHTMLEditor](defaultweboptions-checkifwordisdefaulthtmleditor-property-word.md)| **True** if Microsoft Word checks to see whether it is the default HTML editor when you start Word. Read/write **Boolean** .|
|[Creator](defaultweboptions-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Encoding](defaultweboptions-encoding-property-word.md)|Returns or sets the document encoding (code page or character set) to be used by the Web browser when you view the saved document. Read/write  **MsoEncoding** .|
|[FolderSuffix](defaultweboptions-foldersuffix-property-word.md)|Returns a  **String** that represents the folder suffix that Microsoft Word uses when you save a document as a Web page, use long file names, or save supporting files in a separate folder. Read-only.|
|[Fonts](defaultweboptions-fonts-property-word.md)|Returns the  **WebPageFonts** collection representing the set of fonts that Microsoft Word uses when you open a Web page in Word.|
|[OptimizeForBrowser](defaultweboptions-optimizeforbrowser-property-word.md)| **True** if Microsoft Word optimizes new Web pages created in Word for the Web browser specified by the **[BrowserLevel](defaultweboptions-browserlevel-property-word.md)** property. Read/write **Boolean** .|
|[OrganizeInFolder](defaultweboptions-organizeinfolder-property-word.md)| **True** if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a Web page. **False** if supporting files are saved in the same folder as the Web page. The default value is **True** . Read/write **Boolean** .|
|[Parent](defaultweboptions-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **DefaultWebOptions** object.|
|[PixelsPerInch](defaultweboptions-pixelsperinch-property-word.md)|Returns or sets the density (pixels per inch) of graphics images and table cells on a Web page. Read/write  **Long** .|
|[RelyOnCSS](defaultweboptions-relyoncss-property-word.md)| **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. Read/write **Boolean** .|
|[RelyOnVML](defaultweboptions-relyonvml-property-word.md)| **True** if image files are not generated from drawing objects when you save a document as a Web page. **False** if images are generated. The default value is **False** . Read/write **Boolean** .|
|[SaveNewWebPagesAsWebArchives](defaultweboptions-savenewwebpagesaswebarchives-property-word.md)| **True** for Microsoft Word to save new Web pages in the Single File Web Page (formerly known as Web Archive) format. Read/write **Boolean** .|
|[ScreenSize](defaultweboptions-screensize-property-word.md)|Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a Web browser. Read/write  **MsoScreenSize** .|
|[TargetBrowser](defaultweboptions-targetbrowser-property-word.md)|Sets or returns an  **MsoTargetBrowser** constant representing the target browser for documents viewed in a Web browser. Read/write.|
|[UpdateLinksOnSave](defaultweboptions-updatelinksonsave-property-word.md)| **True** if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page. Read/write **Boolean** .|
|[UseLongFileNames](defaultweboptions-uselongfilenames-property-word.md)| **True** if long file names are used when you save the document as a Web page. **False** if long file names are not used and the DOS file name format (8.3) is used. The default value is **True** . Read/write **Boolean** .|

