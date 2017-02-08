---
title: WebOptions Members (Word)
ms.prod: WORD
ms.assetid: f4fb7f5c-d82a-3a94-bcae-9e9f1fb43872
---


# WebOptions Members (Word)
Contains document-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page.

Contains document-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[UseDefaultFolderSuffix](weboptions-usedefaultfoldersuffix-method-word.md)|Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowPNG](weboptions-allowpng-property-word.md)| **True** if PNG (Portable Network Graphics) is allowed as an image format when you save a document as a Web page. **False** if PNG is not allowed as an output format. The default value is **False** . Read/write **Boolean** .|
|[Application](weboptions-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[BrowserLevel](weboptions-browserlevel-property-word.md)|Returns or sets  **WdBrowserLevel** that represents the level of Web browser at which you want to target the specified Web page. Read/write.|
|[Creator](weboptions-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Encoding](weboptions-encoding-property-word.md)|Returns or sets the document encoding (code page or character set) to be used by the Web browser when you view the saved document. Read/write  **MsoEncoding** .|
|[FolderSuffix](weboptions-foldersuffix-property-word.md)|Returns the folder suffix that Microsoft Word uses when you save a document as a Web page, use long file names, and choose to save supporting files in a separate folder (that is, if the  **UseLongFileNames** and **OrganizeInFolder** properties are set to **True** ). Read-only **String** .|
|[OptimizeForBrowser](weboptions-optimizeforbrowser-property-word.md)| **True** if Word optimizes the specified Web page for the Web browser specified by the **[BrowserLevel](weboptions-browserlevel-property-word.md)** property. Read/write **Boolean** .|
|[OrganizeInFolder](weboptions-organizeinfolder-property-word.md)| **True** if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a Web page. **False** if supporting files are saved in the same folder as the Web page. The default value is **True** . Read/write **Boolean** .|
|[Parent](weboptions-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **WebOptions** object.|
|[PixelsPerInch](weboptions-pixelsperinch-property-word.md)|Returns or sets the density (pixels per inch) of graphics images and table cells on a Web page. Read/write  **Long** .|
|[RelyOnCSS](weboptions-relyoncss-property-word.md)| **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. The default value is **True** . Read/write **Boolean** .|
|[RelyOnVML](weboptions-relyonvml-property-word.md)| **True** if image files are not generated from drawing objects when you save a document as a Web page. **False** if images are generated. The default value is **False** . Read/write **Boolean** .|
|[ScreenSize](weboptions-screensize-property-word.md)|Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a Web browser. Read/write  **MsoScreenSize** .|
|[TargetBrowser](weboptions-targetbrowser-property-word.md)|Sets or returns an  **MsoTargetBrowser** constant representing the target browser for documents viewed in a Web browser. Read/write.|
|[UseLongFileNames](weboptions-uselongfilenames-property-word.md)| **True** if long file names are used when you save the document as a Web page. **False** if long file names are not used and the DOS file name format (8.3) is used. The default value is **True** . Read/write **Boolean** .|

