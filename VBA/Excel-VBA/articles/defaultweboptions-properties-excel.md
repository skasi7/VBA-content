---
title: DefaultWebOptions Properties (Excel)
ms.prod: EXCEL
ms.assetid: a8177c8a-c5c1-4cd7-ba31-0d0981bb2662
---


# DefaultWebOptions Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowPNG](defaultweboptions-allowpng-property-excel.md)| **True** if PNG (Portable Network Graphics) is allowed as an image format when you save documents as a Web page. **False** if PNG is not allowed as an output format. The default value is **False** . Read/write **Boolean** .|
|[AlwaysSaveInDefaultEncoding](defaultweboptions-alwayssaveindefaultencoding-property-excel.md)| **True** if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened. **False** if the original encoding of the file is used. The default value is **False** . Read/write **Boolean** .|
|[Application](defaultweboptions-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[CheckIfOfficeIsHTMLEditor](defaultweboptions-checkifofficeishtmleditor-property-excel.md)| **True** if Microsoft Excel checks to see whether an Office application is the default HTML editor when you start Excel. **False** if Excel does not perform this check. The default value is **True** . Read/write **Boolean** .|
|[Creator](defaultweboptions-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DownloadComponents](defaultweboptions-downloadcomponents-property-excel.md)| **True** if the necessary Microsoft Office Web components are downloaded when you view the saved document in a Web browser, but only if the components are not already installed. **False** if the components are not downloaded. The default value is **False** . Read/write **Boolean** .|
|[Encoding](defaultweboptions-encoding-property-excel.md)|Returns or sets the document encoding (code page or character set) to be used by the Web browser when you view the saved document. The default is the system code page. Read/write  **[MsoEncoding](msoencoding-enumeration-office.md)** .|
|[FolderSuffix](defaultweboptions-foldersuffix-property-excel.md)|Returns the folder suffix that Microsoft Excel uses when you save a document as a Web page, use long file names, and choose to save supporting files in a separate folder (that is, if the  **[UseLongFileNames](defaultweboptions-uselongfilenames-property-excel.md)** and **[OrganizeInFolder](defaultweboptions-organizeinfolder-property-excel.md)** properties are set to **True** ). Read-only **String** .|
|[Fonts](defaultweboptions-fonts-property-excel.md)|Returns the  **[WebPageFonts](webpagefonts-object-office.md)** collection representing the set of fonts Microsoft Excel uses when you open a Web page in Excel and there is either no font information specified in the Web page, or the current default font can't display the character set in the Web page. Read-only.|
|[LoadPictures](defaultweboptions-loadpictures-property-excel.md)| **True** if images are loaded when you open a document in Microsoft Excel, usually when the images and document were not created in Microsoft Excel. **False** if the images are not loaded. The default value is **True** . Read/write **Boolean** .|
|[LocationOfComponents](defaultweboptions-locationofcomponents-property-excel.md)|Returns or sets the central URL (on the intranet or Web) or path (local or network) to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. The default value is the local or network installation path for Microsoft Office. Read/write  **String** .|
|[OrganizeInFolder](defaultweboptions-organizeinfolder-property-excel.md)| **True** if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a Web page. **False** if supporting files are saved in the same folder as the Web page. The default value is **True** . Read/write **Boolean** .|
|[Parent](defaultweboptions-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PixelsPerInch](defaultweboptions-pixelsperinch-property-excel.md)|Returns or sets the density (pixels per inch) of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. The default setting is 96. Read/write  **Long** .|
|[RelyOnCSS](defaultweboptions-relyoncss-property-excel.md)| **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. Microsoft Excel creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your Web page, depending on the value of the **[OrganizeInFolder](defaultweboptions-organizeinfolder-property-excel.md)** property. **False** if HTML <FONT> tags and cascading style sheets are used. The default value is **True** . Read/write **Boolean** .|
|[RelyOnVML](defaultweboptions-relyonvml-property-excel.md)| **True** if image files are not generated from drawing objects when you save a document as a Web page. **False** if images are generated. The default value is **False** . Read/write **Boolean** .|
|[SaveHiddenData](defaultweboptions-savehiddendata-property-excel.md)| **True** if data outside of the specified range is saved when you save the document as a Web page. This data may be necessary for maintaining formulas. **False** if data outside of the specified range is not saved with the Web page. The default value is **True** . Read/write **Boolean** .|
|[SaveNewWebPagesAsWebArchives](defaultweboptions-savenewwebpagesaswebarchives-property-excel.md)| **True** if new Web pages can be saved as Web archives. Read/write **Boolean** .|
|[ScreenSize](defaultweboptions-screensize-property-excel.md)|Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a Web browser. Can be one of the  **[MsoScreenSize](msoscreensize-enumeration-office.md)** constants. The default constant is **msoScreenSize800x600** . Read/write **MsoScreenSize** .|
|[TargetBrowser](defaultweboptions-targetbrowser-property-excel.md)|Returns or sets an  **[MsoTargetBrowser](msotargetbrowser-enumeration-office.md)** constant indicating the browser version. Read/write.|
|[UpdateLinksOnSave](defaultweboptions-updatelinksonsave-property-excel.md)| **True** if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved. **False** if the links are not updated. The default value is **True** . Read/write **Boolean** .|
|[UseLongFileNames](defaultweboptions-uselongfilenames-property-excel.md)| **True** if long file names are used when you save the document as a Web page. **False** if long file names are not used and the DOS file name format (8.3) is used. The default value is **True** . Read/write **Boolean** .|

