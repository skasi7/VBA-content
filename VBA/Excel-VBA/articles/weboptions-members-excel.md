---
title: WebOptions Members (Excel)
ms.prod: EXCEL
ms.assetid: 4188ab11-5d84-aed8-2a2e-17881dcebe67
---


# WebOptions Members (Excel)
Contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.

Contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[UseDefaultFolderSuffix](weboptions-usedefaultfoldersuffix-method-excel.md)|Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowPNG](weboptions-allowpng-property-excel.md)| **True** if PNG (Portable Network Graphics) is allowed as an image format when you save documents as a Web page. **False** if PNG is not allowed as an output format. The default value is **False** . Read/write **Boolean** .|
|[Application](weboptions-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](weboptions-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DownloadComponents](weboptions-downloadcomponents-property-excel.md)| **True** if the necessary Microsoft Office Web components are downloaded when you view the saved document in a Web browser, but only if the components are not already installed. **False** if the components are not downloaded. The default value is **False** . Read/write **Boolean** .|
|[Encoding](weboptions-encoding-property-excel.md)|Returns or sets the document encoding (code page or character set) to be used by the Web browser when you view the saved document. The default is the system code page. Read/write  **[MsoEncoding](msoencoding-enumeration-office.md)** .|
|[FolderSuffix](weboptions-foldersuffix-property-excel.md)|Returns the folder suffix that Microsoft Excel uses when you save a document as a Web page, use long file names, and choose to save supporting files in a separate folder (that is, if the  **[UseLongFileNames](weboptions-uselongfilenames-property-excel.md)** and **[OrganizeInFolder](weboptions-organizeinfolder-property-excel.md)** properties are set to **True** ). Read-only **String** .|
|[LocationOfComponents](weboptions-locationofcomponents-property-excel.md)|Returns or sets the central URL (on the intranet or Web) or path (local or network) to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. The default value is the local or network installation path for Microsoft Office. Read/write  **String** .|
|[OrganizeInFolder](weboptions-organizeinfolder-property-excel.md)| **True** if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a Web page. **False** if supporting files are saved in the same folder as the Web page. The default value is **True** . Read/write **Boolean** .|
|[Parent](weboptions-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PixelsPerInch](weboptions-pixelsperinch-property-excel.md)|Returns or sets the density (pixels per inch) of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. The default setting is 96. Read/write  **Long** .|
|[RelyOnCSS](weboptions-relyoncss-property-excel.md)| **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. Microsoft Excel creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your Web page, depending on the value of the **[OrganizeInFolder](weboptions-organizeinfolder-property-excel.md)** property. **False** if HTML <FONT> tags and cascading style sheets are used. The default value is **True** . Read/write **Boolean** .|
|[RelyOnVML](weboptions-relyonvml-property-excel.md)| **True** if image files are not generated from drawing objects when you save a document as a Web page. **False** if images are generated. The default value is **False** . Read/write **Boolean** .|
|[ScreenSize](weboptions-screensize-property-excel.md)|Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a Web browser. Can be one of the  **[MsoScreenSize](msoscreensize-enumeration-office.md)** constants. The default constant is **msoScreenSize800x600** . Read/write **MsoScreenSize** .|
|[TargetBrowser](weboptions-targetbrowser-property-excel.md)|Returns or sets an  **[MsoTargetBrowser](msotargetbrowser-enumeration-office.md)** constant indicating the browser version. Read/write.|
|[UseLongFileNames](weboptions-uselongfilenames-property-excel.md)| **True** if long file names are used when you save the document as a Web page. **False** if long file names are not used and the DOS file name format (8.3) is used. The default value is **True** . Read/write **Boolean** .|

