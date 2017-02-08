---
title: TextConnection Properties (Excel)
ms.prod: EXCEL
ms.assetid: 91d71e58-2760-4969-5948-71a1926569eb
---


# TextConnection Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](textconnection-application-property-excel.md)|Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.|
|[Connection](textconnection-connection-property-excel.md)|Returns or sets a string that contains text file names that enable Microsoft Excel to connect to text data sources.  **Variant** Read/Write|
|[Creator](textconnection-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Parent](textconnection-parent-property-excel.md)|Returns an  **Object** that represents the parent object of the specified[TextConnection Object (Excel)](textconnection-object-excel.md) object. Read-only.|
|[TextFileColumnDataTypes](textconnection-textfilecolumndatatypes-property-excel.md)|Returns or sets an ordered array of constants that specify the data types applied to the corresponding columns in the text file that you're importing into a query table. The default constant for each column is  **xlGeneral** . **Variant** . Read/Write|
|[TextFileCommaDelimiter](textconnection-textfilecommadelimiter-property-excel.md)| **True** if the comma is the delimiter when you import a text file into a query table. **False** if you want to use some other character as the delimiter. The default value is **False** . Read/Write **Boolean** .|
|[TextFileConsecutiveDelimiter](textconnection-textfileconsecutivedelimiter-property-excel.md)| **True** if consecutive delimiters are treated as a single delimiter when you import a text file into a query table. The default value is **False** . **Boolean** Read/Write|
|[TextFileDecimalSeparator](textconnection-textfiledecimalseparator-property-excel.md)|Returns or sets the decimal separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system decimal separator character. Read/Write  **String** .|
|[TextFileFixedColumnWidths](textconnection-textfilefixedcolumnwidths-property-excel.md)|Returns or sets an array of integers that correspond to the widths of the columns (in characters) in the text file that you're importing into a query table. Valid widths are from 1 through 32767 characters. Read/Write  **Variant** .|
|[TextFileHeaderRow](textconnection-textfileheaderrow-property-excel.md)|Returns or sets value that specifies whether or not the first row (from the starting row) should be treated as a header row.  **Boolean** Read/Write|
|[TextFileOtherDelimiter](textconnection-textfileotherdelimiter-property-excel.md)|Returns or sets the character used as the delimiter when you import a text file into a query table. The default value is  **null** . Read/Write **String** .|
|[TextFileParseType](textconnection-textfileparsetype-property-excel.md)|Returns or sets the column format for the data in the text file that you're importing into a query table. Read/Write [XlTextParsingType Enumeration (Excel)](xltextparsingtype-enumeration-excel.md)|
|[TextFilePlatform](textconnection-textfileplatform-property-excel.md)|Returns or sets the origin of the text file you're importing into the query table. This property determines which code page is used during the data import. Read/write [XlPlatform Enumeration (Excel)](xlplatform-enumeration-excel.md)|
|[TextFilePromptOnRefresh](textconnection-textfilepromptonrefresh-property-excel.md)| **True** if you want to specify the name of the imported text file each time the query table is refreshed. The **Import Text File** dialog box allows you to specify the path and file name. The default value is **False** . Read/Write **Boolean** .|
|[TextFileSemicolonDelimiter](textconnection-textfilesemicolondelimiter-property-excel.md)| **True** if the semicolon is the delimiter when you import a text file into a query table, and if the value of the[TextConnection.TextFileParseType Property (Excel)](textconnection-textfileparsetype-property-excel.md) property is **xlDelimited** . The default value is **False** . Read/Write **Boolean** .|
|[TextFileSpaceDelimiter](textconnection-textfilespacedelimiter-property-excel.md)| **True** if the space character is the delimiter when you import a text file into a query table. The default value is **False** . Read/Write **Boolean**|
|[TextFileStartRow](textconnection-textfilestartrow-property-excel.md)|Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1. Read/Write  **Long** .|
|[TextFileTabDelimiter](textconnection-textfiletabdelimiter-property-excel.md)| **True** if the tab character is the delimiter when you import a text file into a query table. The default value is **False** . Read/Write **Boolean**|
|[TextFileTextQualifier](textconnection-textfiletextqualifier-property-excel.md)|Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format. Read/Write [XlTextQualifier Enumeration (Excel)](xltextqualifier-enumeration-excel.md)|
|[TextFileThousandsSeparator](textconnection-textfilethousandsseparator-property-excel.md)|Returns or sets the thousands separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system thousands separator character. Read/Write  **String**|
|[TextFileTrailingMinusNumbers](textconnection-textfiletrailingminusnumbers-property-excel.md)| **True** for Microsoft Excel to treat numbers imported as text that begin with a " **-** " symbol as a negative symbol. **False** for Excel to treat numbers imported as text that begin with a " **-** " symbol as text. Read/Write **Boolean**|
|[TextFileVisualLayout](textconnection-textfilevisuallayout-property-excel.md)|Returns or sets a [XlTextVisualLayoutType Enumeration (Excel)](xltextvisuallayouttype-enumeration-excel.md) enumeration that indicates whether the visual layout of the text being imported is left-to-right or right-to-left. Read/Write|

