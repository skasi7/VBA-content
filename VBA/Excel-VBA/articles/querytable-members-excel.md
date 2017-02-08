---
title: QueryTable Members (Excel)
ms.prod: EXCEL
ms.assetid: 9a61f024-c1dc-c11b-942f-ff2a6617bdc4
---


# QueryTable Members (Excel)
Represents a worksheet table built from data returned from an external data source, such as an SQL server or a Microsoft Access database.

Represents a worksheet table built from data returned from an external data source, such as an SQL server or a Microsoft Access database.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterRefresh](querytable-afterrefresh-event-excel.md)|Occurs after a query is completed or canceled.|
|[BeforeRefresh](querytable-beforerefresh-event-excel.md)|Occurs before any refreshes of the query table. This includes refreshes resulting from calling the  **Refresh** method, from the user's actions in the product, and from opening the workbook containing the query table.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[CancelRefresh](querytable-cancelrefresh-method-excel.md)|Cancels all background queries for the specified query table. Use the  **[Refreshing](querytable-refreshing-property-excel.md)** property to determine whether a background query is currently in progress.|
|[Delete](querytable-delete-method-excel.md)|Deletes the object.|
|[Refresh](querytable-refresh-method-excel.md)|Updates an external data range ( **[QueryTable](querytable-object-excel.md)** ).|
|[ResetTimer](querytable-resettimer-method-excel.md)|Resets the refresh timer for the specified query table or PivotTable report to the last interval you set using the  **[RefreshPeriod](querytable-refreshperiod-property-excel.md)** property.|
|[SaveAsODC](querytable-saveasodc-method-excel.md)|Saves the QueryTable cache source as an Microsoft Office Data Connection file.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AdjustColumnWidth](querytable-adjustcolumnwidth-property-excel.md)| **True** if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. **False** if the column widths are not automatically adjusted with each refresh. The default value is **True** . Read/write **Boolean** .|
|[Application](querytable-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[BackgroundQuery](querytable-backgroundquery-property-excel.md)| **True** if queries for the query table are performed asynchronously (in the background). Read/write **Boolean** .|
|[CommandText](querytable-commandtext-property-excel.md)|Returns or sets the command string for the specified data source. Read/write  **Variant** .|
|[CommandType](querytable-commandtype-property-excel.md)|Returns or sets one of the  **[XlCmdType](xlcmdtype-enumeration-excel.md)** constants listed in the following table in the remarks section. The constant that is returned or set describes the value of the **[CommandText](querytable-commandtext-property-excel.md)** property. The default value is **xlCmdSQL** . Read/write **XlCmdType** .|
|[Connection](querytable-connection-property-excel.md)|Returns or sets a string that contains one of the following: OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source; ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; the path to and file name of a text file, or the path to and file name of a file that specifies a database or Web query. Read/write  **Variant** .|
|[Creator](querytable-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Destination](querytable-destination-property-excel.md)|Returns the cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains the  **QueryTable** object. Read-only **[Range](range-object-excel.md)** .|
|[EditWebPage](querytable-editwebpage-property-excel.md)|Returns or sets the web page Uniform Resource Locator (URL) for a web query. Read/write  **Variant** .|
|[EnableEditing](querytable-enableediting-property-excel.md)| **True** if the user can edit the specified query table. **False** if the user can only refresh the query table. Read/write **Boolean** .|
|[EnableRefresh](querytable-enablerefresh-property-excel.md)| **True** if the PivotTable cache or query table can be refreshed by the user. The default value is **True** . Read/write **Boolean** .|
|[FetchedRowOverflow](querytable-fetchedrowoverflow-property-excel.md)| **True** if the number of rows returned by the last use of the **[Refresh](querytable-refresh-method-excel.md)** method is greater than the number of rows available on the worksheet. Read-only **Boolean** .|
|[FieldNames](querytable-fieldnames-property-excel.md)| **True** if field names from the data source appear as column headings for the returned data. The default value is **True** . Read/write **Boolean** .|
|[FillAdjacentFormulas](querytable-filladjacentformulas-property-excel.md)| **True** if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed. Read/write **Boolean** .|
|[ListObject](querytable-listobject-property-excel.md)|Returns a  **[ListObject](listobject-object-excel.md)** object for the **[QueryTable](querytable-object-excel.md)** object. Read-only **ListObject** object.|
|[MaintainConnection](querytable-maintainconnection-property-excel.md)| **True** if the connection to the specified data source is maintained after the refresh and until the workbook is closed. The default value is **True** . Read/write **Boolean** .|
|[Name](querytable-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[Parameters](querytable-parameters-property-excel.md)|Returns a  **[Parameters](parameters-object-excel.md)** collection that represents the query table parameters. Read-only.|
|[Parent](querytable-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PostText](querytable-posttext-property-excel.md)|Returns or sets the string used with the post method of inputting data into a Web server to return data from a Web query. Read/write  **String** .|
|[PreserveColumnInfo](querytable-preservecolumninfo-property-excel.md)| **True** if column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value is **True** . Read/write **Boolean** .|
|[PreserveFormatting](querytable-preserveformatting-property-excel.md)| **True** if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property is **False** if the last AutoFormat applied to the query table is applied to new rows of data. The default value is **True** .|
|[QueryType](querytable-querytype-property-excel.md)|Indicates the type of query used by Microsoft Excel to populate the query table. Read-only  **[XlQueryType](xlquerytype-enumeration-excel.md)** .|
|[Recordset](querytable-recordset-property-excel.md)|Returns or sets a  **Recordset** object that's used as the data source for the specified query table. Read/write.|
|[Refreshing](querytable-refreshing-property-excel.md)| **True** if there is a background query in progress for the specified query table. Read only **Boolean** .|
|[RefreshOnFileOpen](querytable-refreshonfileopen-property-excel.md)| **True** if the PivotTable cache or query table is automatically updated each time the workbook is opened. The default value is **False** . Read/write **Boolean** .|
|[RefreshPeriod](querytable-refreshperiod-property-excel.md)|Returns or sets the number of minutes between refreshes. Read/write  **Long** .|
|[RefreshStyle](querytable-refreshstyle-property-excel.md)|Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query. Read/write  **[XlCellInsertionMode](xlcellinsertionmode-enumeration-excel.md)** .|
|[ResultRange](querytable-resultrange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the area of the worksheet occupied by the specified query table. Read-only.|
|[RobustConnect](querytable-robustconnect-property-excel.md)|Returns or sets how the query table connects to its data source. Read/write  **[XlRobustConnect](xlrobustconnect-enumeration-excel.md)** .|
|[RowNumbers](querytable-rownumbers-property-excel.md)| **True** if row numbers are added as the first column of the specified query table. Read/write **Boolean** .|
|[SaveData](querytable-savedata-property-excel.md)| **True** if data for the QueryTable report is saved with the workbook. **False** if only the report definition is saved. Read/write **Boolean** .|
|[SavePassword](querytable-savepassword-property-excel.md)| **True** if password information in an ODBC connection string is saved with the specified query. **False** if the password is removed. Read/write **Boolean** .|
|[Sort](querytable-sort-property-excel.md)|Returns the sort criteria for the query table range. Read-only.|
|[SourceConnectionFile](querytable-sourceconnectionfile-property-excel.md)|Returns or sets a  **String** indicating the Microsoft Office Data Connection file or similar file that was used to create the QueryTable. Read/write.|
|[SourceDataFile](querytable-sourcedatafile-property-excel.md)|Returns or sets a  **String** value that indicates the source data file for a query table.|
|[TextFileColumnDataTypes](querytable-textfilecolumndatatypes-property-excel.md)|Returns or sets an ordered array of constants that specify the data types applied to the corresponding columns in the text file that you are importing into a query table. The default constant for each column is  **xlGeneral** . Read/write **Variant** .|
|[TextFileCommaDelimiter](querytable-textfilecommadelimiter-property-excel.md)| **True** if the comma is the delimiter when you import a text file into a query table. **False** if you want to use some other character as the delimiter. The default value is **False** . Read/write **Boolean** .|
|[TextFileConsecutiveDelimiter](querytable-textfileconsecutivedelimiter-property-excel.md)| **True** if consecutive delimiters are treated as a single delimiter when you import a text file into a query table. The default value is **False** . Read/write **Boolean** .|
|[TextFileDecimalSeparator](querytable-textfiledecimalseparator-property-excel.md)|Returns or sets the decimal separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system decimal separator character. Read/write  **String** .|
|[TextFileFixedColumnWidths](querytable-textfilefixedcolumnwidths-property-excel.md)|Returns or sets an array of integers that correspond to the widths of the columns (in characters) in the text file that you're importing into a query table. Valid widths are from 1 through 32767 characters. Read/write  **Variant** .|
|[TextFileOtherDelimiter](querytable-textfileotherdelimiter-property-excel.md)|Returns or sets the character used as the delimiter when you import a text file into a query table. The default value is  **null** . Read/write **String** .|
|[TextFileParseType](querytable-textfileparsetype-property-excel.md)|Returns or sets the column format for the data in the text file that you're importing into a query table. Read/write  **[XlTextParsingType](xltextparsingtype-enumeration-excel.md)** .|
|[TextFilePlatform](querytable-textfileplatform-property-excel.md)|Returns or sets the origin of the text file you're importing into the query table. This property determines which code page is used during the data import. Read/write  **[XlPlatform](xlplatform-enumeration-excel.md)** .|
|[TextFilePromptOnRefresh](querytable-textfilepromptonrefresh-property-excel.md)| **True** if you want to specify the name of the imported text file each time the query table is refreshed. The **Import Text File** dialog box allows you to specify the path and file name. The default value is **False** . Read/write **Boolean** .|
|[TextFileSemicolonDelimiter](querytable-textfilesemicolondelimiter-property-excel.md)| **True** if the semicolon is the delimiter when you import a text file into a query table, and if the value of the **[TextFileParseType](querytable-textfileparsetype-property-excel.md)** property is **xlDelimited** . The default value is **False** . Read/write **Boolean** .|
|[TextFileSpaceDelimiter](querytable-textfilespacedelimiter-property-excel.md)| **True** if the space character is the delimiter when you import a text file into a query table. The default value is **False** . Read/write **Boolean** .|
|[TextFileStartRow](querytable-textfilestartrow-property-excel.md)|Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1. Read/write  **Long** .|
|[TextFileTabDelimiter](querytable-textfiletabdelimiter-property-excel.md)| **True** if the tab character is the delimiter when you import a text file into a query table. The default value is **False** . Read/write **Boolean** .|
|[TextFileTextQualifier](querytable-textfiletextqualifier-property-excel.md)|Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format. Read/write  **[XlTextQualifier](xltextqualifier-enumeration-excel.md)** .|
|[TextFileThousandsSeparator](querytable-textfilethousandsseparator-property-excel.md)|Returns or sets the thousands separator character thatMicrosoft Excel uses when you import a text file into a query table. The default is the system thousands separator character. Read/write  **String** .|
|[TextFileTrailingMinusNumbers](querytable-textfiletrailingminusnumbers-property-excel.md)| **True** for Microsoft Excel to treat numbers imported as text that begin with a "-" symbol as a negative symbol. **False** for Excel to treat numbers imported as text that begin with a "-" symbol as text. Read/write **Boolean** .|
|[TextFileVisualLayout](querytable-textfilevisuallayout-property-excel.md)|Returns or sets a  **[XlTextVisualLayoutType](xltextvisuallayouttype-enumeration-excel.md)** enumeration that indicates whether the visual layout of the text being imported is left-to-right or right-to-left.|
|[WebConsecutiveDelimitersAsOne](querytable-webconsecutivedelimitersasone-property-excel.md)| **True** if consecutive delimiters are treated as a single delimiter when you import data from HTML <PRE> tags in a Web page into a query table, and if the data is to be parsed into columns. **False** if you want to treat consecutive delimiters as multiple delimiters. The default value is **True** . Read/write **Boolean** .|
|[WebDisableDateRecognition](querytable-webdisabledaterecognition-property-excel.md)| **True** if data that resembles dates is parsed as text when you import a Web page into a query table. **False** if date recognition is used. The default value is **False** . Read/write **Boolean** .|
|[WebDisableRedirections](querytable-webdisableredirections-property-excel.md)| **True** if Web query redirections are disabled for a **QueryTable** object. The default value is **False** . Read/write **Boolean** .|
|[WebFormatting](querytable-webformatting-property-excel.md)|Returns or sets a value that determines how much formatting from a Web page, if any, is applied when you import the page into a query table. Read/write  **[XlWebFormatting](xlwebformatting-enumeration-excel.md)** .|
|[WebPreFormattedTextToColumns](querytable-webpreformattedtexttocolumns-property-excel.md)|Returns or sets whether data contained within HTML <PRE> tags in the Web page is parsed into columns when you import the page into a query table. The default is  **True** . Read/write **Boolean** .|
|[WebSelectionType](querytable-webselectiontype-property-excel.md)|Returns or sets a value that determines whether an entire Web page, all tables on the Web page, or only specific tables on the Web page are imported into a query table. Read/write  **[XlWebSelectionType](xlwebselectiontype-enumeration-excel.md)** .|
|[WebSingleBlockTextImport](querytable-websingleblocktextimport-property-excel.md)| **True** if data from the HTML <PRE> tags in the specified Web page is processed all at once when you import the page into a query table. **False** if the data is imported in blocks of contiguous rows so that header rows will be recognized as such. The default value is **False** . Read/write **Boolean** .|
|[WebTables](querytable-webtables-property-excel.md)|Returns or sets a comma-delimited list of table names or table index numbers when you import a Web page into a query table. Read/write  **String** .|
|[WorkbookConnection](querytable-workbookconnection-property-excel.md)|Returns the  **[WorkbookConnection](workbookconnection-object-excel.md)** object that the query table uses. Read-only.|

