---
title: Application Methods (Access)
ms.prod: ACCESS
ms.assetid: 8e4448f2-d89a-4deb-806d-75159ee61f23
---


# Application Methods (Access)

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AccessError](application-accesserror-method-access.md)|You can use the  **AccessError** method to return the descriptive string associated with a Microsoft Access or DAO error.|
|[AddToFavorites](application-addtofavorites-method-access.md)|The  **AddToFavorites** method adds a hyperlink address to the Favorites folder.|
|[BuildCriteria](application-buildcriteria-method-access.md)|The  **BuildCriteria** method returns a parsed criteria string as it would appear in the query design grid, in Filter By Form or Server Filter By Form mode. For example, you may want to set a form's **Filter** or **[ServerFilter](form-serverfilter-property-access.md)** property based on varying criteria from the user. You can use the **BuildCriteria** method to construct the string expression argument for the **Filter** or **ServerFilter** property. **String**.|
|[CloseCurrentDatabase](application-closecurrentdatabase-method-access.md)|You can use the  **CloseCurrentDatabase** method to close the current database (either a Microsoft Access database or an Access project (.adp) from another application that has opened a database through Automation.|
|[CodeDb](application-codedb-method-access.md)|You can use the  **CodeDb** method in a code module to determine the name of the **Database** object that refers to the database in which code is currently running. Use the **CodeDb** method to access Data Access Objects (DAO) that are part of a library database.|
|[ColumnHistory](application-columnhistory-method-access.md)|Gets the history of values that have been stored in a Memo field.|
|[CompactRepair](application-compactrepair-method-access.md)|Compacts and repairs the specified database or Access project (.adp) file. Returns a  **Boolean**; **True** if the process was successful.|
|[ConvertAccessProject](application-convertaccessproject-method-access.md)|Converts the specified Microsoft Access file from one version to another.|
|[CreateAccessProject](application-createaccessproject-method-access.md)|You can use the  **CreateAccessProject** method to create a new Microsoft Access project (.adp) on disk.|
|[CreateAdditionalData](application-createadditionaldata-method-access.md)|Creates an  **[AdditionalData](additionaldata-object-access.md)** object that can be used to add additional tables and queries to the parent table that is being exported by the **[ExportXML](application-exportxml-method-access.md)** method.|
|[CreateControl](application-createcontrol-method-access.md)|The  **CreateControl** method creates a control on a specified open form. For example, suppose you are building a custom wizard that allows users to easily construct a particular form. You can use the **CreateControl** method in your wizard to add the appropriate controls to the form.|
|[CreateForm](application-createform-method-access.md)|The  **CreateForm** method creates a form and returns a **[Form](form-object-access.md)** object.|
|[CreateGroupLevel](application-creategrouplevel-method-access.md)|You can use the  **CreateGroupLevel** method to specify a field or expression on which to group or sort data in a report. .|
|[CreateReport](application-createreport-method-access.md)|The  **CreateReport** method creates a report and returns a **[Report](report-object-access.md)** object. For example, suppose you are building a custom wizard to create a sales report. You can use the **CreateReport** method in your wizard to create a new report based on a specified report template.|
|[CreateReportControl](application-createreportcontrol-method-access.md)|The  **CreateReportControl** method creates a control on a specified open report. For more information, see the **[CreateControl](application-createcontrol-method-access.md)** method.|
|[CurrentDb](application-currentdb-method-access.md)|The  **CurrentDb** method returns an object variable of type **Database** that represents the database currently open in the Microsoft Access window.|
|[CurrentUser](application-currentuser-method-access.md)|You can use the  **CurrentUser** method to return the name of the current user of the database. .|
|[CurrentWebUser](application-currentwebuser-method-access.md)|Gets information about the current user of a Web database on Microsoft SharePoint Foundation 2010.|
|[CurrentWebUserGroups](application-currentwebusergroups-method-access.md)|Gets the collection of Microsoft SharePoint Foundation 2010 groups of which the user is a member. |
|[DAvg](application-davg-method-access.md)|You can use the  **DAvg** function to calculate the average of a set of values in a specified set of records (a domain).|
|[DCount](application-dcount-method-access.md)|You can use the  **DCount** function to determine the number of records that are in a specified set of records (a domain).|
|[DDEExecute](application-ddeexecute-method-access.md)|You can use the  **DDEExecute** statement to send a command from a client application to a server application over an open dynamic data exchange (DDE) channel.|
|[DDEInitiate](application-ddeinitiate-method-access.md)|You can use the  **DDEInitiate** function to begin a dynamic data exchange (DDE) conversation with another application. The **DDEInitiate** function opens a DDE channel for transfer of data between a DDE server and client application.|
|[DDEPoke](application-ddepoke-method-access.md)|You can use the  **DDEPoke** statement to supply text data from a client application to a server application over an open dynamic data exchange (DDE) channel.|
|[DDERequest](application-dderequest-method-access.md)|You can use the  **DDERequest** function over an open dynamic data exchange (DDE) channel to request an item of information from a DDE server application.|
|[DDETerminate](application-ddeterminate-method-access.md)|You can use the  **DDETerminate** statement to close a specified dynamic data exchange (DDE) channel.|
|[DDETerminateAll](application-ddeterminateall-method-access.md)|You can use the  **DDETerminateAll** statement to close all open dynamic data exchange (DDE) channels.|
|[DefaultWorkspaceClone](application-defaultworkspaceclone-method-access.md)|You can use the  **DefaultWorkspaceClone** method to create a new **Workspace** object without requiring the user to log on again. For example, if you need to conduct two sets of transactions simultaneously in separate workspaces, you can use the **DefaultWorkspaceClone** method to create a second **Workspace** object with the same user name and password without prompting the user for this information again.|
|[DeleteControl](application-deletecontrol-method-access.md)|The  **DeleteControl** method deletes a specified control from a form.|
|[DeleteReportControl](application-deletereportcontrol-method-access.md)|The  **DeleteReportControl** method deletes a specified control from a report.|
|[DFirst](application-dfirst-method-access.md)|You can use the  **DFirst** function to return a random record from a particular field in a table or query when you simply need any value from that field.|
|[DirtyObject](application-dirtyobject-method-access.md)|Marks a form or report as dirty.|
|[DLast](application-dlast-method-access.md)|You can use the  **DLast** function to return a random record from a particular field in a table or query when you simply need any value from that field. .|
|[DLookup](application-dlookup-method-access.md)|You can use the  **DLookup** function to get the value of a particular field from a specified set of records (a domain).|
|[DMax](application-dmax-method-access.md)|You can use  **DMax** function to determine maximum value in a specified set of records (a domain).|
|[DMin](application-dmin-method-access.md)|You can use  **DMin** function to determine minnimum value in a specified set of records (a domain). .|
|[DStDev](application-dstdev-method-access.md)|Estimates the standard deviation across a population sample in a specified set of records (a domain). .|
|[DStDevP](application-dstdevp-method-access.md)|Estimates the standard deviation across a population in a specified set of records (a domain). .|
|[DSum](application-dsum-method-access.md)|You can use the  **DSum** function to calculate the sum of a set of values in a specified set of records (a domain). .|
|[DVar](application-dvar-method-access.md)|Estimates the variance across a sample in a specified set of records (a domain).|
|[DVarP](application-dvarp-method-access.md)|Calculates the variance of a population in a specified set of records (a domain).|
|[Echo](application-echo-method-access.md)|The  **Echo** method specifies whether Microsoft Access repaints the display screen.|
|[EuroConvert](application-euroconvert-method-access.md)|You can use the  **EuroConvert** function to convert a number to euro or from euro to a participating currency. You can also use it to convert a number from one participating currency to another by using the euro as an intermediary (triangulation). The **EuroConvert** function uses fixed conversion rates established by the European Union.|
|[Eval](application-eval-method-access.md)|You can use the  **Eval** function to evaluate an expression that results in a text string or a numeric value.|
|[ExportNavigationPane](application-exportnavigationpane-method-access.md)|Saves the current configuration of the Navigation Pane to an XML file.|
|[ExportXML](application-exportxml-method-access.md)|The  **ExportXML** method allows developers to export XML data, schemas, and presentation information from Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 6.5 or later, or the Microsoft Access database engine.|
|[FollowHyperlink](application-followhyperlink-method-access.md)|The  **FollowHyperlink** method opens the document or Web page specified by a hyperlink address.|
|[GetHiddenAttribute](application-gethiddenattribute-method-access.md)|The  **GetHiddenAttribute** method returns the value of hidden attribute of a Microsoft Access object in the object's **Properties** dialog box, available by selecting the object in the Database window and clicking **Properties** on the **View** menu.|
|[GetOption](application-getoption-method-access.md)|The  **GetOption** method returns the current value of an option in the **Access Options** dialog box, available by clicking the Microsoft Office Button
![File menu button](/images/O12FileMenuButton_ZA10077102.gif) and then clicking **Access Options**.  **Variant**.|
|[GUIDFromString](application-guidfromstring-method-access.md)|The  **GUIDFromString** function converts a string to a GUID, which is an array of type **Byte**.|
|[HtmlEncode](application-htmlencode-method-access.md)|Converts a string to an HTML-encoded string.|
|[hWndAccessApp](application-hwndaccessapp-method-access.md)|You can use the  **hWndAccessApp** method to determine the handle assigned by Microsoft Windows to the main Microsoft Access window.|
|[HyperlinkPart](application-hyperlinkpart-method-access.md)|The  **HyperlinkPart** method returns information about data stored as a Hyperlink data type. .|
|[ImportNavigationPane](application-importnavigationpane-method-access.md)|Loads a saved Navigation Pane configuration from disk.|
|[ImportXML](application-importxml-method-access.md)|The  **ImportXML** method allows developers to import XML data and/or schema information into Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 7.0 or later, or the Microsoft Access database engine.|
|[InstantiateTemplate](application-instantiatetemplate-method-access.md)|Opens a new database and applies the specified template.|
|[IsCurrentWebUserInGroup](application-iscurrentwebuseringroup-method-access.md)|Gets whether or not the current user of a Web databse is a member of the specified Microsoft SharePoint Foundation 2010 group.|
|[LoadCustomUI](application-loadcustomui-method-access.md)|Loads XML markup that represents a customized ribbon.|
|[LoadFromAXL](application-loadfromaxl-method-access.md)|Imports the object defined in an Application XML (AXL) file into the database. |
|[LoadPicture](application-loadpicture-method-access.md)|The  **LoadPicture** method loads a graphic into an ActiveX control.|
|[NewAccessProject](application-newaccessproject-method-access.md)|You can use the  **NewAccessProject** method to create and open a new Microsoft Access project (.adp) as the current Access project in the Microsoft Access window.|
|[NewCurrentDatabase](application-newcurrentdatabase-method-access.md)|Creates a new Microsoft Access database.|
|[Nz](application-nz-method-access.md)|You can use the  **Nz** function to return zero, a zero-length string (" "), or another specified value when a **Variant** is **Null**. For example, you can use this function to convert a **Null** value to another value and prevent it from propagating through an expression.|
|[OpenAccessProject](application-openaccessproject-method-access.md)|You can use the  **OpenAccessProject** method to open an existing Microsoft Access project (.adp) as the current Access project in the Microsoft Access window.|
|[OpenCurrentDatabase](application-opencurrentdatabase-method-access.md)|You can use the  **OpenCurrentDatabase** method to open an existing Microsoft Access database as the current database.|
|[PlainText](application-plaintext-method-access.md)|Strips the rich text formatting from a string and returns an unformatted text string.|
|[Quit](application-quit-method-access.md)|The [Quit](application-quit-method-access.md) method quits Microsoft Access. You can select one of several options for saving a database object before quitting.|
|[RefreshDatabaseWindow](application-refreshdatabasewindow-method-access.md)|The  **RefreshDatabaseWindow** method updates the Database window after a database object has been created, deleted, or renamed.|
|[RefreshTitleBar](application-refreshtitlebar-method-access.md)|The  **RefreshTitleBar** method refreshes the Microsoft Access title bar after the **[AppTitle](apptitle-property.md)** or **[AppIcon](appicon-property.md)** property has been set in Visual Basic.|
|[Run](application-run-method-access.md)|You can use the  **Run** method to carry out a specified Microsoft Access or user-defined **Function** or **Sub** procedure. **Variant**.|
|[RunCommand](application-runcommand-method-access.md)|The  **RunCommand** method runs a built-in command.|
|[SaveAsAXL](application-saveasaxl-method-access.md)|Exports the specified object to an Application XML (AXL) file.|
|[SaveAsTemplate](application-saveastemplate-method-access.md)|Converts an existing Access database file to a database template (*.accdt) format file.|
|[SetDefaultWorkgroupFile](application-setdefaultworkgroupfile-method-access.md)|Sets the default workgroup file to the specified file.|
|[SetHiddenAttribute](application-sethiddenattribute-method-access.md)|The  **SetHiddenAttribute** method sets the hidden attribute of an Access object.|
|[SetOption](application-setoption-method-access.md)|The  **SetOption** method sets the current value of an option in the **Access Options** dialog box.|
|[StringFromGUID](application-stringfromguid-method-access.md)|The  **StringFromGUID** function converts a GUID, which is an array of type **Byte**, to a string.|
|[SysCmd](application-syscmd-method-access.md)|You can use the  **SysCmd** method to, display a progress meter or optional specified text in the status bar, return information about Microsoft Access and its associated files, or return the state of a specified database object (to indicate whether the object is open, is a new object, or has been changed but not saved). **Variant**.|
|[TransformXML](application-transformxml-method-access.md)|Applies an Extensible Stylesheet Language (XSL) stylesheet to an XML data file and writes the resulting XML to an XML data file.|

