---
title: Execute Method (ADO Connection)
ms.prod: ACCESS
ms.assetid: af190bd9-7167-df59-29ca-a9a86c4957fd
---


# Execute Method (ADO Connection)

 **Last modified:** December 30, 2015

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection1)
[Return Value](#sectionSection2)
[Parameters](#sectionSection3)
[Remarks](#sectionSection4)



Executes the specified query, SQL statement, stored procedure, or provider-specific text.

## Syntax
<a name="sectionSection1"> </a>

For a non-row-returning command string:

For a row-returning command string:


## Return Value
<a name="sectionSection2"> </a>

Returns a [Recordset](http://msdn.microsoft.com/library/recordset-object-ado%28Office.15%29.aspx) object reference.


## Parameters
<a name="sectionSection3"> </a>


-  _CommandText_
    
- A  **String** value that contains the SQL statement, stored procedure, a URL, or provider-specific text to execute. Optionally, table names can be used but only if the provider is SQL aware. For example if a table name of "Customers" is used, ADO will automatically prepend the standard SQL Select syntax to form and pass "SELECT * FROM Customers" as a T-SQL statement to the provider.
    
-  _RecordsAffected_
    
- Optional. A  **Long** variable to which the provider returns the number of records that the operation affected.
    
-  _Options_
    
- Optional. A  **Long** value that indicates how the provider should evaluate the _CommandText_ argument. Can be a bitmask of one or more[CommandTypeEnum](http://msdn.microsoft.com/library/commandtypeenum%28Office.15%29.aspx) or[ExecuteOptionEnum](http://msdn.microsoft.com/library/executeoptionenum%28Office.15%29.aspx) values.
    

 **Note**  Use the  **ExecuteOptionEnum** value **adExecuteNoRecords** to improve performance by minimizing internal processing.

Do not use the  **CommandTypeEnum** values of **adCmdFile** or **adCmdTableDirect** with **Execute**. These values can only be used as options with the[Open](http://msdn.microsoft.com/library/open-method-ado-recordset%28Office.15%29.aspx) and[Requery](http://msdn.microsoft.com/library/requery-method-ado%28Office.15%29.aspx) methods of a **Recordset**.


## Remarks
<a name="sectionSection4"> </a>

Using the  **Execute** method on a[Connection](http://msdn.microsoft.com/library/connection-object-ado%28Office.15%29.aspx) object executes whatever query you pass to the method in the _CommandText_ argument on the specified connection. If the _CommandText_ argument specifies a row-returning query, any results that the execution generates are stored in a new **Recordset** object. If the command is not intended to return results (for example, an SQL UPDATE query) the provider returns **Nothing** as long as the option **adExecuteNoRecords** is specified; otherwise Execute returns a closed **Recordset**.

The returned  **Recordset** object is always a read-only, forward-only cursor. If you need a **Recordset** object with more functionality, first create a **Recordset** object with the desired property settings, then use the **Recordset** object's[Open](http://msdn.microsoft.com/library/open-method-ado-recordset%28Office.15%29.aspx) method to execute the query and return the desired cursor type.

The contents of the  _CommandText_ argument are specific to the provider and can be standard SQL syntax or any special command format that the provider supports.

An [ExecuteComplete](http://msdn.microsoft.com/library/executecomplete-event-ado%28Office.15%29.aspx) event will be issued when this operation concludes.


 **Note**  URLs using the http scheme will automatically invoke the [Microsoft OLE DB Provider for Internet Publishing](http://msdn.microsoft.com/library/microsoft-ole-db-provider-for-internet-publishing%28Office.15%29.aspx). For more information, see [Absolute and Relative URLs](http://msdn.microsoft.com/library/absolute-and-relative-urls%28Office.15%29.aspx).

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

