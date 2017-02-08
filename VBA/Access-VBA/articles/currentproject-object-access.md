---
title: CurrentProject Object (Access)
keywords: vbaac10.chm12739
f1_keywords:
- vbaac10.chm12739
ms.prod: ACCESS
api_name:
- Access.CurrentProject
ms.assetid: e6baae73-1eeb-b48f-d35e-b3e921378561
---


# CurrentProject Object (Access)

The  **CurrentProject** object refers to the project for the current Microsoft Access project (.adp) or Access database.


## Remarks

The  **CurrentProject** object has several collections that contain specific **[AccessObject](accessobject-object-access.md)** objects within the current database. The following table lists the name of each collection and the types of objects it contains.



|**Collections**|**Object type**|
|:-----|:-----|
|**[AllForms](allforms-object-access.md)**|All forms|
|**[AllReports](http://msdn.microsoft.com/library/allreports-object-access%28Office.15%29.aspx)**|All reports|
|**[AllMacros](http://msdn.microsoft.com/library/allmacros-object-access%28Office.15%29.aspx)**|All macros|
|**[AllModules](http://msdn.microsoft.com/library/allmodules-object-access%28Office.15%29.aspx)**|All modules|

 **Note**  The collections in the preceding table contain all of the respective objects in the database regardless if they are opened or closed.

For example, an  **AccessObject** object representing a form is a member of the **AllForms** collection, which is a collection of **AccessObject** objects within the current database. Within the **AllForms** collection, individual members of the collection are indexed beginning with zero. You can refer to an individual **AccessObject** object in the **AllForms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllForms** collection, it's better to refer to it by name because a item's collection index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**AllForms** ! _formname_|AllForms!OrderForm|
|**AllForms** ![ _form name_]|AllForms![Order Form]|
|**AllForms** (" _formname_")|AllForms("OrderForm")|
|**AllForms** ( _formname_)|AllForms(0)|

## Example

The following example prints some current property settings of the  **CurrentProject** object and then sets an option to display hidden objects within the application:


```
Sub ApplicationInformation() 
 ' Print name and type of current object. 
 Debug.Print Application.CurrentProject.FullName 
 Debug.Print Application.CurrentProject.ProjectType 
 ' Set Hidden Objects option under Show on View Tab 
 'of the Options dialog box. 
 Application.SetOption "Show Hidden Objects", True 
End Sub
```

The next example shows how to use the CurrentProject object using Automation from another Microsoft Office application. First, from the other application, create a reference to Microsoft Access by clicking  **References** on the **Tools** menu in the Module window. Select the check box next to **Microsoft Access Object Library**. Then enter the following code in a Visual Basic module within that application and call the GetAccessData procedure.

The example passes a database name and report name to a procedure that creates a new instance of the  **Application** class, opens the database, and verifies that the specified report exists using the **CurrentProject** object and **AllReports** collection.




```
Sub GetAccessData() 
' Declare object variable in declarations section of a module 
 Dim appAccess As Access.Application 
 Dim strDB As String 
 Dim strReportName As String 
 
 strDB = "C:\Program Files\Microsoft "_ 
 &amp; "Office\Office11\Samples\Northwind.mdb" 
 strReportName = InputBox("Enter name of report to be verified", _ 
 "Report Verification") 
 VerifyAccessReport strDB, strReportName 
End Sub 
 
Sub VerifyAccessReport(strDB As String, _ 
 strReportName As String) 
 ' Return reference to Microsoft Access 
 ' Application object. 
 Set appAccess = New Access.Application 
 ' Open database in Microsoft Access. 
 appAccess.OpenCurrentDatabase strDB 
 ' Verify report exists. 
 On Error Goto ErrorHandler 
 appAccess.CurrentProject.AllReports(strReportName) 
 MsgBox "Report " &amp; strReportName &amp; _ 
 " verified within Northwind database." 
 appAccess.CloseCurrentDatabase 
 Set appAccess = Nothing 
Exit Sub 
ErrorHandler: 
 MsgBox "Report " &amp; strReportName &amp; _ 
 " does not exist within Northwind database." 
 appAccess.CloseCurrentDatabase 
 Set appAccess = Nothing 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddSharedImage](http://msdn.microsoft.com/library/currentproject-addsharedimage-method-access%28Office.15%29.aspx)|
|[CloseConnection](http://msdn.microsoft.com/library/currentproject-closeconnection-method-access%28Office.15%29.aspx)|
|[OpenConnection](http://msdn.microsoft.com/library/currentproject-openconnection-method-access%28Office.15%29.aspx)|
|[UpdateDependencyInfo](http://msdn.microsoft.com/library/currentproject-updatedependencyinfo-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AccessConnection](http://msdn.microsoft.com/library/currentproject-accessconnection-property-access%28Office.15%29.aspx)|
|[AllForms](http://msdn.microsoft.com/library/currentproject-allforms-property-access%28Office.15%29.aspx)|
|[AllMacros](http://msdn.microsoft.com/library/currentproject-allmacros-property-access%28Office.15%29.aspx)|
|[AllModules](http://msdn.microsoft.com/library/currentproject-allmodules-property-access%28Office.15%29.aspx)|
|[AllReports](http://msdn.microsoft.com/library/currentproject-allreports-property-access%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/currentproject-application-property-access%28Office.15%29.aspx)|
|[BaseConnectionString](http://msdn.microsoft.com/library/currentproject-baseconnectionstring-property-access%28Office.15%29.aspx)|
|[Connection](http://msdn.microsoft.com/library/currentproject-connection-property-access%28Office.15%29.aspx)|
|[FileFormat](http://msdn.microsoft.com/library/currentproject-fileformat-property-access%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/currentproject-fullname-property-access%28Office.15%29.aspx)|
|[ImportExportSpecifications](http://msdn.microsoft.com/library/currentproject-importexportspecifications-property-access%28Office.15%29.aspx)|
|[IsConnected](http://msdn.microsoft.com/library/currentproject-isconnected-property-access%28Office.15%29.aspx)|
|[IsTrusted](http://msdn.microsoft.com/library/currentproject-istrusted-property-access%28Office.15%29.aspx)|
|[IsWeb](http://msdn.microsoft.com/library/currentproject-isweb-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/currentproject-name-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/currentproject-parent-property-access%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/currentproject-path-property-access%28Office.15%29.aspx)|
|[ProjectType](http://msdn.microsoft.com/library/currentproject-projecttype-property-access%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/currentproject-properties-property-access%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/currentproject-removepersonalinformation-property-access%28Office.15%29.aspx)|
|[Resources](http://msdn.microsoft.com/library/currentproject-resources-property-access%28Office.15%29.aspx)|
|[WebSite](http://msdn.microsoft.com/library/currentproject-website-property-access%28Office.15%29.aspx)|
|[IsSQLBackend](http://msdn.microsoft.com/library/currentproject-issqlbackend-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[CurrentProject Object Members](http://msdn.microsoft.com/library/currentproject-members-access%28Office.15%29.aspx)
