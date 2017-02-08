---
title: Command (ADO for Visual C++ Syntax)
ms.prod: ACCESS
ms.assetid: a397daf5-2bcd-6c1a-3fb6-667c1309d0e3
---


# Command (ADO for Visual C++ Syntax)

 **Last modified:** December 30, 2015

**Applies to:** Access 2013 | Access 2016

 **Methods**

[Cancel](http://msdn.microsoft.com/library/cancel-method-ado%28Office.15%29.aspx)(void)[CreateParameter](http://msdn.microsoft.com/library/createparameter-method-ado%28Office.15%29.aspx)(BSTR  _Name,_ DataTypeEnum _Type,_ ParameterDirectionEnum _Direction,_ long _Size,_ VARIANT _Value,_ _ADOParameter ** _ppiprm_ )[Execute](execute-method-ado-command.md)(VARIANT * _RecordsAffected,_ VARIANT * _Parameters,_ long _Options,_ _ADORecordset ** _ppirs_ )
 **Properties**
[get_ActiveConnection](http://msdn.microsoft.com/library/activeconnection-property-ado%28Office.15%29.aspx)(_ADOConnection ** _ppvObject_ ) **put_ActiveConnection** (VARIANT _vConn_ ) **putref_ActiveConnection** (_ADOConnection * _pCon_ )[get_CommandText](http://msdn.microsoft.com/library/commandtext-property-ado%28Office.15%29.aspx)(BSTR * _pbstr_ ) **put_CommandText** (BSTR _bstr_ )[get_CommandTimeout](http://msdn.microsoft.com/library/commandtimeout-property-ado%28Office.15%29.aspx)(LONG * _pl_ ) **put_CommandTimeout** (LONG _Timeout_ )[get_CommandType](http://msdn.microsoft.com/library/commandtype-property-ado%28Office.15%29.aspx)(CommandTypeEnum * _plCmdType_ ) **put_CommandType** (CommandTypeEnum _lCmdType_ )[get_Name](http://msdn.microsoft.com/library/name-property-ado%28Office.15%29.aspx)(BSTR * _pbstrName_ ) **put_Name** (BSTR _bstrName_ )[get_Prepared](http://msdn.microsoft.com/library/prepared-property-ado%28Office.15%29.aspx)(VARIANT_BOOL * _pfPrepared_ ) **put_Prepared** (VARIANT_BOOL _fPrepared_ )[get_State](http://msdn.microsoft.com/library/state-property-ado%28Office.15%29.aspx)(LONG * _plObjState_ )[get_Parameters](http://msdn.microsoft.com/library/parameters-collection-ado%28Office.15%29.aspx)(ADOParameters ** _ppvObject_ )
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

