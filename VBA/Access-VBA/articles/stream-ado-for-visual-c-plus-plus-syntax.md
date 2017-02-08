---
title: Stream (ADO for Visual C++ Syntax)
ms.prod: ACCESS
ms.assetid: e1482f15-9ef6-9485-06c2-1123762afc9f
---


# Stream (ADO for Visual C++ Syntax)

 **Last modified:** December 30, 2015

**Applies to:** Access 2013 | Access 2016

 **Methods**

[Cancel](http://msdn.microsoft.com/library/cancel-method-ado%28Office.15%29.aspx)(void)[Close](http://msdn.microsoft.com/library/close-method-ado%28Office.15%29.aspx)(void)[CopyTo](http://msdn.microsoft.com/library/copyto-method-ado%28Office.15%29.aspx)(_ADOStream  _*DestStream,_ LONG _CharNumber_ = -1)[Flush](http://msdn.microsoft.com/library/flush-method-ado%28Office.15%29.aspx)(void)[LoadFromFile](http://msdn.microsoft.com/library/loadfromfile-method-ado%28Office.15%29.aspx)(BSTR _ FileName_ )[Open](http://msdn.microsoft.com/library/open-method-ado-stream%28Office.15%29.aspx)(VARIANT _ Source,_ ConnectModeEnum _ Mode,_ StreamOpenOptionsEnum _ Options,_ BSTR _ UserName,_ BSTR _ Password_ )[Read](http://msdn.microsoft.com/library/read-method-ado%28Office.15%29.aspx)(long _ NumBytes,_ VARIANT _*pVal_ )[ReadText](http://msdn.microsoft.com/library/readtext-method-ado%28Office.15%29.aspx)(long _ NumChars,_ BSTR _*pbstr_ )[SaveToFile](http://msdn.microsoft.com/library/savetofile-method-ado%28Office.15%29.aspx)(BSTR _ FileName,_ SaveOptionsEnum _Options_ =adSaveCreateNotExist)[SetEOS](http://msdn.microsoft.com/library/seteos-method-ado%28Office.15%29.aspx)(void)[SkipLine](http://msdn.microsoft.com/library/skipline-method-ado%28Office.15%29.aspx)(void)[Write](http://msdn.microsoft.com/library/write-method-ado%28Office.15%29.aspx)(VARIANT _ Buffer_ )[WriteText](http://msdn.microsoft.com/library/writetext-method-ado%28Office.15%29.aspx)(BSTR _ Data,_ StreamWriteEnum _Options_ =adWriteChar)
 **Properties**
[get_Charset](http://msdn.microsoft.com/library/charset-property-ado%28Office.15%29.aspx)(BSTR  _*pbstrCharset_ ) **put_Charset** (BSTR _ Charset_ )[get_EOS](http://msdn.microsoft.com/library/eos-property-ado%28Office.15%29.aspx)(VARIANT_BOOL  _*pEOS_ )[get_LineSeparator](http://msdn.microsoft.com/library/lineseparator-property-ado%28Office.15%29.aspx)(LineSeparatorEnum  _*pLS_ ) **put_LineSeparator** (LineSeparatorEnum _ LineSeparator_ )[get_Mode](http://msdn.microsoft.com/library/mode-property-ado%28Office.15%29.aspx)(ConnectModeEnum  _*pMode_ ) **put_Mode** (ConnectModeEnum _ Mode_ )[get_Position](http://msdn.microsoft.com/library/position-property-ado%28Office.15%29.aspx)(LONG  _*pPos_ ) **put_Position** (LONG _ Position_ )[get_Size](size-property-ado-stream.md)(LONG  _*pSize_ )[get_State](http://msdn.microsoft.com/library/state-property-ado%28Office.15%29.aspx)(ObjectStateEnum  _*pState_ )[get_Type](http://msdn.microsoft.com/library/type-property-ado-stream%28Office.15%29.aspx)(StreamTypeEnum  _*pType_ ) **put_Type** (StreamTypeEnum _ Type_ )
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

