---
title: LanguageSettings Object (Office)
keywords: vbaof11.chm231000
f1_keywords:
- vbaof11.chm231000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.LanguageSettings
ms.assetid: 936f7d61-87e5-e153-08d4-f8c5c8ef0710
---


# LanguageSettings Object (Office)

Returns information about the language settings in a Microsoft Office application.


## Remarks

Use Application.LanguageSettings.LanguageID( _MsoAppLanguageID_ ), where[MsoAppLanguageID](http://msdn.microsoft.com/library/msoapplanguageid-enumeration-office%28Office.15%29.aspx) is a constant used to return locale identifier (LCID) information to the specified application.


## Example

The following example returns the install language, user interface language, and Help language LCIDs in a message box.


```
MsgBox "The following locale IDs are registered " &amp; _ 
 "for this application: Install Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDInstall) &amp; _ 
 " User Interface Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDUI) &amp; _ 
 " Help Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDHelp)
```

Use  **Application.LanguageSettings.LanguagePreferredForEditing** to determine which LCIDs are registered as preferred editing languages for the application, as in the following example.




```
If Application.LanguageSettings. _ 
 LanguagePreferredForEditing(msoLanguageIDEnglishUS) Then 
 MsgBox "U.S. English is one of the chosen editing languagess." 
End If
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/languagesettings-application-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/languagesettings-creator-property-office%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/languagesettings-languageid-property-office%28Office.15%29.aspx)|
|[LanguagePreferredForEditing](http://msdn.microsoft.com/library/languagesettings-languagepreferredforediting-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/languagesettings-parent-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[LanguageSettings Object Members](http://msdn.microsoft.com/library/languagesettings-members-office%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
