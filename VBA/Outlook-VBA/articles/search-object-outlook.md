---
title: Search Object (Outlook)
keywords: vbaol11.chm2248
f1_keywords:
- vbaol11.chm2248
ms.prod: OUTLOOK
api_name:
- Outlook.Search
ms.assetid: 226a5d49-3caf-90dd-725c-265404d1939f
---


# Search Object (Outlook)

Contains information about individual searches performed against Outlook items.


## Remarks

The  **Search** object contains properties that define the type of search and the parameters of the search itself.

Use the  **[Application](http://msdn.microsoft.com/library/application-object-outlook%28Office.15%29.aspx)** object's **[AdvancedSearch](http://msdn.microsoft.com/library/application-advancedsearch-method-outlook%28Office.15%29.aspx)** method to return a **Search** object.

Use the  **[AdvancedSearchComplete](http://msdn.microsoft.com/library/application-advancedsearchcomplete-event-outlook%28Office.15%29.aspx)** event to determine when a given search has completed.


## Example

The following Microsoft Visual Basic for Applications (VBA) example returns a search object named "SubjectSearch" and displays the object's  **[Tag](http://msdn.microsoft.com/library/search-tag-property-outlook%28Office.15%29.aspx)** and **[Filter](http://msdn.microsoft.com/library/search-filter-property-outlook%28Office.15%29.aspx)** property values. The **Tag** property is used to identify a specific search once it has completed.


```
Sub SearchInboxFolder() 
 
'Searches the Inbox 
 
 
 
 Dim objSch As Search 
 
 Const strF As String = _ 
 
 "urn:schemas:mailheader:subject = 'Office Christmas Party'" 
 
 Const strS As String = "Inbox" 
 
 Const strTag As String = "SubjectSearch" 
 
 Set objSch = Application.AdvancedSearch(Scope:=strS, _ 
 
 Filter:=strF, SearchSubFolders:=True, Tag:=strTag) 
 
 
 
End Sub 
 

```

The following VBA example displays information about the search and the results of the search.




```
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 
 
 Dim objRsts As Results 
 
 MsgBox "The search " &amp; SearchObject.Tag &amp; "has completed. 
 
 Set objRsts = SearchObject.Results 
 
 'Print out number in Results collection 
 
 Debug.Print objRsts.Count 
 
 'Print out each member of Results collection 
 
 For Each Item In objRsts 
 
 Debug.Print Item 
 
 Next 
 
 
 
End Sub 
 

```


## Methods



|**Name**|
|:-----|
|[GetTable](http://msdn.microsoft.com/library/search-gettable-method-outlook%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/search-save-method-outlook%28Office.15%29.aspx)|
|[Stop](http://msdn.microsoft.com/library/search-stop-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/search-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/search-class-property-outlook%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/search-filter-property-outlook%28Office.15%29.aspx)|
|[IsSynchronous](http://msdn.microsoft.com/library/search-issynchronous-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/search-parent-property-outlook%28Office.15%29.aspx)|
|[Results](http://msdn.microsoft.com/library/search-results-property-outlook%28Office.15%29.aspx)|
|[Scope](http://msdn.microsoft.com/library/search-scope-property-outlook%28Office.15%29.aspx)|
|[SearchSubFolders](http://msdn.microsoft.com/library/search-searchsubfolders-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/search-session-property-outlook%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/search-tag-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Search Object Members](http://msdn.microsoft.com/library/search-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
