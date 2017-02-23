---
title: Search Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 543773b8-9f38-8d3e-2279-8f2a581ccd18
---


# Search Members (Outlook)
Contains information about individual searches performed against Outlook items.

Contains information about individual searches performed against Outlook items.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[GetTable](search-gettable-method-outlook.md)|Obtains a  **[Table](table-object-outlook.md)** object that contains items filtered by the _Filter_ parameter in a preceding **[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)** method call.|
|[Save](search-save-method-outlook.md)|Saves the search results to a Search Folder.|
|[Stop](search-stop-method-outlook.md)|Immediately ends the search that is being performed currently.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](search-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Class](search-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Filter](search-filter-property-outlook.md)|Returns a  **String** value that represents the DASL statement used to restrict the search to a specified subset of data. Read-only|
|[IsSynchronous](search-issynchronous-property-outlook.md)|Returns a  **Boolean** indicating whether the search is synchronous. Read-only.|
|[Parent](search-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Results](search-results-property-outlook.md)|Returns a  **[Results](results-object-outlook.md)** collection that specifies the results of the search. Read-only.|
|[Scope](search-scope-property-outlook.md)|Returns a  **String** that specifies the scope of the specified search. Read-only.|
|[SearchSubFolders](search-searchsubfolders-property-outlook.md)|Returns a  **Boolean** indicating whether the scope of the specified search included the subfolders of any folders searched. Read-only.|
|[Session](search-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Tag](search-tag-property-outlook.md)|Returns a  **String** specifying the name of the current search. The **Tag** property is used to identify a specific search. Read-only.|

