---
title: OutlineCode Members (Project)
ms.prod: PROJECTSERVER
ms.assetid: 4fa62f6d-9c70-bc1d-92fa-4d30be8b164a
---


# OutlineCode Members (Project)





## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](outlinecode-delete-method-project.md)|Deletes the  **OutlineCode** object from an **OutlineCodes** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](outlinecode-application-property-project.md)|Gets the  **[Application](application-object-project.md)** object. Read-only **Application**.|
|[CodeMask](outlinecode-codemask-property-project.md)|Gets a  **[CodeMask](codemask-object-project.md)** collection representing the code mask for an outline code in Project. Read-only **CodeMask**.|
|[DefaultValue](outlinecode-defaultvalue-property-project.md)|Gets or sets the default value of the  **[OutlineCode](outlinecode-object-project.md)** object. Read/write **String**.|
|[FieldID](outlinecode-fieldid-property-project.md)|Gets the identification number of the local outline code. Read-only  **PjCustomField**.|
|[Index](outlinecode-index-property-project.md)|Gets the index of an  **OutlineCode** object in the containing **OutlineCodes** collection. Read-only **Long**.|
|[LinkedFieldID](outlinecode-linkedfieldid-property-project.md)|Gets or sets the outline code field ID for a linked lookup table. Obsolete in Project. Read/write  **Long**.|
|[LookupTable](outlinecode-lookuptable-property-project.md)|Gets a  **[LookupTable](lookuptable-object-project.md)** collection of lookup table entries for the outline code. Read-only **LookupTable**.|
|[MatchGeneric](outlinecode-matchgeneric-property-project.md)|**True** if Project uses the enterprise text custom field (which is equivalent to an outline code) in the Resource Substitution Wizard. Read/write **Boolean**.|
|[Name](outlinecode-name-property-project.md)|Gets the name of the  **OutlineCode** object. Read/write **String**.|
|[OnlyCompleteCodes](outlinecode-onlycompletecodes-property-project.md)|**True** if only outline codes with values at all levels of the code mask can be used. Read/write **Boolean**.|
|[OnlyLeaves](outlinecode-onlyleaves-property-project.md)|**True** if only outline code lookup table values without children can be selected. Read/write **Boolean**.|
|[OnlyLookUpTableCodes](outlinecode-onlylookuptablecodes-property-project.md)|**True** if only entries listed in the local outline code lookup table can be used. Read/write **Boolean**.|
|[Parent](outlinecode-parent-property-project.md)|Gets the parent of the  **OutlineCode** object. Read-only **Project**.|
|[RequiredCode](outlinecode-requiredcode-property-project.md)|True if a value for the local outline code must be set before the project can be saved. Read/write  **Boolean**.|
|[SortOrder](outlinecode-sortorder-property-project.md)|Gets or sets the order by which the outline code items are sorted. Read/write  **[PjListOrder](pjlistorder-enumeration-project.md)**.|

