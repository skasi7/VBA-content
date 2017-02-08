---
title: Group2 Members (Project)
ms.prod: PROJECTSERVER
ms.assetid: 69c5069c-3fd6-fbb5-d886-ebbda667cba4
---


# Group2 Members (Project)
Represents a group definition where the group hierarchy can be maintained. A  **Group2** object is a member of a **[Groups2](groups2-object-project.md)**, **[ResourceGroups2](resourcegroups2-object-project.md)**, or **[TaskGroups2](taskgroups2-object-project.md)** collection.

Represents a group definition where the group hierarchy can be maintained. A  **Group2** object is a member of a **[Groups2](groups2-object-project.md)**, **[ResourceGroups2](resourcegroups2-object-project.md)**, or **[TaskGroups2](taskgroups2-object-project.md)** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](group2-delete-method-project.md)|Deletes the  **Group2** object from a **Groups2** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](group2-application-property-project.md)|Gets the  **[Application](application-object-project.md)** object. Read-only **Application**.|
|[GroupAssignments](group2-groupassignments-property-project.md)|**True** if assignments are grouped, rather than tasks or resources. Read/write **Boolean**.|
|[GroupCriteria](group2-groupcriteria-property-project.md)|Gets or sets the  **[GroupCriteria2](groupcriteria2-object-project.md)** collection representing the fields in a group definition. Read/write **GroupCriteria2**.|
|[Index](group2-index-property-project.md)|Gets the index of a  **Group2** object in a **ResourceGroups2** collection or **TaskGroups2** collection. Read-only **Long**.|
|[MaintainHierarchy](group2-maintainhierarchy-property-project.md)|Gets or sets a value that specifies whether hierarchy is maintained in the group view. Read/write  **Boolean**.|
|[Name](group2-name-property-project.md)|Gets or sets the name of a  **Group2** object. Read/write **String**.|
|[Parent](group2-parent-property-project.md)|Gets the parent of the object. Read-only  **Project**.|
|[ShowSummary](group2-showsummary-property-project.md)|**True** if summary tasks are displayed in a task view that is organized by group. Read/write **Boolean**.|

