---
title: ModelChanges Members (Excel)
ms.prod: EXCEL
ms.assetid: 9ecee580-b4aa-9e89-1a6e-70ee31552ec7
---


# ModelChanges Members (Excel)
Represents changes made to the data model. 

Represents changes made to the data model. 


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](modelchanges-application-property-excel.md)|Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.|
|[ColumnsAdded](modelchanges-columnsadded-property-excel.md)|Returns a [ModelColumnNames Object (Excel)](modelcolumnnames-object-excel.md) collection of[ModelColumnName Object (Excel)](modelcolumnname-object-excel.md) objects which represent all columns added as part of a model operation. Read-only.|
|[ColumnsChanged](modelchanges-columnschanged-property-excel.md)|Returns a [ModelColumnChanges Object (Excel)](modelcolumnchanges-object-excel.md) collection of[ModelColumnChange Object (Excel)](modelcolumnchange-object-excel.md) objects which represent table names and column names of all table columns for which the data type was changed (might add more types of changes here in the future) as part of a model operation. Read-only.|
|[ColumnsDeleted](modelchanges-columnsdeleted-property-excel.md)|Returns a [ModelColumnNames Object (Excel)](modelcolumnnames-object-excel.md) collection of[ModelColumnName Object (Excel)](modelcolumnname-object-excel.md) objects which represent all columns which were deleted as part of a model operation. Read-only.|
|[Creator](modelchanges-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[MeasuresAdded](modelchanges-measuresadded-property-excel.md)|Returns a [ModelMeasureNames Object (Excel)](modelmeasurenames-object-excel.md) collection of[ModelMeasureName Object (Excel)](modelmeasurename-object-excel.md) objects which represent all measures which were added as part of a model operation. Read-only.|
|[Parent](modelchanges-parent-property-excel.md)|Returns an  **Object** that represents the parent object of the specified[ModelChanges](modelchanges-object-excel.md) object. Read-only.|
|[RelationshipChange](modelchanges-relationshipchange-property-excel.md)| When **True** , one or more relationships in the model were changed (added, deleted or modified) as part of a model operation. When **False** , no relationships were changed during the operation. **Boolean** Read-only.|
|[Source](modelchanges-source-property-excel.md)||
|[TableNamesChanged](modelchanges-tablenameschanged-property-excel.md)|Returns a [ModelTableNameChanges Object (Excel)](modeltablenamechanges-object-excel.md) collection of[ModelTableNameChange Object (Excel)](modeltablenamechange-object-excel.md) objects representing old and new names of all tables which were renamed in the model as part of a model operation. Read-only.|
|[TablesAdded](modelchanges-tablesadded-property-excel.md)|Returns a [ModelTableNames Object (Excel)](modeltablenames-object-excel.md) collection of table names as strings representing all tables which were added to the model as part of a model operation. Read-only.|
|[TablesDeleted](modelchanges-tablesdeleted-property-excel.md)|Returns a [ModelTableNames Object (Excel)](modeltablenames-object-excel.md) collection of table names as strings representing all tables which were deleted from the model as part of a model operation. Read-only.|
|[TablesModified](modelchanges-tablesmodified-property-excel.md)|Returns a [ModelTableNames Object (Excel)](modeltablenames-object-excel.md) collection of table names as strings representing all tables which were refreshed or recalculated as part of a model operation. Read-only.|
|[UnknownChange](modelchanges-unknownchange-property-excel.md)| **True** when a non-specified change was made to the model as part of a model transaction. **Boolean** Read-only|

