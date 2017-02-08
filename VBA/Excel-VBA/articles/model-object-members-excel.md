---
title: Model Object Members (Excel)
ms.prod: EXCEL
ms.assetid: 2fb97d8c-f889-469a-95c7-97a2bcc24635
---


# Model Object Members (Excel)
Stores references to workbook connections and information about the tables and relationships contained within the Data Model.

Stores references to workbook connections and information about the tables and relationships contained within the Data Model.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddConnection](model-addconnection-method-excel.md)|Adds a new Workbook Connection to the model with the same properties as the one supplied as an argument.|
|[CreateModelWorkbookConnection](model-createmodelworkbookconnection-method-excel.md)|Returns a [WorkbookConnection Object (Excel)](workbookconnection-object-excel.md) object of type[ModelConnection Object (Excel)](modelconnection-object-excel.md). |
|[Initialize](model-initialize-method-excel.md)|Initializes the Workbook's Data Model. This is called by default the first time the model is used.|
|[Refresh](model-refresh-method-excel.md)|Refreshes all data sources associated with the model, fully reprocesses the model and updates all Excel data features associated with the Model.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](model-application-property-excel.md)|Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.|
|[Creator](model-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DataModelConnection](model-datamodelconnection-property-excel.md)|Returns the model workbook connection object from the workbook connections collection which connects to the model.|
|[ModelRelationships](model-modelrelationships-property-excel.md)|Returns a collection of relationships between Data Model tables. Read-only|
|[ModelTables](model-modeltables-property-excel.md)|Returns a collection of tables inside the Data Model. Read-only|
|[Name](model-name-property-excel.md)|Returns a  **String** value representing the name of the **Model** object. Read-only|
|[Parent](model-parent-property-excel.md)|Returns an  **Object** that represents the parent object of the specified object. Read-only.|
|[ModelFormatBoolean](model-modelformatboolean-property-excel.md)|Returns a [ModelFormatBoolean](modelformatboolean-object-excel.md) object that represents formatting of type true/false in the data model. Read-only .|
|[ModelFormatCurrency](model-modelformatcurrency-property-excel.md)|Returns a [ModelFormatCurrency](modelformatcurrency-object-excel.md) object that represents formatting of type currency in the data model. Read-only .|
|[ModelFormatDate](model-modelformatdate-property-excel.md)|Returns a [ModelFormatDate](modelformatdate-object-excel.md) object that represents formatting of type date in the data model. Read-only .|
|[ModelFormatDecimalNumber](model-modelformatdecimalnumber-property-excel.md)|Returns a [ModelFormatDecimalNumber](modelformatdecimalnumber-object-excel.md) object that represents formatting of type decimal number in the data model. Read-only.|
|[ModelFormatGeneral](model-modelformatgeneral-property-excel.md)|Returns a [ModelFormatGeneral](modelformatgeneral-object-excel.md) object that represents formatting of type general in the data model. Read-only.|
|[ModelFormatPercentageNumber](model-modelformatpercentagenumber-property-excel.md)|Returns a [ModelFormatPercentageNumber](modelformatpercentagenumber-object-excel.md) object that represents formatting of type percentage number in the data model. Read-only .|
|[ModelFormatScientificNumber](model-modelformatscientificnumber-property-excel.md)|Returns a [ModelFormatScientificNumber](modelformatscientificnumber-object-excel.md) object that represents formatting of type scientific number in the data model. Read-only.|
|[ModelFormatWholeNumber](model-modelformatwholenumber-property-excel.md)|Returns a [ModelFormatWholeNumber](modelformatwholenumber-object-excel.md) object that represents formatting of type whole number in the data model. Read-only .|
|[ModelMeasures](model-modelmeasures-property-excel.md)|Returns a [ModelMeasures Object (Excel)](modelmeasures-object-excel.md) object that represents the collection of model measures in the data model. Read-only.|

