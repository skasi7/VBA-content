---
title: XMLMapping Members (Word)
ms.prod: WORD
ms.assetid: 8fb27e7a-1d02-4754-87ca-f117cc67cdff
---


# XMLMapping Members (Word)
Represents the XML mapping on a  **[ContentControl](contentcontrol-object-word.md)** object between custom XML and a content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document.

Represents the XML mapping on a  **[ContentControl](contentcontrol-object-word.md)** object between custom XML and a content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](xmlmapping-delete-method-word.md)|Deletes the XML mapping from the parent content control.|
|[SetMapping](xmlmapping-setmapping-method-word.md)|Allows creating or changing the XML mapping on a content control. Returns  **True** if Microsoft Word maps the content control to a custom XML node in the document?s custom XML data store.|
|[SetMappingByNode](xmlmapping-setmappingbynode-method-word.md)|Allows creating or changing the XML data mapping on a content control. Returns  **True** if Microsoft Word maps the content control to a custom XML node in the document?s custom XML data store.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](xmlmapping-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](xmlmapping-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .|
|[CustomXMLNode](xmlmapping-customxmlnode-property-word.md)|Returns a  **CustomXMLNode** object that represents the custom XML node in the data store to which the content control in the document maps.|
|[CustomXMLPart](xmlmapping-customxmlpart-property-word.md)|Returns a  **CustomXMLPart** object that represents the custom XML part to which the content control in the document maps.|
|[IsMapped](xmlmapping-ismapped-property-word.md)|Returns a  **Boolean** that represents whether the content control in the document is mapped to an XML node in the document's XML data store. Read-only.|
|[Parent](xmlmapping-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **XMLMapping** object.|
|[PrefixMappings](xmlmapping-prefixmappings-property-word.md)|Returns a  **String** that represents the prefix mappings used to evaluate the XPath for the current XML mapping. Read-only.|
|[XPath](xmlmapping-xpath-property-word.md)|Returns a  **String** that represents the XPath for the XML mapping, which evaluates to the currently mapped XML node. Read-only.|

