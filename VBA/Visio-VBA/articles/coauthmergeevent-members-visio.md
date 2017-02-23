---
title: CoauthMergeEvent Members (Visio)
ms.prod: VISIO
ms.assetid: 268dee02-6c6f-80bf-abc7-762174406ec9
---


# CoauthMergeEvent Members (Visio)
Represents the merging of changes by various authors to a document or documents. Passed as a parameter to the [Document.AfterDocumentMerge](document-afterdocumentmerge-event-visio.md) and[Documents.AfterDocumentMerge](documents-afterdocumentmerge-event-visio.md) events.

Represents the merging of changes by various authors to a document or documents. Passed as a parameter to the [Document.AfterDocumentMerge](document-afterdocumentmerge-event-visio.md) and[Documents.AfterDocumentMerge](documents-afterdocumentmerge-event-visio.md) events.


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](coauthmergeevent-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[BaseDocument](coauthmergeevent-basedocument-property-visio.md)|Returns a [Document](document-object-visio.md) object that represents the state of the original document before any changes by any users are merged into it. Read-only.|
|[DownloadDocument](coauthmergeevent-downloaddocument-property-visio.md)|Returns a [Document](document-object-visio.md) object that represents a document that includes changes from all users other than the current user. Read-only.|
|[ObjectType](coauthmergeevent-objecttype-property-visio.md)|Returns the type of an object. Read-only.|
|[Stat](coauthmergeevent-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[WorkingDocument](coauthmergeevent-workingdocument-property-visio.md)|Returns a [Document](document-object-visio.md) object that represents a merged document that includes changes by the current user only. Read-only.|

