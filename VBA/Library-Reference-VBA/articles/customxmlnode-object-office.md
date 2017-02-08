---
title: CustomXMLNode Object (Office)
keywords: vbaof11.chm294000
f1_keywords:
- vbaof11.chm294000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CustomXMLNode
ms.assetid: e90213f5-6d62-52d8-3043-2399eaa5aaba
---


# CustomXMLNode Object (Office)

Represents an XML node in a tree in a document. The  **CustomXMLNode** object is a member of the **CustomXMLNodes** collection.


## Remarks

The  **CustomXMLNode** object is designed to have functional parity with the **IXMLDOMNode** interface. In addition, it contains an **XPath** property, which is a great improvement over the objects provided by MSXML.


## Example

The following example selects a single node from a  **CustomXMLPart** object by using an XPath expression and assigns it to a **CustomXMLNode** object.


```
Sub CustomXmlNodes()  
    Dim cxp1 As CustomXMLPart 
    Dim cxn As CustomXMLNode 
 
    With ActiveDocument 
 
        ' Returns the first custom xml part with the given root namespace. 
        Set cxp1 = .CustomXMLParts("urn:invoice:namespace")  
         
        ' Get the first node matching the XPath expression.                              
        Set cxn = cxp1.SelectSingleNode("//*[@quantity < 4]") 
                 
    End With 
     
End Sub
```


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AppendChildNode](http://msdn.microsoft.com/library/customxmlnode-appendchildnode-method-office%28Office.15%29.aspx)||
|[AppendChildSubtree](http://msdn.microsoft.com/library/customxmlnode-appendchildsubtree-method-office%28Office.15%29.aspx)||
|[Delete](http://msdn.microsoft.com/library/customxmlnode-delete-method-office%28Office.15%29.aspx)||
|[HasChildNodes](http://msdn.microsoft.com/library/customxmlnode-haschildnodes-method-office%28Office.15%29.aspx)||
|[InsertNodeBefore](http://msdn.microsoft.com/library/customxmlnode-insertnodebefore-method-office%28Office.15%29.aspx)||
|[InsertSubtreeBefore](http://msdn.microsoft.com/library/customxmlnode-insertsubtreebefore-method-office%28Office.15%29.aspx)||
|[RemoveChild](http://msdn.microsoft.com/library/customxmlnode-removechild-method-office%28Office.15%29.aspx)||
|[ReplaceChildNode](http://msdn.microsoft.com/library/customxmlnode-replacechildnode-method-office%28Office.15%29.aspx)||
|[ReplaceChildSubtree](http://msdn.microsoft.com/library/customxmlnode-replacechildsubtree-method-office%28Office.15%29.aspx)||
|[SelectNodes](http://msdn.microsoft.com/library/customxmlnode-selectnodes-method-office%28Office.15%29.aspx)||
|[SelectSingleNode](http://msdn.microsoft.com/library/customxmlnode-selectsinglenode-method-office%28Office.15%29.aspx)||

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/customxmlnode-application-property-office%28Office.15%29.aspx)|
|[Attributes](http://msdn.microsoft.com/library/customxmlnode-attributes-property-office%28Office.15%29.aspx)|
|[BaseName](http://msdn.microsoft.com/library/customxmlnode-basename-property-office%28Office.15%29.aspx)|
|[ChildNodes](http://msdn.microsoft.com/library/customxmlnode-childnodes-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/customxmlnode-creator-property-office%28Office.15%29.aspx)|
|[FirstChild](http://msdn.microsoft.com/library/customxmlnode-firstchild-property-office%28Office.15%29.aspx)|
|[LastChild](http://msdn.microsoft.com/library/customxmlnode-lastchild-property-office%28Office.15%29.aspx)|
|[NamespaceURI](http://msdn.microsoft.com/library/customxmlnode-namespaceuri-property-office%28Office.15%29.aspx)|
|[NextSibling](http://msdn.microsoft.com/library/customxmlnode-nextsibling-property-office%28Office.15%29.aspx)|
|[NodeType](http://msdn.microsoft.com/library/customxmlnode-nodetype-property-office%28Office.15%29.aspx)|
|[NodeValue](http://msdn.microsoft.com/library/customxmlnode-nodevalue-property-office%28Office.15%29.aspx)|
|[OwnerDocument](http://msdn.microsoft.com/library/customxmlnode-ownerdocument-property-office%28Office.15%29.aspx)|
|[OwnerPart](http://msdn.microsoft.com/library/customxmlnode-ownerpart-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/customxmlnode-parent-property-office%28Office.15%29.aspx)|
|[ParentNode](http://msdn.microsoft.com/library/customxmlnode-parentnode-property-office%28Office.15%29.aspx)|
|[PreviousSibling](http://msdn.microsoft.com/library/customxmlnode-previoussibling-property-office%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/customxmlnode-text-property-office%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/customxmlnode-xml-property-office%28Office.15%29.aspx)|
|[XPath](http://msdn.microsoft.com/library/customxmlnode-xpath-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[CustomXMLNode Object Members](http://msdn.microsoft.com/library/customxmlnode-members-office%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
