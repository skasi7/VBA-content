---
title: SignatureSet Object (Office)
keywords: vbaof11.chm247000
f1_keywords:
- vbaof11.chm247000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SignatureSet
ms.assetid: 574cba16-c632-ab66-f014-58172ff1c091
---


# SignatureSet Object (Office)

A collection of  **Signature** objects that correspond to the digital signature attached to a document.


## Remarks

Use the  **Signatures** property of the **Document** object to return a **SignatureSet** collection; for example:


```
Set sigs = ActiveDocument.Signatures
```

You can add a  **Signature** object to a **SignatureSet** collection using the **Add** method and you can return an existing member using the **Item** method. The **AddSignatureLine** method also adds a **Signature** object to the collection. Also see the **Subset** property, which acts as a filter for whether certain **Signature** objects appear in the collection. To remove a **Signature** from a **SignatureSet** collection, use the **Delete** method of the **Signature** object.


## Example

The following example prompts the user to select a digital signature with which to sign the active document in Microsoft Word. To use this example, open a document in Word and pass this function the name of a certificate issuer and the name of a certificate signer that match the  **Issued By** and **Issued To** fields of a digital certificate in the **Digital Certificates** dialog box. This example will test to make sure that the digital signature that the user selects meets certain criteria, such as not having expired, before the new signature is committed to the disk.


```
Function AddSignature(ByVal strIssuer As String, _ 
 strSigner As String) As Boolean 
 
 Dim sig As Signature 
 
 'Display the dialog box that lets the 
 'user select a digital signature. 
 'If the user selects a signature, then 
 'it is added to the Signatures 
 'collection. If the user doesn't, then 
 'an error is returned. 
 Set sig = ActiveDocument.Signatures.Add 
 
 'Test several properties before committing the Signature object to disk. 
 If sig.Issuer = strIssuer And _ 
 sig.Signer = strSigner And _ 
 sig.IsCertificateExpired = False And _ 
 sig.IsCertificateRevoked = False And _ 
 sig.IsValid = True Then 
 
 MsgBox "Signed" 
 AddSignature = True 
 'Otherwise, remove the Signature object from the SignatureSet collection. 
 Else 
 sig.Delete 
 MsgBox "Not signed" 
 AddSignature = False 
 End If 
 
End Function
```


## Methods



|**Name**|
|:-----|
|[AddNonVisibleSignature](http://msdn.microsoft.com/library/signatureset-addnonvisiblesignature-method-office%28Office.15%29.aspx)|
|[AddSignatureLine](http://msdn.microsoft.com/library/signatureset-addsignatureline-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/signatureset-application-property-office%28Office.15%29.aspx)|
|[CanAddSignatureLine](http://msdn.microsoft.com/library/signatureset-canaddsignatureline-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/signatureset-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/signatureset-creator-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/signatureset-item-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/signatureset-parent-property-office%28Office.15%29.aspx)|
|[ShowSignaturesPane](http://msdn.microsoft.com/library/signatureset-showsignaturespane-property-office%28Office.15%29.aspx)|
|[Subset](http://msdn.microsoft.com/library/signatureset-subset-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[SignatureSet Object Members](http://msdn.microsoft.com/library/signatureset-members-office%28Office.15%29.aspx)
