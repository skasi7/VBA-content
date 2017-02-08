---
title: PictureEffect Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.PictureEffect
ms.assetid: af3f742a-e082-1abd-7df2-d1fb2f57c8a2
---


# PictureEffect Object (Office)

Represents a Picture Effect.


## Remarks

Picture Effects are processed as a chain composed of individual items which are applied in sequence to create the final composited image. An Effects chain will allow an effect to be added to the chain, reordered, or removed from the chain.


## Example

The following code sets several Picture Effect fill properties on a shape in a Microsoft PowerPoint slide.


```
Sub PictureEffectSample() 
' Setup a slide with one picture shape. 
With ActivePresentation.Slides(1).Shapes(1).Fill.PictureEffects 
 
 ' Insert a 150% Saturation effect. 
 .Insert(msoEffectSaturation).EffectParameters(1).Value = 1.5 
 
 ' Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast. 
 Dim brightnessContrast As PictureEffect 
 Set brightnessContrast = .Insert(msoEffectBrightnessContrast) 
 brightnessContrast.EffectParameters(1).Value = -0.5 
 brightnessContrast.EffectParameters(2).Value = 0.25 
 
 ' Remove all Picture effects. 
 While .Count > 0 
 .Delete (1) 
 Wend 
 
End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/pictureeffect-delete-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/pictureeffect-application-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/pictureeffect-creator-property-office%28Office.15%29.aspx)|
|[EffectParameters](http://msdn.microsoft.com/library/pictureeffect-effectparameters-property-office%28Office.15%29.aspx)|
|[Position](http://msdn.microsoft.com/library/pictureeffect-position-property-office%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/pictureeffect-type-property-office%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/pictureeffect-visible-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[PictureEffect Object Members](http://msdn.microsoft.com/library/pictureeffect-members-office%28Office.15%29.aspx)
