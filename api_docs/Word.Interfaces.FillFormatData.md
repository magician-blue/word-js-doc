# Word.Interfaces.FillFormatData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface describing the data returned by calling `fillFormat.toJSON()`.

## Properties

- `backgroundColor`: Returns a `ColorFormat` object that represents the background color for the fill.
- `foregroundColor`: Returns a `ColorFormat` object that represents the foreground color for the fill.
- `gradientAngle`: Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.
- `gradientColorType`: Gets the gradient color type.
- `gradientDegree`: Returns how dark or light a one-color gradient fill is. A value of 0 means that black is mixed in with the shape's foreground color to form the gradient. A value of 1 means that white is mixed in. Values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in.
- `gradientStyle`: Returns the gradient style for the fill.
- `gradientVariant`: Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.
- `isVisible`: Specifies if the object, or the formatting applied to it, is visible.
- `pattern`: Returns a `PatternType` value that represents the pattern applied to the fill or line.
- `presetGradientType`: Returns the preset gradient type for the fill.
- `presetTexture`: Gets the preset texture.
- `rotateWithObject`: Specifies whether the fill rotates with the shape.
- `textureAlignment`: Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.
- `textureHorizontalScale`: Specifies the horizontal scaling factor for the texture fill.
- `textureName`: Returns the name of the custom texture file for the fill.
- `textureOffsetX`: Specifies the horizontal offset of the texture from the origin in points.
- `textureOffsetY`: Specifies the vertical offset of the texture.
- `textureTile`: Specifies whether the texture is tiled.
- `textureType`: Returns the texture type for the fill.
- `textureVerticalScale`: Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.
- `transparency`: Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).
- `type`: Gets the fill format type.

## Property details

### backgroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the background color for the fill.

```typescript
backgroundColor?: Word.Interfaces.ColorFormatData;
```

Property value: [Word.Interfaces.ColorFormatData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.colorformatdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### foregroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the foreground color for the fill.

```typescript
foregroundColor?: Word.Interfaces.ColorFormatData;
```

Property value: [Word.Interfaces.ColorFormatData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.colorformatdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### gradientAngle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.

```typescript
gradientAngle?: number;
```

Property value: `number`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### gradientColorType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the gradient color type.

```typescript
gradientColorType?: Word.GradientColorType | "Mixed" | "OneColor" | "TwoColors" | "PresetColors" | "MultiColor";
```

Property value: [Word.GradientColorType](https://learn.microsoft.com/en-us/javascript/api/word/word.gradientcolortype) | "Mixed" | "OneColor" | "TwoColors" | "PresetColors" | "MultiColor"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### gradientDegree

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns how dark or light a one-color gradient fill is. A value of 0 means that black is mixed in with the shape's foreground color to form the gradient. A value of 1 means that white is mixed in. Values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in.

```typescript
gradientDegree?: number;
```

Property value: `number`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### gradientStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the gradient style for the fill.

```typescript
gradientStyle?: Word.GradientStyle | "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter";
```

Property value: [Word.GradientStyle](https://learn.microsoft.com/en-us/javascript/api/word/word.gradientstyle) | "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### gradientVariant

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.

```typescript
gradientVariant?: number;
```

Property value: `number`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the object, or the formatting applied to it, is visible.

```typescript
isVisible?: boolean;
```

Property value: `boolean`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pattern

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `PatternType` value that represents the pattern applied to the fill or line.

```typescript
pattern?: Word.PatternType | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross";
```

Property value: [Word.PatternType](https://learn.microsoft.com/en-us/javascript/api/word/word.patterntype) | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetGradientType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the preset gradient type for the fill.

```typescript
presetGradientType?: Word.PresetGradientType | "Mixed" | "EarlySunset" | "LateSunset" | "Nightfall" | "Daybreak" | "Horizon" | "Desert" | "Ocean" | "CalmWater" | "Fire" | "Fog" | "Moss" | "Peacock" | "Wheat" | "Parchment" | "Mahogany" | "Rainbow" | "RainbowII" | "Gold" | "GoldII" | "Brass" | "Chrome" | "ChromeII" | "Silver" | "Sapphire";
```

Property value: [Word.PresetGradientType](https://learn.microsoft.com/en-us/javascript/api/word/word.presetgradienttype) | "Mixed" | "EarlySunset" | "LateSunset" | "Nightfall" | "Daybreak" | "Horizon" | "Desert" | "Ocean" | "CalmWater" | "Fire" | "Fog" | "Moss" | "Peacock" | "Wheat" | "Parchment" | "Mahogany" | "Rainbow" | "RainbowII" | "Gold" | "GoldII" | "Brass" | "Chrome" | "ChromeII" | "Silver" | "Sapphire"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetTexture

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the preset texture.

```typescript
presetTexture?: Word.PresetTexture | "Mixed" | "Papyrus" | "Canvas" | "Denim" | "WovenMat" | "WaterDroplets" | "PaperBag" | "FishFossil" | "Sand" | "GreenMarble" | "WhiteMarble" | "BrownMarble" | "Granite" | "Newsprint" | "RecycledPaper" | "Parchment" | "Stationery" | "BlueTissuePaper" | "PinkTissuePaper" | "PurpleMesh" | "Bouquet" | "Cork" | "Walnut" | "Oak" | "MediumWood";
```

Property value: [Word.PresetTexture](https://learn.microsoft.com/en-us/javascript/api/word/word.presettexture) | "Mixed" | "Papyrus" | "Canvas" | "Denim" | "WovenMat" | "WaterDroplets" | "PaperBag" | "FishFossil" | "Sand" | "GreenMarble" | "WhiteMarble" | "BrownMarble" | "Granite" | "Newsprint" | "RecycledPaper" | "Parchment" | "Stationery" | "BlueTissuePaper" | "PinkTissuePaper" | "PurpleMesh" | "Bouquet" | "Cork" | "Walnut" | "Oak" | "MediumWood"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotateWithObject

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the fill rotates with the shape.

```typescript
rotateWithObject?: boolean;
```

Property value: `boolean`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textureAlignment

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.

```typescript
textureAlignment?: Word.TextureAlignment | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "Center" | "Right" | "BottomLeft" | "Bottom" | "BottomRight";
```

Property value: [Word.TextureAlignment](https://learn.microsoft.com/en-us/javascript/api/word/word.texturealignment) | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "Center" | "Right" | "BottomLeft" | "Bottom" | "BottomRight"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textureHorizontalScale

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal scaling factor for the texture fill.

```typescript
textureHorizontalScale?: number;
```

Property value: `number`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textureName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the name of the custom texture file for the fill.

```typescript
textureName?: string;
```

Property value: `string`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textureOffsetX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal offset of the texture from the origin in points.

```typescript
textureOffsetX?: number;
```

Property value: `number`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textureOffsetY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical offset of the texture.

```typescript
textureOffsetY?: number;
```

Property value: `number`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textureTile

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the texture is tiled.

```typescript
textureTile?: boolean;
```

Property value: `boolean`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textureType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the texture type for the fill.

```typescript
textureType?: Word.TextureType | "Mixed" | "Preset" | "UserDefined";
```

Property value: [Word.TextureType](https://learn.microsoft.com/en-us/javascript/api/word/word.texturetype) | "Mixed" | "Preset" | "UserDefined"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textureVerticalScale

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.

```typescript
textureVerticalScale?: number;
```

Property value: `number`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: number;
```

Property value: `number`

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the fill format type.

```typescript
type?: Word.FillType | "Mixed" | "Solid" | "Patterned" | "Gradient" | "Textured" | "Background" | "Picture";
```

Property value: [Word.FillType](https://learn.microsoft.com/en-us/javascript/api/word/word.filltype) | "Mixed" | "Solid" | "Patterned" | "Gradient" | "Textured" | "Background" | "Picture"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)