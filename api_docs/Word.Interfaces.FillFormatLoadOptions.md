# Word.Interfaces.FillFormatLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the fill formatting for a shape or text.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- backgroundColor — Returns a ColorFormat object that represents the background color for the fill.
- foregroundColor — Returns a ColorFormat object that represents the foreground color for the fill.
- gradientAngle — Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.
- gradientColorType — Gets the gradient color type.
- gradientDegree — Returns how dark or light a one-color gradient fill is. A value of 0 means that black is mixed in with the shape's foreground color to form the gradient. A value of 1 means that white is mixed in. Values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in.
- gradientStyle — Returns the gradient style for the fill.
- gradientVariant — Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.
- isVisible — Specifies if the object, or the formatting applied to it, is visible.
- pattern — Returns a PatternType value that represents the pattern applied to the fill or line.
- presetGradientType — Returns the preset gradient type for the fill.
- presetTexture — Gets the preset texture.
- rotateWithObject — Specifies whether the fill rotates with the shape.
- textureAlignment — Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.
- textureHorizontalScale — Specifies the horizontal scaling factor for the texture fill.
- textureName — Returns the name of the custom texture file for the fill.
- textureOffsetX — Specifies the horizontal offset of the texture from the origin in points.
- textureOffsetY — Specifies the vertical offset of the texture.
- textureTile — Specifies whether the texture is tiled.
- textureType — Returns the texture type for the fill.
- textureVerticalScale — Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.
- transparency — Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).
- type — Gets the fill format type.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

- Type: boolean

### backgroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the background color for the fill.

```typescript
backgroundColor?: Word.Interfaces.ColorFormatLoadOptions;
```

- Type: [Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### foregroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the foreground color for the fill.

```typescript
foregroundColor?: Word.Interfaces.ColorFormatLoadOptions;
```

- Type: [Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientAngle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.

```typescript
gradientAngle?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientColorType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the gradient color type.

```typescript
gradientColorType?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientDegree

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns how dark or light a one-color gradient fill is. A value of 0 means that black is mixed in with the shape's foreground color to form the gradient. A value of 1 means that white is mixed in. Values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in.

```typescript
gradientDegree?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the gradient style for the fill.

```typescript
gradientStyle?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientVariant

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.

```typescript
gradientVariant?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the object, or the formatting applied to it, is visible.

```typescript
isVisible?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pattern

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `PatternType` value that represents the pattern applied to the fill or line.

```typescript
pattern?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### presetGradientType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the preset gradient type for the fill.

```typescript
presetGradientType?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### presetTexture

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the preset texture.

```typescript
presetTexture?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotateWithObject

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the fill rotates with the shape.

```typescript
rotateWithObject?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureAlignment

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.

```typescript
textureAlignment?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureHorizontalScale

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal scaling factor for the texture fill.

```typescript
textureHorizontalScale?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the name of the custom texture file for the fill.

```typescript
textureName?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureOffsetX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal offset of the texture from the origin in points.

```typescript
textureOffsetX?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureOffsetY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical offset of the texture.

```typescript
textureOffsetY?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureTile

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the texture is tiled.

```typescript
textureTile?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the texture type for the fill.

```typescript
textureType?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureVerticalScale

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.

```typescript
textureVerticalScale?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the fill format type.

```typescript
type?: boolean;
```

- Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)