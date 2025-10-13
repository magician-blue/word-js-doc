# Word.Interfaces.FillFormatUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the FillFormat object, for use in `fillFormat.set({ ... })`.

## Properties

- backgroundColor: Returns a ColorFormat object that represents the background color for the fill.
- foregroundColor: Returns a ColorFormat object that represents the foreground color for the fill.
- gradientAngle: Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.
- isVisible: Specifies if the object, or the formatting applied to it, is visible.
- rotateWithObject: Specifies whether the fill rotates with the shape.
- textureAlignment: Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.
- textureHorizontalScale: Specifies the horizontal scaling factor for the texture fill.
- textureOffsetX: Specifies the horizontal offset of the texture from the origin in points.
- textureOffsetY: Specifies the vertical offset of the texture.
- textureTile: Specifies whether the texture is tiled.
- textureVerticalScale: Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.
- transparency: Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).

## Property Details

### backgroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the background color for the fill.

```typescript
backgroundColor?: Word.Interfaces.ColorFormatUpdateData;
```

Type: [Word.Interfaces.ColorFormatUpdateData](/en-us/javascript/api/word/word.interfaces.colorformatupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### foregroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the foreground color for the fill.

```typescript
foregroundColor?: Word.Interfaces.ColorFormatUpdateData;
```

Type: [Word.Interfaces.ColorFormatUpdateData](/en-us/javascript/api/word/word.interfaces.colorformatupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientAngle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.

```typescript
gradientAngle?: number;
```

Type: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the object, or the formatting applied to it, is visible.

```typescript
isVisible?: boolean;
```

Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotateWithObject

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the fill rotates with the shape.

```typescript
rotateWithObject?: boolean;
```

Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureAlignment

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.

```typescript
textureAlignment?: Word.TextureAlignment | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "Center" | "Right" | "BottomLeft" | "Bottom" | "BottomRight";
```

Type: [Word.TextureAlignment](/en-us/javascript/api/word/word.texturealignment) | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "Center" | "Right" | "BottomLeft" | "Bottom" | "BottomRight"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureHorizontalScale

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal scaling factor for the texture fill.

```typescript
textureHorizontalScale?: number;
```

Type: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureOffsetX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal offset of the texture from the origin in points.

```typescript
textureOffsetX?: number;
```

Type: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureOffsetY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical offset of the texture.

```typescript
textureOffsetY?: number;
```

Type: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureTile

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the texture is tiled.

```typescript
textureTile?: boolean;
```

Type: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureVerticalScale

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.

```typescript
textureVerticalScale?: number;
```

Type: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: number;
```

Type: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)