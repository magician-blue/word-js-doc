# Word.Interfaces.ShadingUniversalUpdateData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface for updating data on the ShadingUniversal object, for use in shadingUniversal.set({ ... }).

## Properties

- backgroundPatternColor
  - Specifies the color that's applied to the background of the ShadingUniversal object. You can provide the value in the '#RRGGBB' format.
- backgroundPatternColorIndex
  - Specifies the color that's applied to the background of the ShadingUniversal object.
- foregroundPatternColor
  - Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern. You can provide the value in the '#RRGGBB' format.
- foregroundPatternColorIndex
  - Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern.
- texture
  - Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515.

## Property Details

### backgroundPatternColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the background of the ShadingUniversal object. You can provide the value in the '#RRGGBB' format.

```typescript
backgroundPatternColor?: string;
```

Type: string

Remarks: API set: WordApi BETA (PREVIEW ONLY): https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### backgroundPatternColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the background of the ShadingUniversal object.

```typescript
backgroundPatternColorIndex?: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.colorindex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

Remarks: API set: WordApi BETA (PREVIEW ONLY): https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### foregroundPatternColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern. You can provide the value in the '#RRGGBB' format.

```typescript
foregroundPatternColor?: string;
```

Type: string

Remarks: API set: WordApi BETA (PREVIEW ONLY): https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### foregroundPatternColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern.

```typescript
foregroundPatternColorIndex?: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.colorindex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

Remarks: API set: WordApi BETA (PREVIEW ONLY): https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### texture

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515.

```typescript
texture?: Word.ShadingTextureType | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid";
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.shadingtexturetype | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid"

Remarks: API set: WordApi BETA (PREVIEW ONLY): https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets