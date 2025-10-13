# Word.Interfaces.ShadingUpdateData interface

- Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `Shading` object, for use in `shading.set({ ... })`.

## Properties

- backgroundPatternColor: Specifies the color for the background of the object. You can provide the value in the '#RRGGBB' format or the color name.
- foregroundPatternColor: Specifies the color for the foreground of the object. You can provide the value in the '#RRGGBB' format or the color name.
- texture: Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see [Add, change, or delete the background color in Word](https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515).

## Property Details

### backgroundPatternColor

Specifies the color for the background of the object. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
backgroundPatternColor?: string;
```

Property value: string

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### foregroundPatternColor

Specifies the color for the foreground of the object. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
foregroundPatternColor?: string;
```

Property value: string

Remarks: [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### texture

Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see [Add, change, or delete the background color in Word](https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515).

```typescript
texture?: Word.ShadingTextureType | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid";
```

Property value: [Word.ShadingTextureType](/en-us/javascript/api/word/word.shadingtexturetype) | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid"

Remarks: [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)