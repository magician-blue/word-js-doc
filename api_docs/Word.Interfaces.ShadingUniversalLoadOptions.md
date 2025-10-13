# Word.Interfaces.ShadingUniversalLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the ShadingUniversal object, which manages shading for a range, paragraph, frame, or table.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [$all](#all)  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- [backgroundPatternColor](#backgroundpatterncolor)  
  Specifies the color that's applied to the background of the ShadingUniversal object. You can provide the value in the '#RRGGBB' format.

- [backgroundPatternColorIndex](#backgroundpatterncolorindex)  
  Specifies the color that's applied to the background of the ShadingUniversal object.

- [foregroundPatternColor](#foregroundpatterncolor)  
  Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern. You can provide the value in the '#RRGGBB' format.

- [foregroundPatternColorIndex](#foregroundpatterncolorindex)  
  Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern.

- [texture](#texture)  
  Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see Add, change, or delete the background color in Word.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

- Property value: boolean

### backgroundPatternColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the background of the ShadingUniversal object. You can provide the value in the '#RRGGBB' format.

```typescript
backgroundPatternColor?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### backgroundPatternColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the background of the ShadingUniversal object.

```typescript
backgroundPatternColorIndex?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### foregroundPatternColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern. You can provide the value in the '#RRGGBB' format.

```typescript
foregroundPatternColor?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### foregroundPatternColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern.

```typescript
foregroundPatternColorIndex?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### texture

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see [Add, change, or delete the background color in Word](https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515).

```typescript
texture?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)