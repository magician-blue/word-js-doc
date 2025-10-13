# Word.Interfaces.GlowFormatData interface

- Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `glowFormat.toJSON()`.

## Properties

- [color](#color): Returns a `ColorFormat` object that represents the color for a glow effect.
- [radius](#radius): Specifies the length of the radius for a glow effect.
- [transparency](#transparency): Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).

## Property Details

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the color for a glow effect.

```typescript
color?: Word.Interfaces.ColorFormatData;
```

Property value: [Word.Interfaces.ColorFormatData](/en-us/javascript/api/word/word.interfaces.colorformatdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### radius

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the radius for a glow effect.

```typescript
radius?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)