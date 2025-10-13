# Word.Interfaces.GlowFormatLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the glow formatting for the font used by the range of text.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all
  - Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- color
  - Returns a ColorFormat object that represents the color for a glow effect.
- radius
  - Specifies the length of the radius for a glow effect.
- transparency
  - Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value
- boolean

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ColorFormat object that represents the color for a glow effect.

```typescript
color?: Word.Interfaces.ColorFormatLoadOptions;
```

Property Value
- [Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### radius

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the radius for a glow effect.

```typescript
radius?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)