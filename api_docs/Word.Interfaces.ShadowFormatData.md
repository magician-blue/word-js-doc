# Word.Interfaces.ShadowFormatData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `shadowFormat.toJSON()`.

## Properties

- [blur](#blur) - Specifies the blur level for a shadow format as a value between 0.0 and 100.0.
- [foregroundColor](#foregroundcolor) - Returns a `ColorFormat` object that represents the foreground color for the fill, line, or shadow.
- [isVisible](#isvisible) - Specifies whether the object or the formatting applied to it is visible.
- [obscured](#obscured) - Specifies `true` if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill, `false` if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
- [offsetX](#offsetx) - Specifies the horizontal offset (in points) of the shadow from the shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
- [offsetY](#offsety) - Specifies the vertical offset (in points) of the shadow from the shape. A positive value offsets the shadow to the top of the shape; a negative value offsets it to the bottom.
- [rotateWithShape](#rotatewithshape) - Specifies whether to rotate the shadow when rotating the shape.
- [size](#size) - Specifies the width of the shadow.
- [style](#style) - Specifies the type of shadow formatting to apply to a shape.
- [transparency](#transparency) - Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).
- [type](#type) - Specifies the shape shadow type.

## Property Details

### blur

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the blur level for a shadow format as a value between 0.0 and 100.0.

```typescript
blur?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### foregroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the foreground color for the fill, line, or shadow.

```typescript
foregroundColor?: Word.Interfaces.ColorFormatData;
```

Property value: [Word.Interfaces.ColorFormatData](/en-us/javascript/api/word/word.interfaces.colorformatdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the object or the formatting applied to it is visible.

```typescript
isVisible?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### obscured

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies `true` if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill, `false` if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.

```typescript
obscured?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### offsetX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal offset (in points) of the shadow from the shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.

```typescript
offsetX?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### offsetY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical offset (in points) of the shadow from the shape. A positive value offsets the shadow to the top of the shape; a negative value offsets it to the bottom.

```typescript
offsetY?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotateWithShape

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to rotate the shadow when rotating the shape.

```typescript
rotateWithShape?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### size

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the shadow.

```typescript
size?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### style

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the type of shadow formatting to apply to a shape.

```typescript
style?: Word.ShadowStyle | "Mixed" | "OuterShadow" | "InnerShadow";
```

Property value: [Word.ShadowStyle](/en-us/javascript/api/word/word.shadowstyle) | "Mixed" | "OuterShadow" | "InnerShadow"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the shape shadow type.

```typescript
type?: Word.ShadowType | "Mixed" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9" | "Type10" | "Type11" | "Type12" | "Type13" | "Type14" | "Type15" | "Type16" | "Type17" | "Type18" | "Type19" | "Type20" | "Type21" | "Type22" | "Type23" | "Type24" | "Type25" | "Type26" | "Type27" | "Type28" | "Type29" | "Type30" | "Type31" | "Type32" | "Type33" | "Type34" | "Type35" | "Type36" | "Type37" | "Type38" | "Type39" | "Type40" | "Type41" | "Type42" | "Type43";
```

Property value: [Word.ShadowType](/en-us/javascript/api/word/word.shadowtype) | "Mixed" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9" | "Type10" | "Type11" | "Type12" | "Type13" | "Type14" | "Type15" | "Type16" | "Type17" | "Type18" | "Type19" | "Type20" | "Type21" | "Type22" | "Type23" | "Type24" | "Type25" | "Type26" | "Type27" | "Type28" | "Type29" | "Type30" | "Type31" | "Type32" | "Type33" | "Type34" | "Type35" | "Type36" | "Type37" | "Type38" | "Type39" | "Type40" | "Type41" | "Type42" | "Type43"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)