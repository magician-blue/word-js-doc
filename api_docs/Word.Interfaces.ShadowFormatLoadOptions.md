# Word.Interfaces.ShadowFormatLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the shadow formatting for a shape or text in Word.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [$all](#word-word-interfaces-shadowformatloadoptions-all-member)
  - Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- [blur](#word-word-interfaces-shadowformatloadoptions-blur-member)
  - Specifies the blur level for a shadow format as a value between 0.0 and 100.0.
- [foregroundColor](#word-word-interfaces-shadowformatloadoptions-foregroundcolor-member)
  - Returns a ColorFormat object that represents the foreground color for the fill, line, or shadow.
- [isVisible](#word-word-interfaces-shadowformatloadoptions-isvisible-member)
  - Specifies whether the object or the formatting applied to it is visible.
- [obscured](#word-word-interfaces-shadowformatloadoptions-obscured-member)
  - Specifies true if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill, false if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
- [offsetX](#word-word-interfaces-shadowformatloadoptions-offsetx-member)
  - Specifies the horizontal offset (in points) of the shadow from the shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
- [offsetY](#word-word-interfaces-shadowformatloadoptions-offsety-member)
  - Specifies the vertical offset (in points) of the shadow from the shape. A positive value offsets the shadow to the top of the shape; a negative value offsets it to the bottom.
- [rotateWithShape](#word-word-interfaces-shadowformatloadoptions-rotatewithshape-member)
  - Specifies whether to rotate the shadow when rotating the shape.
- [size](#word-word-interfaces-shadowformatloadoptions-size-member)
  - Specifies the width of the shadow.
- [style](#word-word-interfaces-shadowformatloadoptions-style-member)
  - Specifies the type of shadow formatting to apply to a shape.
- [transparency](#word-word-interfaces-shadowformatloadoptions-transparency-member)
  - Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).
- [type](#word-word-interfaces-shadowformatloadoptions-type-member)
  - Specifies the shape shadow type.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value
- boolean

### blur

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the blur level for a shadow format as a value between 0.0 and 100.0.

```typescript
blur?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### foregroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ColorFormat object that represents the foreground color for the fill, line, or shadow.

```typescript
foregroundColor?: Word.Interfaces.ColorFormatLoadOptions;
```

Property Value
- [Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the object or the formatting applied to it is visible.

```typescript
isVisible?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### obscured

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies true if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill, false if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.

```typescript
obscured?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### offsetX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal offset (in points) of the shadow from the shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.

```typescript
offsetX?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### offsetY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical offset (in points) of the shadow from the shape. A positive value offsets the shadow to the top of the shape; a negative value offsets it to the bottom.

```typescript
offsetY?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotateWithShape

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to rotate the shadow when rotating the shape.

```typescript
rotateWithShape?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### size

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the shadow.

```typescript
size?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### style

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the type of shadow formatting to apply to a shape.

```typescript
style?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the shape shadow type.

```typescript
type?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)