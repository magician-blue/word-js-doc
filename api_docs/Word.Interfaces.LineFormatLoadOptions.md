# Word.Interfaces.LineFormatLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents line and arrowhead formatting. For a line, the LineFormat object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all
  - Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- backgroundColor
  - Gets a ColorFormat object that represents the background color for a patterned line.
- beginArrowheadLength
  - Specifies the length of the arrowhead at the beginning of the line.
- beginArrowheadStyle
  - Specifies the style of the arrowhead at the beginning of the line.
- beginArrowheadWidth
  - Specifies the width of the arrowhead at the beginning of the line.
- dashStyle
  - Specifies the dash style for the line.
- endArrowheadLength
  - Specifies the length of the arrowhead at the end of the line.
- endArrowheadStyle
  - Specifies the style of the arrowhead at the end of the line.
- endArrowheadWidth
  - Specifies the width of the arrowhead at the end of the line.
- foregroundColor
  - Gets a ColorFormat object that represents the foreground color for the line.
- insetPen
  - Specifies if to draw lines inside a shape.
- isVisible
  - Specifies if the object, or the formatting applied to it, is visible.
- pattern
  - Specifies the pattern applied to the line.
- style
  - Specifies the line format style.
- transparency
  - Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).
- weight
  - Specifies the thickness of the line in points.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```ts
$all?: boolean;
```

Property Value
- boolean

### backgroundColor

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `ColorFormat` object that represents the background color for a patterned line.

```ts
backgroundColor?: Word.Interfaces.ColorFormatLoadOptions;
```

Property Value
- [Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### beginArrowheadLength

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the arrowhead at the beginning of the line.

```ts
beginArrowheadLength?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### beginArrowheadStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the style of the arrowhead at the beginning of the line.

```ts
beginArrowheadStyle?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### beginArrowheadWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the arrowhead at the beginning of the line.

```ts
beginArrowheadWidth?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### dashStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the dash style for the line.

```ts
dashStyle?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### endArrowheadLength

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the arrowhead at the end of the line.

```ts
endArrowheadLength?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### endArrowheadStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the style of the arrowhead at the end of the line.

```ts
endArrowheadStyle?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### endArrowheadWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the arrowhead at the end of the line.

```ts
endArrowheadWidth?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### foregroundColor

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `ColorFormat` object that represents the foreground color for the line.

```ts
foregroundColor?: Word.Interfaces.ColorFormatLoadOptions;
```

Property Value
- [Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### insetPen

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if to draw lines inside a shape.

```ts
insetPen?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the object, or the formatting applied to it, is visible.

```ts
isVisible?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pattern

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the pattern applied to the line.

```ts
pattern?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### style

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line format style.

```ts
style?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).

```ts
transparency?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### weight

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the thickness of the line in points.

```ts
weight?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)