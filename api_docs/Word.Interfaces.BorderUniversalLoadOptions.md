# Word.Interfaces.BorderUniversalLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the BorderUniversal object, which manages borders for a range, paragraph, table, or frame.

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- artStyle: Specifies the graphical page-border design for the document.
- artWidth: Specifies the width (in points) of the graphical page border specified in the artStyle property.
- color: Specifies the color for the BorderUniversal object. You can provide the value in the '#RRGGBB' format.
- colorIndex: Specifies the color for the BorderUniversal or [Word.Font](/en-us/javascript/api/word/word.font) object.
- inside: Returns true if an inside border can be applied to the specified object.
- isVisible: Specifies whether the border is visible.
- lineStyle: Specifies the line style of the border.
- lineWidth: Specifies the line width of an object's border.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

### artStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the graphical page-border design for the document.

```typescript
artStyle?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### artWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in points) of the graphical page border specified in the artStyle property.

```typescript
artWidth?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### color

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the BorderUniversal object. You can provide the value in the '#RRGGBB' format.

```typescript
color?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### colorIndex

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the BorderUniversal or [Word.Font](/en-us/javascript/api/word/word.font) object.

```typescript
colorIndex?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### inside

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if an inside border can be applied to the specified object.

```typescript
inside?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the border is visible.

```typescript
isVisible?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line style of the border.

```typescript
lineStyle?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line width of an object's border.

```typescript
lineWidth?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)