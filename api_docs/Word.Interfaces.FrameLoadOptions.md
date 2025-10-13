# Word.Interfaces.FrameLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a frame. The `Frame` object is a member of the [Word.FrameCollection](/en-us/javascript/api/word/word.framecollection) object.

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

## Properties
- [$all](#all): Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- [height](#height): Specifies the height (in points) of the frame.
- [heightRule](#heightrule): Specifies a `FrameSizeRule` value that represents the rule for determining the height of the frame.
- [horizontalDistanceFromText](#horizontaldistancefromtext): Specifies the horizontal distance between the frame and the surrounding text, in points.
- [horizontalPosition](#horizontalposition): Specifies the horizontal distance between the edge of the frame and the item specified by the `relativeHorizontalPosition` property.
- [lockAnchor](#lockanchor): Specifies if the frame is locked.
- [range](#range): Returns a `Range` object that represents the portion of the document that's contained within the frame.
- [relativeHorizontalPosition](#relativehorizontalposition): Specifies the relative horizontal position of the frame.
- [relativeVerticalPosition](#relativeverticalposition): Specifies the relative vertical position of the frame.
- [shading](#shading): Returns a `ShadingUniversal` object that refers to the shading formatting for the frame.
- [textWrap](#textwrap): Specifies if document text wraps around the frame.
- [verticalDistanceFromText](#verticaldistancefromtext): Specifies the vertical distance (in points) between the frame and the surrounding text.
- [verticalPosition](#verticalposition): Specifies the vertical distance between the edge of the frame and the item specified by the `relativeVerticalPosition` property.
- [width](#width): Specifies the width (in points) of the frame.
- [widthRule](#widthrule): Specifies the rule used to determine the width of the frame.

## Property Details

### $all
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

### height
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the height (in points) of the frame.

```typescript
height?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### heightRule
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `FrameSizeRule` value that represents the rule for determining the height of the frame.

```typescript
heightRule?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### horizontalDistanceFromText
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal distance between the frame and the surrounding text, in points.

```typescript
horizontalDistanceFromText?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### horizontalPosition
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal distance between the edge of the frame and the item specified by the `relativeHorizontalPosition` property.

```typescript
horizontalPosition?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### lockAnchor
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the frame is locked.

```typescript
lockAnchor?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### range
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Range` object that represents the portion of the document that's contained within the frame.

```typescript
range?: Word.Interfaces.RangeLoadOptions;
```

Property Value
- [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### relativeHorizontalPosition
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the relative horizontal position of the frame.

```typescript
relativeHorizontalPosition?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### relativeVerticalPosition
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the relative vertical position of the frame.

```typescript
relativeVerticalPosition?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### shading
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ShadingUniversal` object that refers to the shading formatting for the frame.

```typescript
shading?: Word.Interfaces.ShadingUniversalLoadOptions;
```

Property Value
- [Word.Interfaces.ShadingUniversalLoadOptions](/en-us/javascript/api/word/word.interfaces.shadinguniversalloadoptions)

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### textWrap
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if document text wraps around the frame.

```typescript
textWrap?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### verticalDistanceFromText
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical distance (in points) between the frame and the surrounding text.

```typescript
verticalDistanceFromText?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### verticalPosition
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical distance between the edge of the frame and the item specified by the `relativeVerticalPosition` property.

```typescript
verticalPosition?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### width
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in points) of the frame.

```typescript
width?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### widthRule
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rule used to determine the width of the frame.

```typescript
widthRule?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]