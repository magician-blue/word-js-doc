# Word.Interfaces.ShapeTextWrapUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the ShapeTextWrap object, for use in shapeTextWrap.set({ ... }).

## Properties

- [bottomDistance](#bottomdistance)
  - Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.
- [leftDistance](#leftdistance)
  - Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.
- [rightDistance](#rightdistance)
  - Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.
- [side](#side)
  - Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.
- [topDistance](#topdistance)
  - Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.
- [type](#type)
  - Specifies the text wrap type around the shape. See [Word.ShapeTextWrapType](/en-us/javascript/api/word/word.shapetextwraptype) for details.

## Property Details

### bottomDistance

Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.

```typescript
bottomDistance?: number;
```

- Type: number  
- Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leftDistance

Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.

```typescript
leftDistance?: number;
```

- Type: number  
- Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rightDistance

Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.

```typescript
rightDistance?: number;
```

- Type: number  
- Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### side

Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.

```typescript
side?: Word.ShapeTextWrapSide | "None" | "Both" | "Left" | "Right" | "Largest";
```

- Type: [Word.ShapeTextWrapSide](/en-us/javascript/api/word/word.shapetextwrapside) | "None" | "Both" | "Left" | "Right" | "Largest"  
- Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### topDistance

Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.

```typescript
topDistance?: number;
```

- Type: number  
- Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Specifies the text wrap type around the shape. See [Word.ShapeTextWrapType](/en-us/javascript/api/word/word.shapetextwraptype) for details.

```typescript
type?: Word.ShapeTextWrapType | "Inline" | "Square" | "Tight" | "Through" | "TopBottom" | "Behind" | "Front";
```

- Type: [Word.ShapeTextWrapType](/en-us/javascript/api/word/word.shapetextwraptype) | "Inline" | "Square" | "Tight" | "Through" | "TopBottom" | "Behind" | "Front"  
- Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)