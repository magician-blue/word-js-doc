# Word.Interfaces.ShapeTextWrapData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling shapeTextWrap.toJSON().

## Properties

- bottomDistance  
  Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.
- leftDistance  
  Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.
- rightDistance  
  Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.
- side  
  Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.
- topDistance  
  Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.
- type  
  Specifies the text wrap type around the shape. See Word.ShapeTextWrapType for details.

## Property Details

### bottomDistance

Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.

```typescript
bottomDistance?: number;
```

Property Value
- number

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leftDistance

Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.

```typescript
leftDistance?: number;
```

Property Value
- number

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rightDistance

Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.

```typescript
rightDistance?: number;
```

Property Value
- number

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### side

Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.

```typescript
side?: Word.ShapeTextWrapSide | "None" | "Both" | "Left" | "Right" | "Largest";
```

Property Value
- [Word.ShapeTextWrapSide](/en-us/javascript/api/word/word.shapetextwrapside) | "None" | "Both" | "Left" | "Right" | "Largest"

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### topDistance

Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.

```typescript
topDistance?: number;
```

Property Value
- number

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Specifies the text wrap type around the shape. See Word.ShapeTextWrapType for details.

```typescript
type?: Word.ShapeTextWrapType | "Inline" | "Square" | "Tight" | "Through" | "TopBottom" | "Behind" | "Front";
```

Property Value
- [Word.ShapeTextWrapType](/en-us/javascript/api/word/word.shapetextwraptype) | "Inline" | "Square" | "Tight" | "Through" | "TopBottom" | "Behind" | "Front"

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)