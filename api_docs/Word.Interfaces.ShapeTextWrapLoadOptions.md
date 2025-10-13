# Word.Interfaces.ShapeTextWrapLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Represents all the properties for wrapping text around a shape.

## Remarks

[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

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
  Specifies the text wrap type around the shape. See `Word.ShapeTextWrapType` for details.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

#### Property Value
- boolean

---

### bottomDistance

Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.

```typescript
bottomDistance?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leftDistance

Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.

```typescript
leftDistance?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rightDistance

Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.

```typescript
rightDistance?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### side

Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.

```typescript
side?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### topDistance

Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.

```typescript
topDistance?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Specifies the text wrap type around the shape. See `Word.ShapeTextWrapType` for details.

```typescript
type?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)