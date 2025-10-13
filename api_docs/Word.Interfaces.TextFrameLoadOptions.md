# Word.Interfaces.TextFrameLoadOptions interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Represents the text frame of a shape object.

## Remarks

[ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- `$all`: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- `autoSizeSetting`: The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
- `bottomMargin`: Represents the bottom margin, in points, of the text frame.
- `hasText`: Specifies if the text frame contains text.
- `leftMargin`: Represents the left margin, in points, of the text frame.
- `noTextRotation`: Returns True if text in the text frame shouldn't rotate when the shape is rotated.
- `orientation`: Represents the angle to which the text is oriented for the text frame. See `Word.ShapeTextOrientation` for details.
- `rightMargin`: Represents the right margin, in points, of the text frame.
- `topMargin`: Represents the top margin, in points, of the text frame.
- `verticalAlignment`: Represents the vertical alignment of the text frame. See `Word.ShapeTextVerticalAlignment` for details.
- `wordWrap`: Determines whether lines break automatically to fit text inside the shape.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### autoSizeSetting

The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.

```typescript
autoSizeSetting?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bottomMargin

Represents the bottom margin, in points, of the text frame.

```typescript
bottomMargin?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hasText

Specifies if the text frame contains text.

```typescript
hasText?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leftMargin

Represents the left margin, in points, of the text frame.

```typescript
leftMargin?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### noTextRotation

Returns True if text in the text frame shouldn't rotate when the shape is rotated.

```typescript
noTextRotation?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### orientation

Represents the angle to which the text is oriented for the text frame. See `Word.ShapeTextOrientation` for details.

```typescript
orientation?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rightMargin

Represents the right margin, in points, of the text frame.

```typescript
rightMargin?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### topMargin

Represents the top margin, in points, of the text frame.

```typescript
topMargin?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### verticalAlignment

Represents the vertical alignment of the text frame. See `Word.ShapeTextVerticalAlignment` for details.

```typescript
verticalAlignment?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### wordWrap

Determines whether lines break automatically to fit text inside the shape.

```typescript
wordWrap?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)