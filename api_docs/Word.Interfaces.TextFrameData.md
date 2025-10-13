# Word.Interfaces.TextFrameData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `textFrame.toJSON()`.

## Properties

- [autoSizeSetting](#autosizesetting): The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
- [bottomMargin](#bottommargin): Represents the bottom margin, in points, of the text frame.
- [hasText](#hastext): Specifies if the text frame contains text.
- [leftMargin](#leftmargin): Represents the left margin, in points, of the text frame.
- [noTextRotation](#notextrotation): Returns True if text in the text frame shouldn't rotate when the shape is rotated.
- [orientation](#orientation): Represents the angle to which the text is oriented for the text frame. See `Word.ShapeTextOrientation` for details.
- [rightMargin](#rightmargin): Represents the right margin, in points, of the text frame.
- [topMargin](#topmargin): Represents the top margin, in points, of the text frame.
- [verticalAlignment](#verticalalignment): Represents the vertical alignment of the text frame. See `Word.ShapeTextVerticalAlignment` for details.
- [wordWrap](#wordwrap): Determines whether lines break automatically to fit text inside the shape.

## Property Details

### autoSizeSetting

The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.

```typescript
autoSizeSetting?: Word.ShapeAutoSize | "None" | "TextToFitShape" | "ShapeToFitText" | "Mixed";
```

Property Value: [Word.ShapeAutoSize](/en-us/javascript/api/word/word.shapeautosize) | "None" | "TextToFitShape" | "ShapeToFitText" | "Mixed"

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bottomMargin

Represents the bottom margin, in points, of the text frame.

```typescript
bottomMargin?: number;
```

Property Value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### hasText

Specifies if the text frame contains text.

```typescript
hasText?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leftMargin

Represents the left margin, in points, of the text frame.

```typescript
leftMargin?: number;
```

Property Value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### noTextRotation

Returns True if text in the text frame shouldn't rotate when the shape is rotated.

```typescript
noTextRotation?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### orientation

Represents the angle to which the text is oriented for the text frame. See `Word.ShapeTextOrientation` for details.

```typescript
orientation?: Word.ShapeTextOrientation | "None" | "Horizontal" | "EastAsianVertical" | "Vertical270" | "Vertical" | "EastAsianHorizontalRotated" | "Mixed";
```

Property Value: [Word.ShapeTextOrientation](/en-us/javascript/api/word/word.shapetextorientation) | "None" | "Horizontal" | "EastAsianVertical" | "Vertical270" | "Vertical" | "EastAsianHorizontalRotated" | "Mixed"

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rightMargin

Represents the right margin, in points, of the text frame.

```typescript
rightMargin?: number;
```

Property Value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### topMargin

Represents the top margin, in points, of the text frame.

```typescript
topMargin?: number;
```

Property Value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalAlignment

Represents the vertical alignment of the text frame. See `Word.ShapeTextVerticalAlignment` for details.

```typescript
verticalAlignment?: Word.ShapeTextVerticalAlignment | "Top" | "Middle" | "Bottom";
```

Property Value: [Word.ShapeTextVerticalAlignment](/en-us/javascript/api/word/word.shapetextverticalalignment) | "Top" | "Middle" | "Bottom"

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### wordWrap

Determines whether lines break automatically to fit text inside the shape.

```typescript
wordWrap?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)