# Word.Interfaces.TextFrameUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `TextFrame` object, for use in `textFrame.set({ ... })`.

## Properties

- `autoSizeSetting`  
  The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
- `bottomMargin`  
  Represents the bottom margin, in points, of the text frame.
- `leftMargin`  
  Represents the left margin, in points, of the text frame.
- `noTextRotation`  
  Returns True if text in the text frame shouldn't rotate when the shape is rotated.
- `orientation`  
  Represents the angle to which the text is oriented for the text frame. See `Word.ShapeTextOrientation` for details.
- `rightMargin`  
  Represents the right margin, in points, of the text frame.
- `topMargin`  
  Represents the top margin, in points, of the text frame.
- `verticalAlignment`  
  Represents the vertical alignment of the text frame. See `Word.ShapeTextVerticalAlignment` for details.
- `wordWrap`  
  Determines whether lines break automatically to fit text inside the shape.

## Property Details

### autoSizeSetting

The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.

```typescript
autoSizeSetting?: Word.ShapeAutoSize | "None" | "TextToFitShape" | "ShapeToFitText" | "Mixed";
```

Property value: [Word.ShapeAutoSize](/en-us/javascript/api/word/word.shapeautosize) | "None" | "TextToFitShape" | "ShapeToFitText" | "Mixed"

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bottomMargin

Represents the bottom margin, in points, of the text frame.

```typescript
bottomMargin?: number;
```

Property value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leftMargin

Represents the left margin, in points, of the text frame.

```typescript
leftMargin?: number;
```

Property value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### noTextRotation

Returns True if text in the text frame shouldn't rotate when the shape is rotated.

```typescript
noTextRotation?: boolean;
```

Property value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### orientation

Represents the angle to which the text is oriented for the text frame. See `Word.ShapeTextOrientation` for details.

```typescript
orientation?: Word.ShapeTextOrientation | "None" | "Horizontal" | "EastAsianVertical" | "Vertical270" | "Vertical" | "EastAsianHorizontalRotated" | "Mixed";
```

Property value: [Word.ShapeTextOrientation](/en-us/javascript/api/word/word.shapetextorientation) | "None" | "Horizontal" | "EastAsianVertical" | "Vertical270" | "Vertical" | "EastAsianHorizontalRotated" | "Mixed"

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rightMargin

Represents the right margin, in points, of the text frame.

```typescript
rightMargin?: number;
```

Property value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### topMargin

Represents the top margin, in points, of the text frame.

```typescript
topMargin?: number;
```

Property value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalAlignment

Represents the vertical alignment of the text frame. See `Word.ShapeTextVerticalAlignment` for details.

```typescript
verticalAlignment?: Word.ShapeTextVerticalAlignment | "Top" | "Middle" | "Bottom";
```

Property value: [Word.ShapeTextVerticalAlignment](/en-us/javascript/api/word/word.shapetextverticalalignment) | "Top" | "Middle" | "Bottom"

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### wordWrap

Determines whether lines break automatically to fit text inside the shape.

```typescript
wordWrap?: boolean;
```

Property value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)