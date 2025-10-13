# Word.Interfaces.FrameData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling frame.toJSON().

## Properties

- borders: Returns a BorderUniversalCollection object that represents all the borders for the frame.
- height: Specifies the height (in points) of the frame.
- heightRule: Specifies a FrameSizeRule value that represents the rule for determining the height of the frame.
- horizontalDistanceFromText: Specifies the horizontal distance between the frame and the surrounding text, in points.
- horizontalPosition: Specifies the horizontal distance between the edge of the frame and the item specified by the relativeHorizontalPosition property.
- lockAnchor: Specifies if the frame is locked.
- range: Returns a Range object that represents the portion of the document that's contained within the frame.
- relativeHorizontalPosition: Specifies the relative horizontal position of the frame.
- relativeVerticalPosition: Specifies the relative vertical position of the frame.
- shading: Returns a ShadingUniversal object that refers to the shading formatting for the frame.
- textWrap: Specifies if document text wraps around the frame.
- verticalDistanceFromText: Specifies the vertical distance (in points) between the frame and the surrounding text.
- verticalPosition: Specifies the vertical distance between the edge of the frame and the item specified by the relativeVerticalPosition property.
- width: Specifies the width (in points) of the frame.
- widthRule: Specifies the rule used to determine the width of the frame.

## Property Details

### borders

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders for the frame.

```typescript
borders?: Word.Interfaces.BorderUniversalData[];
```

- Property value: [Word.Interfaces.BorderUniversalData](/en-us/javascript/api/word/word.interfaces.borderuniversaldata)[]
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### height

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the height (in points) of the frame.

```typescript
height?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### heightRule

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a FrameSizeRule value that represents the rule for determining the height of the frame.

```typescript
heightRule?: Word.FrameSizeRule | "Auto" | "AtLeast" | "Exact";
```

- Property value: [Word.FrameSizeRule](/en-us/javascript/api/word/word.framesizerule) | "Auto" | "AtLeast" | "Exact"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### horizontalDistanceFromText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal distance between the frame and the surrounding text, in points.

```typescript
horizontalDistanceFromText?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### horizontalPosition

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal distance between the edge of the frame and the item specified by the relativeHorizontalPosition property.

```typescript
horizontalPosition?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lockAnchor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the frame is locked.

```typescript
lockAnchor?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Range object that represents the portion of the document that's contained within the frame.

```typescript
range?: Word.Interfaces.RangeData;
```

- Property value: [Word.Interfaces.RangeData](/en-us/javascript/api/word/word.interfaces.rangedata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### relativeHorizontalPosition

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the relative horizontal position of the frame.

```typescript
relativeHorizontalPosition?: Word.RelativeHorizontalPosition | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin";
```

- Property value: [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition) | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### relativeVerticalPosition

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the relative vertical position of the frame.

```typescript
relativeVerticalPosition?: Word.RelativeVerticalPosition | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin";
```

- Property value: [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition) | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the frame.

```typescript
shading?: Word.Interfaces.ShadingUniversalData;
```

- Property value: [Word.Interfaces.ShadingUniversalData](/en-us/javascript/api/word/word.interfaces.shadinguniversaldata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textWrap

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if document text wraps around the frame.

```typescript
textWrap?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalDistanceFromText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical distance (in points) between the frame and the surrounding text.

```typescript
verticalDistanceFromText?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalPosition

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical distance between the edge of the frame and the item specified by the relativeVerticalPosition property.

```typescript
verticalPosition?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in points) of the frame.

```typescript
width?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### widthRule

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rule used to determine the width of the frame.

```typescript
widthRule?: Word.FrameSizeRule | "Auto" | "AtLeast" | "Exact";
```

- Property value: [Word.FrameSizeRule](/en-us/javascript/api/word/word.framesizerule) | "Auto" | "AtLeast" | "Exact"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)