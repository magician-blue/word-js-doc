# Word.Interfaces.CommentContentRangeUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the CommentContentRange object, for use in commentContentRange.set({ ... }).

## Properties

- bold  
  Specifies a value that indicates whether the comment text is bold.

- hyperlink  
  Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.

- italic  
  Specifies a value that indicates whether the comment text is italicized.

- strikeThrough  
  Specifies a value that indicates whether the comment text has a strikethrough.

- underline  
  Specifies a value that indicates the comment text's underline type. 'None' if the comment text isn't underlined.

## Property Details

### bold

Specifies a value that indicates whether the comment text is bold.

```typescript
bold?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### hyperlink

Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.

```typescript
hyperlink?: string;
```

Property Value  
string

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### italic

Specifies a value that indicates whether the comment text is italicized.

```typescript
italic?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### strikeThrough

Specifies a value that indicates whether the comment text has a strikethrough.

```typescript
strikeThrough?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### underline

Specifies a value that indicates the comment text's underline type. 'None' if the comment text isn't underlined.

```typescript
underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
```

Property Value  
[Word.UnderlineType](/en-us/javascript/api/word/word.underlinetype) | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble"

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)