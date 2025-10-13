# Word.Interfaces.ParagraphFormatLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a style of paragraph in a document.

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- alignment: Specifies the alignment for the specified paragraphs.
- firstLineIndent: Specifies the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
- keepTogether: Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.
- keepWithNext: Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.
- leftIndent: Specifies the left indent.
- lineSpacing: Specifies the line spacing (in points) for the specified paragraphs.
- lineUnitAfter: Specifies the amount of spacing (in gridlines) after the specified paragraphs.
- lineUnitBefore: Specifies the amount of spacing (in gridlines) before the specified paragraphs.
- mirrorIndents: Specifies whether left and right indents are the same width.
- outlineLevel: Specifies the outline level for the specified paragraphs.
- rightIndent: Specifies the right indent (in points) for the specified paragraphs.
- spaceAfter: Specifies the amount of spacing (in points) after the specified paragraph or text column.
- spaceBefore: Specifies the spacing (in points) before the specified paragraphs.
- widowControl: Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

### alignment

Specifies the alignment for the specified paragraphs.

```typescript
alignment?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### firstLineIndent

Specifies the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

```typescript
firstLineIndent?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### keepTogether

Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.

```typescript
keepTogether?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### keepWithNext

Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.

```typescript
keepWithNext?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leftIndent

Specifies the left indent.

```typescript
leftIndent?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineSpacing

Specifies the line spacing (in points) for the specified paragraphs.

```typescript
lineSpacing?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineUnitAfter

Specifies the amount of spacing (in gridlines) after the specified paragraphs.

```typescript
lineUnitAfter?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineUnitBefore

Specifies the amount of spacing (in gridlines) before the specified paragraphs.

```typescript
lineUnitBefore?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### mirrorIndents

Specifies whether left and right indents are the same width.

```typescript
mirrorIndents?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### outlineLevel

Specifies the outline level for the specified paragraphs.

```typescript
outlineLevel?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rightIndent

Specifies the right indent (in points) for the specified paragraphs.

```typescript
rightIndent?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### spaceAfter

Specifies the amount of spacing (in points) after the specified paragraph or text column.

```typescript
spaceAfter?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### spaceBefore

Specifies the spacing (in points) before the specified paragraphs.

```typescript
spaceBefore?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### widowControl

Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.

```typescript
widowControl?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)