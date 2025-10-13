# Word.Interfaces.ParagraphFormatData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling paragraphFormat.toJSON().

## Properties

- alignment  
  Specifies the alignment for the specified paragraphs.
- firstLineIndent  
  Specifies the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
- keepTogether  
  Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.
- keepWithNext  
  Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.
- leftIndent  
  Specifies the left indent.
- lineSpacing  
  Specifies the line spacing (in points) for the specified paragraphs.
- lineUnitAfter  
  Specifies the amount of spacing (in gridlines) after the specified paragraphs.
- lineUnitBefore  
  Specifies the amount of spacing (in gridlines) before the specified paragraphs.
- mirrorIndents  
  Specifies whether left and right indents are the same width.
- outlineLevel  
  Specifies the outline level for the specified paragraphs.
- rightIndent  
  Specifies the right indent (in points) for the specified paragraphs.
- spaceAfter  
  Specifies the amount of spacing (in points) after the specified paragraph or text column.
- spaceBefore  
  Specifies the spacing (in points) before the specified paragraphs.
- widowControl  
  Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.

## Property Details

### alignment

Specifies the alignment for the specified paragraphs.

```typescript
alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

Property Value: [Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### firstLineIndent

Specifies the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

```typescript
firstLineIndent?: number;
```

Property Value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### keepTogether

Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.

```typescript
keepTogether?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### keepWithNext

Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.

```typescript
keepWithNext?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leftIndent

Specifies the left indent.

```typescript
leftIndent?: number;
```

Property Value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lineSpacing

Specifies the line spacing (in points) for the specified paragraphs.

```typescript
lineSpacing?: number;
```

Property Value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lineUnitAfter

Specifies the amount of spacing (in gridlines) after the specified paragraphs.

```typescript
lineUnitAfter?: number;
```

Property Value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lineUnitBefore

Specifies the amount of spacing (in gridlines) before the specified paragraphs.

```typescript
lineUnitBefore?: number;
```

Property Value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### mirrorIndents

Specifies whether left and right indents are the same width.

```typescript
mirrorIndents?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### outlineLevel

Specifies the outline level for the specified paragraphs.

```typescript
outlineLevel?: Word.OutlineLevel | "OutlineLevel1" | "OutlineLevel2" | "OutlineLevel3" | "OutlineLevel4" | "OutlineLevel5" | "OutlineLevel6" | "OutlineLevel7" | "OutlineLevel8" | "OutlineLevel9" | "OutlineLevelBodyText";
```

Property Value: [Word.OutlineLevel](/en-us/javascript/api/word/word.outlinelevel) | "OutlineLevel1" | "OutlineLevel2" | "OutlineLevel3" | "OutlineLevel4" | "OutlineLevel5" | "OutlineLevel6" | "OutlineLevel7" | "OutlineLevel8" | "OutlineLevel9" | "OutlineLevelBodyText"

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rightIndent

Specifies the right indent (in points) for the specified paragraphs.

```typescript
rightIndent?: number;
```

Property Value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### spaceAfter

Specifies the amount of spacing (in points) after the specified paragraph or text column.

```typescript
spaceAfter?: number;
```

Property Value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### spaceBefore

Specifies the spacing (in points) before the specified paragraphs.

```typescript
spaceBefore?: number;
```

Property Value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### widowControl

Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.

```typescript
widowControl?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)