# Word.Interfaces.ParagraphUpdateData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface for updating data on the Paragraph object, for use in paragraph.set({ ... }).

## Properties

- alignment  
  Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.

- firstLineIndent  
  Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

- font  
  Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.

- leftIndent  
  Specifies the left indent value, in points, for the paragraph.

- lineSpacing  
  Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.

- lineUnitAfter  
  Specifies the amount of spacing, in grid lines, after the paragraph.

- lineUnitBefore  
  Specifies the amount of spacing, in grid lines, before the paragraph.

- listItem  
  Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.

- listItemOrNullObject  
  Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- outlineLevel  
  Specifies the outline level for the paragraph.

- rightIndent  
  Specifies the right indent value, in points, for the paragraph.

- shading  
  Returns a ShadingUniversal object that refers to the shading formatting for the paragraph.

- spaceAfter  
  Specifies the spacing, in points, after the paragraph.

- spaceBefore  
  Specifies the spacing, in points, before the paragraph.

- style  
  Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

- styleBuiltIn  
  Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

## Property Details

### alignment

Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.

```typescript
alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

Property Value  
[Word.Alignment](https://learn.microsoft.com/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks  
[API set: WordApi 1.1]

---

### firstLineIndent

Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

```typescript
firstLineIndent?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### font

Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontUpdateData;
```

Property Value  
[Word.Interfaces.FontUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.fontupdatedata)

Remarks  
[API set: WordApi 1.1]

---

### leftIndent

Specifies the left indent value, in points, for the paragraph.

```typescript
leftIndent?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### lineSpacing

Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.

```typescript
lineSpacing?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### lineUnitAfter

Specifies the amount of spacing, in grid lines, after the paragraph.

```typescript
lineUnitAfter?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### lineUnitBefore

Specifies the amount of spacing, in grid lines, before the paragraph.

```typescript
lineUnitBefore?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### listItem

Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.

```typescript
listItem?: Word.Interfaces.ListItemUpdateData;
```

Property Value  
[Word.Interfaces.ListItemUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.listitemupdatedata)

Remarks  
[API set: WordApi 1.3]

---

### listItemOrNullObject

Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
listItemOrNullObject?: Word.Interfaces.ListItemUpdateData;
```

Property Value  
[Word.Interfaces.ListItemUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.listitemupdatedata)

Remarks  
[API set: WordApi 1.3]

---

### outlineLevel

Specifies the outline level for the paragraph.

```typescript
outlineLevel?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### rightIndent

Specifies the right indent value, in points, for the paragraph.

```typescript
rightIndent?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the paragraph.

```typescript
shading?: Word.Interfaces.ShadingUniversalUpdateData;
```

Property Value  
[Word.Interfaces.ShadingUniversalUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shadinguniversalupdatedata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### spaceAfter

Specifies the spacing, in points, after the paragraph.

```typescript
spaceAfter?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### spaceBefore

Specifies the spacing, in points, before the paragraph.

```typescript
spaceBefore?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### style

Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: string;
```

Property Value  
string

Remarks  
[API set: WordApi 1.1]

---

### styleBuiltIn

Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

Property Value  
[Word.BuiltInStyleName](https://learn.microsoft.com/en-us/javascript/api/word/word.builtinstylename) | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"

Remarks  
[API set: WordApi 1.3]