# Word.Interfaces.BodyData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling body.toJSON().

## Properties

- contentControls: Gets the collection of rich text content control objects in the body.
- fields: Gets the collection of field objects in the body.
- font: Gets the text format of the body. Use this to get and set font name, size, color, and other properties.
- inlinePictures: Gets the collection of InlinePicture objects in the body. The collection doesn't include floating images.
- lists: Gets the collection of list objects in the body.
- paragraphs: Gets the collection of paragraph objects in the body.
- shapes: Gets the collection of shape objects in the body, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.
- style: Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBuiltIn: Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- tables: Gets the collection of table objects in the body.
- text: Gets the text of the body. Use the insertText method to insert text.
- type: Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Additional types 'Footnote', 'Endnote', and 'NoteItem' are supported in WordAPIOnline 1.1 and later.

## Property Details

### contentControls

Gets the collection of rich text content control objects in the body.

```typescript
contentControls?: Word.Interfaces.ContentControlData[];
```

Property Value:
[Word.Interfaces.ContentControlData](/en-us/javascript/api/word/word.interfaces.contentcontroldata)[]

Remarks:
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fields

Gets the collection of field objects in the body.

```typescript
fields?: Word.Interfaces.FieldData[];
```

Property Value:
[Word.Interfaces.FieldData](/en-us/javascript/api/word/word.interfaces.fielddata)[]

Remarks:
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### font

Gets the text format of the body. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontData;
```

Property Value:
[Word.Interfaces.FontData](/en-us/javascript/api/word/word.interfaces.fontdata)

Remarks:
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### inlinePictures

Gets the collection of InlinePicture objects in the body. The collection doesn't include floating images.

```typescript
inlinePictures?: Word.Interfaces.InlinePictureData[];
```

Property Value:
[Word.Interfaces.InlinePictureData](/en-us/javascript/api/word/word.interfaces.inlinepicturedata)[]

Remarks:
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lists

Gets the collection of list objects in the body.

```typescript
lists?: Word.Interfaces.ListData[];
```

Property Value:
[Word.Interfaces.ListData](/en-us/javascript/api/word/word.interfaces.listdata)[]

Remarks:
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### paragraphs

Gets the collection of paragraph objects in the body.

```typescript
paragraphs?: Word.Interfaces.ParagraphData[];
```

Property Value:
[Word.Interfaces.ParagraphData](/en-us/javascript/api/word/word.interfaces.paragraphdata)[]

Remarks:
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

Important: Paragraphs in tables aren't returned for requirement sets 1.1 and 1.2. From requirement set 1.3, paragraphs in tables are also returned.

---

### shapes

Gets the collection of shape objects in the body, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

```typescript
shapes?: Word.Interfaces.ShapeData[];
```

Property Value:
[Word.Interfaces.ShapeData](/en-us/javascript/api/word/word.interfaces.shapedata)[]

Remarks:
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### style

Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: string;
```

Property Value:
string

Remarks:
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### styleBuiltIn

Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

Property Value:
[Word.BuiltInStyleName](/en-us/javascript/api/word/word.builtinstylename) | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"

Remarks:
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tables

Gets the collection of table objects in the body.

```typescript
tables?: Word.Interfaces.TableData[];
```

Property Value:
[Word.Interfaces.TableData](/en-us/javascript/api/word/word.interfaces.tabledata)[]

Remarks:
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### text

Gets the text of the body. Use the insertText method to insert text.

```typescript
text?: string;
```

Property Value:
string

Remarks:
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Additional types 'Footnote', 'Endnote', and 'NoteItem' are supported in WordAPIOnline 1.1 and later.

```typescript
type?: Word.BodyType | "Unknown" | "MainDoc" | "Section" | "Header" | "Footer" | "TableCell" | "Footnote" | "Endnote" | "NoteItem" | "Shape";
```

Property Value:
[Word.BodyType](/en-us/javascript/api/word/word.bodytype) | "Unknown" | "MainDoc" | "Section" | "Header" | "Footer" | "TableCell" | "Footnote" | "Endnote" | "NoteItem" | "Shape"

Remarks:
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)