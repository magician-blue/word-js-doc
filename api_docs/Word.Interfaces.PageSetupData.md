# Word.Interfaces.PageSetupData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface describing the data returned by calling pageSetup.toJSON().

## Properties

- bookFoldPrinting — Specifies whether Microsoft Word prints the document as a booklet.
- bookFoldPrintingSheets — Specifies the number of pages for each booklet.
- bookFoldReversePrinting — Specifies if Microsoft Word reverses the printing order for book fold printing of bidirectional or Asian language documents.
- bottomMargin — Specifies the distance (in points) between the bottom edge of the page and the bottom boundary of the body text.
- charsLine — Specifies the number of characters per line in the document grid.
- differentFirstPageHeaderFooter — Specifies whether the first page has a different header and footer.
- footerDistance — Specifies the distance between the footer and the bottom of the page in points.
- gutter — Specifies the amount (in points) of extra margin space added to each page in a document or section for binding.
- gutterPosition — Specifies on which side the gutter appears in a document.
- gutterStyle — Specifies whether Microsoft Word uses gutters for the current document based on a right-to-left language or a left-to-right language.
- headerDistance — Specifies the distance between the header and the top of the page in points.
- layoutMode — Specifies the layout mode for the current document.
- leftMargin — Specifies the distance (in points) between the left edge of the page and the left boundary of the body text.
- lineNumbering — Specifies a LineNumbering object that represents the line numbers for the PageSetup object.
- linesPage — Specifies the number of lines per page in the document grid.
- mirrorMargins — Specifies if the inside and outside margins of facing pages are the same width.
- oddAndEvenPagesHeaderFooter — Specifies whether odd and even pages have different headers and footers.
- orientation — Specifies the orientation of the page.
- pageHeight — Specifies the page height in points.
- pageWidth — Specifies the page width in points.
- paperSize — Specifies the paper size of the page.
- rightMargin — Specifies the distance (in points) between the right edge of the page and the right boundary of the body text.
- sectionDirection — Specifies the reading order and alignment for the specified sections.
- sectionStart — Specifies the type of section break for the specified object.
- showGrid — Specifies whether to show the grid.
- suppressEndnotes — Specifies if endnotes are printed at the end of the next section that doesn't suppress endnotes.
- textColumns — Gets a TextColumnCollection object that represents the set of text columns for the PageSetup object.
- topMargin — Specifies the top margin of the page in points.
- twoPagesOnOne — Specifies whether to print two pages per sheet.
- verticalAlignment — Specifies the vertical alignment of text on each page in a document or section.

## Property Details

### bookFoldPrinting

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word prints the document as a booklet.

```typescript
bookFoldPrinting?: boolean;
```

Property Value: boolean

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### bookFoldPrintingSheets

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number of pages for each booklet.

```typescript
bookFoldPrintingSheets?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### bookFoldReversePrinting

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if Microsoft Word reverses the printing order for book fold printing of bidirectional or Asian language documents.

```typescript
bookFoldReversePrinting?: boolean;
```

Property Value: boolean

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### bottomMargin

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the bottom edge of the page and the bottom boundary of the body text.

```typescript
bottomMargin?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### charsLine

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number of characters per line in the document grid.

```typescript
charsLine?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### differentFirstPageHeaderFooter

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the first page has a different header and footer.

```typescript
differentFirstPageHeaderFooter?: boolean;
```

Property Value: boolean

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### footerDistance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance between the footer and the bottom of the page in points.

```typescript
footerDistance?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### gutter

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the amount (in points) of extra margin space added to each page in a document or section for binding.

```typescript
gutter?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### gutterPosition

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies on which side the gutter appears in a document.

```typescript
gutterPosition?: Word.GutterPosition | "Left" | "Right" | "Top";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.gutterposition | "Left" | "Right" | "Top"

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### gutterStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word uses gutters for the current document based on a right-to-left language or a left-to-right language.

```typescript
gutterStyle?: Word.GutterStyle | "Bidirectional" | "Latin";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.gutterstyle | "Bidirectional" | "Latin"

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### headerDistance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance between the header and the top of the page in points.

```typescript
headerDistance?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### layoutMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the layout mode for the current document.

```typescript
layoutMode?: Word.LayoutMode | "Default" | "Grid" | "LineGrid" | "Genko";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.layoutmode | "Default" | "Grid" | "LineGrid" | "Genko"

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### leftMargin

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the left edge of the page and the left boundary of the body text.

```typescript
leftMargin?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### lineNumbering

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `LineNumbering` object that represents the line numbers for the `PageSetup` object.

```typescript
lineNumbering?: Word.Interfaces.LineNumberingData;
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.linenumberingdata

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### linesPage

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number of lines per page in the document grid.

```typescript
linesPage?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### mirrorMargins

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the inside and outside margins of facing pages are the same width.

```typescript
mirrorMargins?: boolean;
```

Property Value: boolean

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### oddAndEvenPagesHeaderFooter

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether odd and even pages have different headers and footers.

```typescript
oddAndEvenPagesHeaderFooter?: boolean;
```

Property Value: boolean

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### orientation

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the orientation of the page.

```typescript
orientation?: Word.PageOrientation | "Portrait" | "Landscape";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.pageorientation | "Portrait" | "Landscape"

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### pageHeight

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page height in points.

```typescript
pageHeight?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### pageWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page width in points.

```typescript
pageWidth?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### paperSize

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the paper size of the page.

```typescript
paperSize?: Word.PaperSize | "Size10x14" | "Size11x17" | "Letter" | "LetterSmall" | "Legal" | "Executive" | "A3" | "A4" | "A4Small" | "A5" | "B4" | "B5" | "CSheet" | "DSheet" | "ESheet" | "FanfoldLegalGerman" | "FanfoldStdGerman" | "FanfoldUS" | "Folio" | "Ledger" | "Note" | "Quarto" | "Statement" | "Tabloid" | "Envelope9" | "Envelope10" | "Envelope11" | "Envelope12" | "Envelope14" | "EnvelopeB4" | "EnvelopeB5" | "EnvelopeB6" | "EnvelopeC3" | "EnvelopeC4" | "EnvelopeC5" | "EnvelopeC6" | "EnvelopeC65" | "EnvelopeDL" | "EnvelopeItaly" | "EnvelopeMonarch" | "EnvelopePersonal" | "Custom";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.papersize | "Size10x14" | "Size11x17" | "Letter" | "LetterSmall" | "Legal" | "Executive" | "A3" | "A4" | "A4Small" | "A5" | "B4" | "B5" | "CSheet" | "DSheet" | "ESheet" | "FanfoldLegalGerman" | "FanfoldStdGerman" | "FanfoldUS" | "Folio" | "Ledger" | "Note" | "Quarto" | "Statement" | "Tabloid" | "Envelope9" | "Envelope10" | "Envelope11" | "Envelope12" | "Envelope14" | "EnvelopeB4" | "EnvelopeB5" | "EnvelopeB6" | "EnvelopeC3" | "EnvelopeC4" | "EnvelopeC5" | "EnvelopeC6" | "EnvelopeC65" | "EnvelopeDL" | "EnvelopeItaly" | "EnvelopeMonarch" | "EnvelopePersonal" | "Custom"

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### rightMargin

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the right edge of the page and the right boundary of the body text.

```typescript
rightMargin?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### sectionDirection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the reading order and alignment for the specified sections.

```typescript
sectionDirection?: Word.SectionDirection | "RightToLeft" | "LeftToRight";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.sectiondirection | "RightToLeft" | "LeftToRight"

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### sectionStart

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the type of section break for the specified object.

```typescript
sectionStart?: Word.SectionStart | "Continuous" | "NewColumn" | "NewPage" | "EvenPage" | "OddPage";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.sectionstart | "Continuous" | "NewColumn" | "NewPage" | "EvenPage" | "OddPage"

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### showGrid

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to show the grid.

```typescript
showGrid?: boolean;
```

Property Value: boolean

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### suppressEndnotes

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if endnotes are printed at the end of the next section that doesn't suppress endnotes.

```typescript
suppressEndnotes?: boolean;
```

Property Value: boolean

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### textColumns

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `TextColumnCollection` object that represents the set of text columns for the `PageSetup` object.

```typescript
textColumns?: Word.Interfaces.TextColumnData[];
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.textcolumndata[]

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### topMargin

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the top margin of the page in points.

```typescript
topMargin?: number;
```

Property Value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### twoPagesOnOne

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to print two pages per sheet.

```typescript
twoPagesOnOne?: boolean;
```

Property Value: boolean

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))

---

### verticalAlignment

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical alignment of text on each page in a document or section.

```typescript
verticalAlignment?: Word.PageSetupVerticalAlignment | "Top" | "Center" | "Justify" | "Bottom";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.pagesetupverticalalignment | "Top" | "Center" | "Justify" | "Bottom"

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApi BETA (PREVIEW ONLY))