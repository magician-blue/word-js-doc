# Word.PageSetup class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the page setup settings for a Word document or section.

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- bookFoldPrinting — Specifies whether Microsoft Word prints the document as a booklet.
- bookFoldPrintingSheets — Specifies the number of pages for each booklet.
- bookFoldReversePrinting — Specifies if Microsoft Word reverses the printing order for book fold printing of bidirectional or Asian language documents.
- bottomMargin — Specifies the distance (in points) between the bottom edge of the page and the bottom boundary of the body text.
- charsLine — Specifies the number of characters per line in the document grid.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
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

## Methods
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- setAsTemplateDefault() — Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.
- togglePortrait() — Switches between portrait and landscape page orientations for a document or section.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.PageSetup object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PageSetupData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property details

### bookFoldPrinting
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word prints the document as a booklet.

```typescript
bookFoldPrinting: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bookFoldPrintingSheets
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number of pages for each booklet.

```typescript
bookFoldPrintingSheets: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bookFoldReversePrinting
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if Microsoft Word reverses the printing order for book fold printing of bidirectional or Asian language documents.

```typescript
bookFoldReversePrinting: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bottomMargin
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the bottom edge of the page and the bottom boundary of the body text.

```typescript
bottomMargin: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### charsLine
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number of characters per line in the document grid.

```typescript
charsLine: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### differentFirstPageHeaderFooter
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the first page has a different header and footer.

```typescript
differentFirstPageHeaderFooter: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### footerDistance
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance between the footer and the bottom of the page in points.

```typescript
footerDistance: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gutter
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the amount (in points) of extra margin space added to each page in a document or section for binding.

```typescript
gutter: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gutterPosition
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies on which side the gutter appears in a document.

```typescript
gutterPosition: Word.GutterPosition | "Left" | "Right" | "Top";
```

Property Value: [Word.GutterPosition](https://learn.microsoft.com/en-us/javascript/api/word/word.gutterposition) | "Left" | "Right" | "Top"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gutterStyle
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word uses gutters for the current document based on a right-to-left language or a left-to-right language.

```typescript
gutterStyle: Word.GutterStyle | "Bidirectional" | "Latin";
```

Property Value: [Word.GutterStyle](https://learn.microsoft.com/en-us/javascript/api/word/word.gutterstyle) | "Bidirectional" | "Latin"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### headerDistance
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance between the header and the top of the page in points.

```typescript
headerDistance: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### layoutMode
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the layout mode for the current document.

```typescript
layoutMode: Word.LayoutMode | "Default" | "Grid" | "LineGrid" | "Genko";
```

Property Value: [Word.LayoutMode](https://learn.microsoft.com/en-us/javascript/api/word/word.layoutmode) | "Default" | "Grid" | "LineGrid" | "Genko"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leftMargin
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the left edge of the page and the left boundary of the body text.

```typescript
leftMargin: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineNumbering
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LineNumbering object that represents the line numbers for the PageSetup object.

```typescript
lineNumbering: Word.LineNumbering;
```

Property Value: [Word.LineNumbering](https://learn.microsoft.com/en-us/javascript/api/word/word.linenumbering)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### linesPage
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number of lines per page in the document grid.

```typescript
linesPage: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### mirrorMargins
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the inside and outside margins of facing pages are the same width.

```typescript
mirrorMargins: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### oddAndEvenPagesHeaderFooter
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether odd and even pages have different headers and footers.

```typescript
oddAndEvenPagesHeaderFooter: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### orientation
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the orientation of the page.

```typescript
orientation: Word.PageOrientation | "Portrait" | "Landscape";
```

Property Value: [Word.PageOrientation](https://learn.microsoft.com/en-us/javascript/api/word/word.pageorientation) | "Portrait" | "Landscape"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pageHeight
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page height in points.

```typescript
pageHeight: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pageWidth
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page width in points.

```typescript
pageWidth: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### paperSize
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the paper size of the page.

```typescript
paperSize: Word.PaperSize | "Size10x14" | "Size11x17" | "Letter" | "LetterSmall" | "Legal" | "Executive" | "A3" | "A4" | "A4Small" | "A5" | "B4" | "B5" | "CSheet" | "DSheet" | "ESheet" | "FanfoldLegalGerman" | "FanfoldStdGerman" | "FanfoldUS" | "Folio" | "Ledger" | "Note" | "Quarto" | "Statement" | "Tabloid" | "Envelope9" | "Envelope10" | "Envelope11" | "Envelope12" | "Envelope14" | "EnvelopeB4" | "EnvelopeB5" | "EnvelopeB6" | "EnvelopeC3" | "EnvelopeC4" | "EnvelopeC5" | "EnvelopeC6" | "EnvelopeC65" | "EnvelopeDL" | "EnvelopeItaly" | "EnvelopeMonarch" | "EnvelopePersonal" | "Custom";
```

Property Value: [Word.PaperSize](https://learn.microsoft.com/en-us/javascript/api/word/word.papersize) | "Size10x14" | "Size11x17" | "Letter" | "LetterSmall" | "Legal" | "Executive" | "A3" | "A4" | "A4Small" | "A5" | "B4" | "B5" | "CSheet" | "DSheet" | "ESheet" | "FanfoldLegalGerman" | "FanfoldStdGerman" | "FanfoldUS" | "Folio" | "Ledger" | "Note" | "Quarto" | "Statement" | "Tabloid" | "Envelope9" | "Envelope10" | "Envelope11" | "Envelope12" | "Envelope14" | "EnvelopeB4" | "EnvelopeB5" | "EnvelopeB6" | "EnvelopeC3" | "EnvelopeC4" | "EnvelopeC5" | "EnvelopeC6" | "EnvelopeC65" | "EnvelopeDL" | "EnvelopeItaly" | "EnvelopeMonarch" | "EnvelopePersonal" | "Custom"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rightMargin
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the right edge of the page and the right boundary of the body text.

```typescript
rightMargin: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sectionDirection
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the reading order and alignment for the specified sections.

```typescript
sectionDirection: Word.SectionDirection | "RightToLeft" | "LeftToRight";
```

Property Value: [Word.SectionDirection](https://learn.microsoft.com/en-us/javascript/api/word/word.sectiondirection) | "RightToLeft" | "LeftToRight"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sectionStart
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the type of section break for the specified object.

```typescript
sectionStart: Word.SectionStart | "Continuous" | "NewColumn" | "NewPage" | "EvenPage" | "OddPage";
```

Property Value: [Word.SectionStart](https://learn.microsoft.com/en-us/javascript/api/word/word.sectionstart) | "Continuous" | "NewColumn" | "NewPage" | "EvenPage" | "OddPage"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### showGrid
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to show the grid.

```typescript
showGrid: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### suppressEndnotes
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if endnotes are printed at the end of the next section that doesn't suppress endnotes.

```typescript
suppressEndnotes: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textColumns
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a TextColumnCollection object that represents the set of text columns for the PageSetup object.

```typescript
readonly textColumns: Word.TextColumnCollection;
```

Property Value: [Word.TextColumnCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.textcolumncollection)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### topMargin
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the top margin of the page in points.

```typescript
topMargin: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### twoPagesOnOne
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to print two pages per sheet.

```typescript
twoPagesOnOne: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalAlignment
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical alignment of text on each page in a document or section.

```typescript
verticalAlignment: Word.PageSetupVerticalAlignment | "Top" | "Center" | "Justify" | "Bottom";
```

Property Value: [Word.PageSetupVerticalAlignment](https://learn.microsoft.com/en-us/javascript/api/word/word.pagesetupverticalalignment) | "Top" | "Center" | "Justify" | "Bottom"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method details

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.PageSetupLoadOptions): Word.PageSetup;
```

Parameters:
- options: [Word.Interfaces.PageSetupLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.pagesetuploadoptions) — Provides options for which properties of the object to load.

Returns: [Word.PageSetup](https://learn.microsoft.com/en-us/javascript/api/word/word.pagesetup)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.PageSetup;
```

Parameters:
- propertyNames: string | string[] — A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.PageSetup](https://learn.microsoft.com/en-us/javascript/api/word/word.pagesetup)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.PageSetup;
```

Parameters:
- propertyNamesAndPaths: { select?: string; expand?: string; } — propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.PageSetup](https://learn.microsoft.com/en-us/javascript/api/word/word.pagesetup)

### set(properties, options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.PageSetupUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.PageSetupUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.pagesetupupdatedata) — A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions) — Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.PageSetup): void;
```

Parameters:
- properties: [Word.PageSetup](https://learn.microsoft.com/en-us/javascript/api/word/word.pagesetup)

Returns: void

### setAsTemplateDefault()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.

```typescript
setAsTemplateDefault(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### togglePortrait()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Switches between portrait and landscape page orientations for a document or section.

```typescript
togglePortrait(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.PageSetup object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PageSetupData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.PageSetupData;
```

Returns: [Word.Interfaces.PageSetupData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.pagesetupdata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.PageSetup;
```

Returns: [Word.PageSetup](https://learn.microsoft.com/en-us/javascript/api/word/word.pagesetup)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.PageSetup;
```

Returns: [Word.PageSetup](https://learn.microsoft.com/en-us/javascript/api/word/word.pagesetup)