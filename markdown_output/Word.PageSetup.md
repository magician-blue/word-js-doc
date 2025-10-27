# Word.PageSetup

**Package:** `https://learn.microsoft.com/en-us/javascript/api/word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `officeextension.clientobject`

## Description

Represents the page setup settings for a Word document or section.

## Properties

### bookFoldPrinting

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether Microsoft Word prints the document as a booklet.

#### Examples

**Example**: Enable book fold printing to format the document as a booklet

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Enable book fold printing
    pageSetup.bookFoldPrinting = true;
    
    await context.sync();
    
    console.log("Book fold printing enabled");
});
```

---

### bookFoldPrintingSheets

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the number of pages for each booklet.

#### Examples

**Example**: Set up a document for booklet printing with 4 pages per booklet sheet

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    
    // Set the number of pages for each booklet
    pageSetup.bookFoldPrintingSheets = 4;
    
    await context.sync();
    
    console.log("Booklet printing configured with 4 pages per sheet");
});
```

---

### bookFoldReversePrinting

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if Microsoft Word reverses the printing order for book fold printing of bidirectional or Asian language documents.

#### Examples

**Example**: Enable reverse printing order for book fold printing to properly print a bidirectional language document

```typescript
await Word.run(async (context) => {
    // Get the page setup of the active document
    const pageSetup = context.document.body.pageSetup;
    
    // Enable reverse printing for book fold printing
    pageSetup.bookFoldReversePrinting = true;
    
    await context.sync();
    
    console.log("Book fold reverse printing has been enabled.");
});
```

---

### bottomMargin

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the distance (in points) between the bottom edge of the page and the bottom boundary of the body text.

#### Examples

**Example**: Set the bottom margin of the document to 72 points (1 inch)

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    pageSetup.bottomMargin = 72;
    
    await context.sync();
});
```

---

### charsLine

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the number of characters per line in the document grid.

#### Examples

**Example**: Set the document grid to display 40 characters per line

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    pageSetup.charsLine = 40;
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the page setup's request context to verify the connection to the Word host application and log its diagnostic information

```typescript
await Word.run(async (context) => {
    // Get the page setup of the active document
    const pageSetup = context.document.body.pageSetup;
    
    // Access the request context associated with the page setup object
    const requestContext = pageSetup.context;
    
    // Use the context to verify connection and get diagnostic info
    console.log("Request context is connected:", requestContext !== null);
    console.log("Context debug info:", requestContext.debugInfo);
    
    // The context connects the add-in process to Word's process
    // and is used internally for all API operations
    await context.sync();
});
```

---

### differentFirstPageHeaderFooter

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the first page has a different header and footer.

#### Examples

**Example**: Configure the document to use a different header and footer on the first page

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document body
    const pageSetup = context.document.body.pageSetup;
    
    // Enable different header/footer for the first page
    pageSetup.differentFirstPageHeaderFooter = true;
    
    await context.sync();
    
    console.log("First page will now have different headers and footers");
});
```

---

### footerDistance

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the distance between the footer and the bottom of the page in points.

#### Examples

**Example**: Set the footer distance to 36 points (0.5 inches) from the bottom of the page

```typescript
await Word.run(async (context) => {
    // Get the page setup of the active document
    const pageSetup = context.document.body.pageSetup;
    
    // Set the footer distance to 36 points (0.5 inches)
    pageSetup.footerDistance = 36;
    
    await context.sync();
    
    console.log("Footer distance set to 36 points");
});
```

---

### gutter

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the amount (in points) of extra margin space added to each page in a document or section for binding.

#### Examples

**Example**: Set a 36-point gutter margin for binding on the left side of the document

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Set the gutter to 36 points for binding
    pageSetup.gutter = 36;
    
    await context.sync();
    
    console.log("Gutter margin set to 36 points for binding.");
});
```

---

### gutterPosition

**Type:** `Word.GutterPosition | "Left" | "Right" | "Top"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies on which side the gutter appears in a document.

#### Examples

**Example**: Set the gutter position to the top of the page for binding at the top edge of the document

```typescript
await Word.run(async (context) => {
    // Get the page setup of the active document
    const pageSetup = context.document.body.pageSetup;
    
    // Set the gutter position to top
    pageSetup.gutterPosition = Word.GutterPosition.top;
    
    await context.sync();
    
    console.log("Gutter position set to top");
});
```

---

### gutterStyle

**Type:** `Word.GutterStyle | "Bidirectional" | "Latin"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether Microsoft Word uses gutters for the current document based on a right-to-left language or a left-to-right language.

#### Examples

**Example**: Set the gutter style to right-to-left (bidirectional) for the current document to accommodate Arabic or Hebrew text layout

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Set the gutter style to bidirectional (right-to-left)
    pageSetup.gutterStyle = Word.GutterStyle.bidirectional;
    
    await context.sync();
    
    console.log("Gutter style set to bidirectional");
});
```

---

### headerDistance

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the distance between the header and the top of the page in points.

#### Examples

**Example**: Set the header distance to 36 points (0.5 inches) from the top of the page

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Set the header distance to 36 points (0.5 inches)
    pageSetup.headerDistance = 36;
    
    await context.sync();
    
    console.log("Header distance set to 36 points");
});
```

---

### layoutMode

**Type:** `Word.LayoutMode | "Default" | "Grid" | "LineGrid" | "Genko"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the layout mode for the current document.

#### Examples

**Example**: Set the document's layout mode to Grid for precise character and line positioning

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Set the layout mode to Grid
    pageSetup.layoutMode = Word.LayoutMode.grid;
    
    await context.sync();
    
    console.log("Layout mode set to Grid");
});
```

---

### leftMargin

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the distance (in points) between the left edge of the page and the left boundary of the body text.

#### Examples

**Example**: Set the left margin of the document to 1 inch (72 points)

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    pageSetup.leftMargin = 72;
    
    await context.sync();
});
```

---

### lineNumbering

**Type:** `Word.LineNumbering`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a LineNumbering object that represents the line numbers for the PageSetup object.

#### Examples

**Example**: Configure line numbering to restart at 1 on each page with numbers appearing every 5 lines

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Configure line numbering settings
    const lineNumbering = pageSetup.lineNumbering;
    lineNumbering.restartMode = Word.LineNumberRestartMode.restartPage;
    lineNumbering.countBy = 5;
    lineNumbering.startingNumber = 1;
    
    await context.sync();
    
    console.log("Line numbering configured successfully");
});
```

---

### linesPage

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the number of lines per page in the document grid.

#### Examples

**Example**: Set the document to display 44 lines per page in the document grid

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Set the number of lines per page to 44
    pageSetup.linesPage = 44;
    
    await context.sync();
});
```

---

### mirrorMargins

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the inside and outside margins of facing pages are the same width.

#### Examples

**Example**: Enable mirror margins for a document so that inside and outside margins are swapped on facing pages for book-style printing

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Enable mirror margins for facing pages
    pageSetup.mirrorMargins = true;
    
    await context.sync();
    
    console.log("Mirror margins enabled for the document");
});
```

---

### oddAndEvenPagesHeaderFooter

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether odd and even pages have different headers and footers.

#### Examples

**Example**: Configure a document to use different headers and footers for odd and even pages

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    
    // Enable different headers/footers for odd and even pages
    pageSetup.oddAndEvenPagesHeaderFooter = true;
    
    await context.sync();
    
    console.log("Odd and even pages will now have different headers and footers");
});
```

---

### orientation

**Type:** `Word.PageOrientation | "Portrait" | "Landscape"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the orientation of the page.

#### Examples

**Example**: Change the page orientation of the active document to landscape mode

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document body
    const pageSetup = context.document.body.pageSetup;
    
    // Set the orientation to landscape
    pageSetup.orientation = Word.PageOrientation.landscape;
    
    await context.sync();
});
```

---

### pageHeight

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the page height in points.

#### Examples

**Example**: Set the page height to 792 points (11 inches) for the active document

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    pageSetup.pageHeight = 792;
    
    await context.sync();
});
```

---

### pageWidth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the page width in points.

#### Examples

**Example**: Set the page width to 8.5 inches (612 points) for the active document

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    
    // Set page width to 8.5 inches (8.5 * 72 = 612 points)
    pageSetup.pageWidth = 612;
    
    await context.sync();
});
```

---

### paperSize

**Type:** `Word.PaperSize | "Size10x14" | "Size11x17" | "Letter" | "LetterSmall" | "Legal" | "Executive" | "A3" | "A4" | "A4Small" | "A5" | "B4" | "B5" | "CSheet" | "DSheet" | "ESheet" | "FanfoldLegalGerman" | "FanfoldStdGerman" | "FanfoldUS" | "Folio" | "Ledger" | "Note" | "Quarto" | "Statement" | "Tabloid" | "Envelope9" | "Envelope10" | "Envelope11" | "Envelope12" | "Envelope14" | "EnvelopeB4" | "EnvelopeB5" | "EnvelopeB6" | "EnvelopeC3" | "EnvelopeC4" | "EnvelopeC5" | "EnvelopeC6" | "EnvelopeC65" | "EnvelopeDL" | "EnvelopeItaly" | "EnvelopeMonarch" | "EnvelopePersonal" | "Custom"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the paper size of the page.

#### Examples

**Example**: Set the document's paper size to A4 format

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document body
    const pageSetup = context.document.body.pageSetup;
    
    // Set the paper size to A4
    pageSetup.paperSize = "A4";
    
    await context.sync();
});
```

---

### rightMargin

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the distance (in points) between the right edge of the page and the right boundary of the body text.

#### Examples

**Example**: Set the right margin of the document to 1 inch (72 points)

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Set the right margin to 72 points (1 inch)
    pageSetup.rightMargin = 72;
    
    await context.sync();
});
```

---

### sectionDirection

**Type:** `Word.SectionDirection | "RightToLeft" | "LeftToRight"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the reading order and alignment for the specified sections.

#### Examples

**Example**: Set the section direction to right-to-left for the first section of the document to support languages like Arabic or Hebrew

```typescript
await Word.run(async (context) => {
    const firstSection = context.document.sections.getFirst();
    firstSection.pageSetup.sectionDirection = Word.SectionDirection.rightToLeft;
    
    await context.sync();
});
```

---

### sectionStart

**Type:** `Word.SectionStart | "Continuous" | "NewColumn" | "NewPage" | "EvenPage" | "OddPage"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the type of section break for the specified object.

#### Examples

**Example**: Set the section to start on a new page

```typescript
await Word.run(async (context) => {
    // Get the first section of the document
    const section = context.document.sections.getFirst();
    
    // Load the page setup
    section.load("pageSetup");
    await context.sync();
    
    // Set the section to start on a new page
    section.pageSetup.sectionStart = Word.SectionStart.newPage;
    
    await context.sync();
});
```

---

### showGrid

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to show the grid.

#### Examples

**Example**: Show the document grid to help with alignment and layout

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Show the grid
    pageSetup.showGrid = true;
    
    await context.sync();
});
```

---

### suppressEndnotes

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if endnotes are printed at the end of the next section that doesn't suppress endnotes.

#### Examples

**Example**: Configure the current section to suppress endnotes so they print at the end of the next section that allows them

```typescript
await Word.run(async (context) => {
    // Get the first section of the document
    const section = context.document.sections.getFirst();
    
    // Access the page setup for this section
    const pageSetup = section.pageSetup;
    
    // Suppress endnotes in this section
    pageSetup.suppressEndnotes = true;
    
    await context.sync();
    
    console.log("Endnotes will be suppressed in this section and printed at the end of the next section that doesn't suppress them.");
});
```

---

### textColumns

**Type:** `Word.TextColumnCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a TextColumnCollection object that represents the set of text columns for the PageSetup object.

#### Examples

**Example**: Configure the document to use a two-column layout with a 0.5 inch spacing between columns

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Access the text columns collection
    const textColumns = pageSetup.textColumns;
    
    // Set the number of columns to 2
    textColumns.set({
        columnCount: 2,
        spacing: 36 // 0.5 inch in points (72 points = 1 inch)
    });
    
    await context.sync();
    
    console.log("Document configured with 2 columns");
});
```

---

### topMargin

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the top margin of the page in points.

#### Examples

**Example**: Set the top margin of the document to 1 inch (72 points)

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    pageSetup.topMargin = 72;
    
    await context.sync();
});
```

---

### twoPagesOnOne

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to print two pages per sheet.

#### Examples

**Example**: Configure the document to print two pages on one sheet of paper

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document
    const pageSetup = context.document.body.pageSetup;
    
    // Enable printing two pages on one sheet
    pageSetup.twoPagesOnOne = true;
    
    await context.sync();
    
    console.log("Document configured to print two pages per sheet");
});
```

---

### verticalAlignment

**Type:** `Word.PageSetupVerticalAlignment | "Top" | "Center" | "Justify" | "Bottom"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical alignment of text on each page in a document or section.

#### Examples

**Example**: Set the vertical alignment of the document to center so that text is vertically centered on each page

```typescript
await Word.run(async (context) => {
    // Get the page setup of the first section
    const pageSetup = context.document.sections.getFirst().pageSetup;
    
    // Set vertical alignment to center
    pageSetup.verticalAlignment = Word.PageSetupVerticalAlignment.center;
    
    await context.sync();
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.PageSetupLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.PageSetup`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.PageSetup`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.PageSetup`

#### Examples

**Example**: Load and display the current page orientation and margins of the active document

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document body
    const pageSetup = context.document.body.pageSetup;
    
    // Load specific properties
    pageSetup.load("orientation, topMargin, bottomMargin, leftMargin, rightMargin");
    
    // Sync to read the loaded properties
    await context.sync();
    
    // Display the loaded properties
    console.log(`Orientation: ${pageSetup.orientation}`);
    console.log(`Top Margin: ${pageSetup.topMargin}`);
    console.log(`Bottom Margin: ${pageSetup.bottomMargin}`);
    console.log(`Left Margin: ${pageSetup.leftMargin}`);
    console.log(`Right Margin: ${pageSetup.rightMargin}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.PageSetupUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.PageSetup` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure page setup for the document by setting multiple properties including orientation, margins, and paper size at once

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    
    pageSetup.set({
        orientation: Word.PageOrientation.landscape,
        topMargin: 72,      // 1 inch in points
        bottomMargin: 72,
        leftMargin: 54,     // 0.75 inches
        rightMargin: 54,
        pageWidth: 792,     // 11 inches (letter size)
        pageHeight: 612     // 8.5 inches (letter size)
    });
    
    await context.sync();
    console.log("Page setup configured successfully");
});
```

---

### setAsTemplateDefault

**Kind:** `configure`

Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Set custom page margins (1 inch top/bottom, 1.5 inch left/right) as the default for all new documents based on the current template

```typescript
await Word.run(async (context) => {
    // Get the page setup for the document
    const pageSetup = context.document.body.pageSetup;
    
    // Configure custom margins
    pageSetup.topMargin = 72;      // 1 inch (72 points)
    pageSetup.bottomMargin = 72;   // 1 inch
    pageSetup.leftMargin = 108;    // 1.5 inches (108 points)
    pageSetup.rightMargin = 108;   // 1.5 inches
    
    // Set these settings as the default for the template
    pageSetup.setAsTemplateDefault();
    
    await context.sync();
    
    console.log("Page setup saved as template default");
});
```

---

### togglePortrait

**Kind:** `configure`

Switches between portrait and landscape page orientations for a document or section.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Toggle the page orientation of the current document from portrait to landscape (or vice versa)

```typescript
await Word.run(async (context) => {
    // Get the page setup of the document body
    const pageSetup = context.document.body.pageSetup;
    
    // Toggle between portrait and landscape orientation
    pageSetup.togglePortrait();
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.PageSetup object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PageSetupData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.PageSetupData`

#### Examples

**Example**: Retrieve page setup settings as a plain JavaScript object and log them to the console for debugging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the page setup of the first section
    const firstSection = context.document.sections.getFirst();
    const pageSetup = firstSection.pageSetup;
    
    // Load the properties we want to serialize
    pageSetup.load("orientation,pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin");
    
    await context.sync();
    
    // Convert the PageSetup object to a plain JavaScript object
    const pageSetupData = pageSetup.toJSON();
    
    // Now we can use the plain object for logging, storage, or comparison
    console.log("Page Setup Data:", pageSetupData);
    console.log("Page Width:", pageSetupData.pageWidth);
    console.log("Page Height:", pageSetupData.pageHeight);
    console.log("Orientation:", pageSetupData.orientation);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.PageSetup`

#### Examples

**Example**: Track a page setup object across multiple sync calls to modify its margins and orientation without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.body.pageSetup;
    pageSetup.load("topMargin,orientation");
    
    // Track the object to use it across multiple sync calls
    pageSetup.track();
    
    await context.sync();
    
    console.log(`Current top margin: ${pageSetup.topMargin}`);
    
    // Modify properties after sync - tracking prevents InvalidObjectPath error
    pageSetup.topMargin = 72; // 1 inch
    pageSetup.orientation = Word.PageOrientation.landscape;
    
    await context.sync();
    
    console.log("Page setup updated successfully");
    
    // Untrack when done to free up memory
    pageSetup.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.PageSetup`

#### Examples

**Example**: Apply page setup changes to a section and then untrack the PageSetup object to free memory after the modifications are complete.

```typescript
await Word.run(async (context) => {
    // Get the first section's page setup
    const pageSetup = context.document.sections.getFirst().pageSetup;
    
    // Track the object to work with it
    pageSetup.track();
    
    // Load and modify page setup properties
    pageSetup.load("orientation");
    await context.sync();
    
    // Make changes to page setup
    pageSetup.orientation = Word.PageOrientation.landscape;
    pageSetup.topMargin = 72; // 1 inch
    
    await context.sync();
    
    // Untrack the object to release memory
    pageSetup.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
