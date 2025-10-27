# Word.Page

**Package:** `https://learn.microsoft.com/en-us/javascript/api/word`

**API Set:** WordApiDesktop 1.2

**Extends:** `https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject`

## Description

Represents a page in the document. Page objects manage the page layout and content.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets pages of the selection.
  const pages: Word.PageCollection = context.document.getSelection().pages;
  pages.load();
  await context.sync();

  // Log info for pages included in selection.
  console.log(pages);
  const pagesIndexes = [];
  const pagesText = [];
  for (let i = 0; i < pages.items.length; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);

    const range = page.getRange();
    range.load('text');
    pagesText.push(range);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Index info for page ${i + 1} in the selection: ${pagesIndexes[i].index}`);
    console.log("Text of that page in the selection:", pagesText[i].text);
  }
});
```

## Properties

### breaks

**Type:** `None`

Gets a BreakCollection object that represents the breaks on the page.

#### Examples

**Example**: Get all page breaks on the first page and display their count and types in the console.

```typescript
await Word.run(async (context) => {
    const firstPage = context.document.body.sections.getFirst().getFirstPage();
    const breaks = firstPage.breaks;
    
    breaks.load("items/type");
    await context.sync();
    
    console.log(`Total breaks on page: ${breaks.items.length}`);
    breaks.items.forEach((pageBreak, index) => {
        console.log(`Break ${index + 1}: ${pageBreak.type}`);
    });
});
```

---

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the page's request context to verify the connection between the add-in and Word before performing page operations

```typescript
await Word.run(async (context) => {
    // Get the first page in the document
    const page = context.document.body.getRange().getPage();
    
    // Access the request context associated with the page object
    const pageContext = page.context;
    
    // Verify the context is valid by using it to load page properties
    page.load("isLastPage");
    
    await pageContext.sync();
    
    console.log("Page context is active and connected to Word");
    console.log(`Is last page: ${page.isLastPage}`);
});
```

---

### height

**Type:** `None`

Gets the height, in points, of the paper defined in the Page Setup dialog box.

#### Examples

**Example**: Display the current page height in points to the user in a message dialog

```typescript
await Word.run(async (context) => {
    // Get the first page of the document
    const page = context.document.body.pages.getFirst();
    
    // Load the height property
    page.load("height");
    
    await context.sync();
    
    // Display the page height
    console.log(`Page height: ${page.height} points`);
    // Or show in a dialog: OfficeExtension.UI.displayDialogAsync(`Page height: ${page.height} points`);
});
```

---

### index

**Type:** `None`

Gets the index of the page. The page index is 1-based and independent of the user's custom page numbering.

#### Examples

**Example**: Display the actual page index of the first page in the document to verify it starts at 1

```typescript
await Word.run(async (context) => {
    const firstPage = context.document.body.paragraphs.getFirst().getRange().getPage();
    firstPage.load("index");
    
    await context.sync();
    
    console.log(`The first page index is: ${firstPage.index}`);
    // Output: The first page index is: 1
});
```

---

### width

**Type:** `None`

Gets the width, in points, of the paper defined in the Page Setup dialog box.

#### Examples

**Example**: Display the current page width in points to verify the paper size settings

```typescript
await Word.run(async (context) => {
    // Get the first page of the document
    const page = context.document.body.sections.getFirst().getFirstPageOrNullObject();
    
    // Load the width property
    page.load("width");
    
    await context.sync();
    
    // Display the page width
    console.log(`Page width: ${page.width} points`);
});
```

---

## Methods

### getNext

**Kind:** `read`

Gets the next page in the pane. Throws an ItemNotFound error if this page is the last one.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get the next page after the first page and highlight all text on that second page with yellow background color.

```typescript
await Word.run(async (context) => {
    // Get the first page in the document
    const firstPage = context.document.body.pages.getFirst();
    
    // Get the next page (second page)
    const secondPage = firstPage.getNext();
    
    // Load the body of the second page
    secondPage.load("body");
    
    await context.sync();
    
    // Highlight all text on the second page with yellow
    secondPage.body.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next page. If this page is the last one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Check if the current page has a next page and display an alert indicating whether it's the last page in the document.

```typescript
await Word.run(async (context) => {
    // Get the first page in the document
    const firstPage = context.document.body.paragraphs.getFirst().getRange().getPage();
    
    // Get the next page or null object
    const nextPage = firstPage.getNextOrNullObject();
    
    // Load the isNullObject property
    nextPage.load("isNullObject");
    
    await context.sync();
    
    // Check if there is a next page
    if (nextPage.isNullObject) {
        console.log("This is the last page in the document.");
    } else {
        console.log("There is a next page after the current page.");
    }
});
```

---

### getRange

**Kind:** `read`

Gets the whole page, or the starting or ending point of the page, as a range.

#### Signature

**Parameters:**
- `rangeLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Highlight the entire content of the first page by getting its range and applying a yellow background color

```typescript
await Word.run(async (context) => {
    // Get the first page in the document
    const firstPage = context.document.body.sections.getFirst().getFirstPageHeader().page;
    
    // Alternative: Get pages from body
    const pages = context.document.body.getRange().getPages();
    const page = pages.items[0];
    
    // Get the range for the entire page
    const pageRange = page.getRange("Whole");
    
    // Apply yellow highlighting to the entire page content
    pageRange.font.highlightColor = "Yellow";
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `None` (required)

  **Returns:** `None`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Load and display the page number and dimensions of the first page in the document

```typescript
await Word.run(async (context) => {
    // Get the first page in the document
    const firstPage = context.document.body.pages.getFirst();
    
    // Load specific properties of the page
    firstPage.load("pageNumber, width, height");
    
    // Sync to execute the load command
    await context.sync();
    
    // Access the loaded properties
    console.log(`Page Number: ${firstPage.pageNumber}`);
    console.log(`Width: ${firstPage.width} points`);
    console.log(`Height: ${firstPage.height} points`);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Page object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PageData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Serialize a page object to JSON format to log or store its properties for debugging purposes

```typescript
await Word.run(async (context) => {
    // Get the first page in the document
    const page = context.document.body.pages.getFirst();
    
    // Load properties we want to serialize
    page.load("pageIndex");
    
    await context.sync();
    
    // Convert the page object to a plain JavaScript object
    const pageData = page.toJSON();
    
    // Now you can use the plain object (e.g., log it, store it, etc.)
    console.log("Page data:", JSON.stringify(pageData, null, 2));
    console.log("Page index:", pageData.pageIndex);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a page object to maintain its reference across multiple sync calls when checking and updating page properties

```typescript
await Word.run(async (context) => {
    // Get the first page
    const page = context.document.body.pages.getFirst();
    
    // Track the page object to use it across multiple sync calls
    page.track();
    
    // Load properties
    page.load("pageIndex");
    await context.sync();
    
    console.log(`Current page index: ${page.pageIndex}`);
    
    // Perform additional operations after sync
    // The tracked object remains valid
    page.load("height,width");
    await context.sync();
    
    console.log(`Page dimensions: ${page.width} x ${page.height}`);
    
    // Untrack when done to free up memory
    page.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is a shorthand for context.trackedObjects.remove(thisObject).

#### Signature

**Returns:** `None`

#### Examples

**Example**: Load page properties, use them for processing, then untrack the page object to free memory when done

```typescript
await Word.run(async (context) => {
    // Get the first page and track it
    const page = context.document.body.pages.getFirst();
    page.track();
    
    // Load properties for use
    page.load("width,height");
    await context.sync();
    
    // Use the page properties
    console.log(`Page dimensions: ${page.width} x ${page.height}`);
    
    // Untrack the page object to release memory
    page.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
