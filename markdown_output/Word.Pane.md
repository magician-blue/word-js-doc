# Word.Pane

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a window pane. The Pane object is a member of the pane collection. The pane collection includes all the window panes for a single window.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets the first paragraph of each page.
  console.log("Getting first paragraph of each page...");

  // Get the active window.
  const activeWindow: Word.Window = context.document.activeWindow;
  activeWindow.load();

  // Get the active pane.
  const activePane: Word.Pane = activeWindow.activePane;
  activePane.load();

  // Get all pages.
  const pages: Word.PageCollection = activePane.pages;
  pages.load();

  await context.sync();

  // Get page index and paragraphs of each page.
  const pagesIndexes = [];
  const pagesNumberOfParagraphs = [];
  const pagesFirstParagraphText = [];
  for (let i = 0; i < pages.items.length; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);

    const paragraphs = page.getRange().paragraphs;
    paragraphs.load('items/length');
    pagesNumberOfParagraphs.push(paragraphs);

    const firstParagraph = paragraphs.getFirst();
    firstParagraph.load('text');
    pagesFirstParagraphText.push(firstParagraph);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Page index: ${pagesIndexes[i].index}`);
    console.log(`Number of paragraphs: ${pagesNumberOfParagraphs[i].items.length}`);
    console.log("First paragraph's text:", pagesFirstParagraphText[i].text);
  }
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a window pane to verify the connection between the add-in and Word, then use it to load and log the pane's index property.

```typescript
await Word.run(async (context) => {
    // Get the active pane
    const activePane = context.document.getActiveWindow().panes.getFirst();
    
    // Access the request context from the pane object
    const paneContext = activePane.context;
    
    // Use the context to load properties
    activePane.load("index");
    
    await paneContext.sync();
    
    console.log(`Pane index: ${activePane.index}`);
    console.log("Request context successfully accessed from pane object");
});
```

---

### pages

**Type:** `Word.PageCollection`

**Since:** WordApiDesktop 1.2

Gets the collection of pages in the pane.

#### Examples

**Example**: Retrieve and display the page index, total number of paragraphs, and first paragraph text for each page in the active document pane.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets the first paragraph of each page.
  console.log("Getting first paragraph of each page...");

  // Get the active window.
  const activeWindow: Word.Window = context.document.activeWindow;
  activeWindow.load();

  // Get the active pane.
  const activePane: Word.Pane = activeWindow.activePane;
  activePane.load();

  // Get all pages.
  const pages: Word.PageCollection = activePane.pages;
  pages.load();

  await context.sync();

  // Get page index and paragraphs of each page.
  const pagesIndexes = [];
  const pagesNumberOfParagraphs = [];
  const pagesFirstParagraphText = [];
  for (let i = 0; i < pages.items.length; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);

    const paragraphs = page.getRange().paragraphs;
    paragraphs.load('items/length');
    pagesNumberOfParagraphs.push(paragraphs);

    const firstParagraph = paragraphs.getFirst();
    firstParagraph.load('text');
    pagesFirstParagraphText.push(firstParagraph);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Page index: ${pagesIndexes[i].index}`);
    console.log(`Number of paragraphs: ${pagesNumberOfParagraphs[i].items.length}`);
    console.log("First paragraph's text:", pagesFirstParagraphText[i].text);
  }
});
```

---

### pagesEnclosingViewport

**Type:** `Word.PageCollection`

**Since:** WordApiDesktop 1.2

Gets the PageCollection shown in the viewport of the pane. If a page is partially visible in the pane, the whole page is returned.

#### Examples

**Example**: Retrieve and log the count and index values of all pages that are currently visible within the active document window's viewport.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets the pages enclosing the viewport.

  // Get the active window.
  const activeWindow: Word.Window = context.document.activeWindow;
  activeWindow.load();

  // Get the active pane.
  const activePane: Word.Pane = activeWindow.activePane;
  activePane.load();

  // Get pages enclosing the viewport.
  const pages: Word.PageCollection = activePane.pagesEnclosingViewport;
  pages.load();

  await context.sync();

  // Log the number of pages.
  const pageCount = pages.items.length;
  console.log(`Number of pages enclosing the viewport: ${pageCount}`);

  // Log index info of these pages.
  const pagesIndexes = [];
  for (let i = 0; i < pageCount; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Page index: ${pagesIndexes[i].index}`);
  }
});
```

---

## Methods

### getNext

**Kind:** `read`

Gets the next pane in the window. Throws an ItemNotFound error if this pane is the last one.

#### Signature

**Returns:** `Word.Pane`

#### Examples

**Example**: Navigate through all window panes sequentially and display their index positions in the document body.

```typescript
await Word.run(async (context) => {
    const firstPane = context.document.application.activeWindow.panes.getFirst();
    firstPane.load("index");
    
    let currentPane = firstPane;
    let paneCount = 1;
    let paneInfo = "";
    
    // Get the first pane info
    await context.sync();
    paneInfo += `Pane ${paneCount}: Index ${currentPane.index}\n`;
    
    // Navigate through remaining panes using getNext()
    let hasNext = true;
    while (hasNext) {
        try {
            currentPane = currentPane.getNext();
            currentPane.load("index");
            await context.sync();
            
            paneCount++;
            paneInfo += `Pane ${paneCount}: Index ${currentPane.index}\n`;
        } catch (error) {
            // ItemNotFound error means we've reached the last pane
            hasNext = false;
        }
    }
    
    // Display results in document
    const body = context.document.body;
    body.insertParagraph(paneInfo, Word.InsertLocation.end);
    
    await context.sync();
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next pane. If this pane is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Pane`

#### Examples

**Example**: Check if there are multiple panes in the active window and display information about the current pane and the next pane if it exists.

```typescript
await Word.run(async (context) => {
    const firstPane = context.document.getActiveWindow().panes.getFirst();
    const nextPane = firstPane.getNextOrNullObject();
    
    firstPane.load("index");
    nextPane.load("isNullObject, index");
    
    await context.sync();
    
    console.log(`Current pane index: ${firstPane.index}`);
    
    if (nextPane.isNullObject) {
        console.log("This is the last pane in the window.");
    } else {
        console.log(`Next pane exists with index: ${nextPane.index}`);
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Pane`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Pane`

#### Examples

**Example**: Load and display the index property of the first pane in the active document's window

```typescript
await Word.run(async (context) => {
    // Get the first pane from the active document's window
    const pane = context.document.application.activeWindow.panes.getFirst();
    
    // Queue a command to load the index property
    pane.load("index");
    
    // Sync to execute the queued command
    await context.sync();
    
    // Now we can read the loaded property
    console.log(`Pane index: ${pane.index}`);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Pane object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PaneData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.PaneData`

#### Examples

**Example**: Serialize a window pane object to JSON format for logging or debugging purposes

```typescript
await Word.run(async (context) => {
    // Get the first pane of the active window
    const pane = context.document.application.windows.getActiveOrNullObject().panes.getFirst();
    
    // Load properties of the pane
    pane.load("index");
    
    await context.sync();
    
    // Convert the pane object to a plain JavaScript object
    const paneData = pane.toJSON();
    
    // Log the serialized data (useful for debugging)
    console.log("Pane data:", JSON.stringify(paneData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Pane`

#### Examples

**Example**: Track a window pane object across multiple sync calls to maintain its reference while monitoring and logging its view type property

```typescript
await Word.run(async (context) => {
    // Get the first pane of the active window
    const pane = context.document.application.activeWindow.panes.getFirst();
    
    // Track the pane object to use it across multiple sync calls
    pane.track();
    
    // Load the view type property
    pane.load("viewType");
    await context.sync();
    
    console.log("Initial view type:", pane.viewType);
    
    // Perform additional operations with the tracked pane
    // The pane reference remains valid across sync calls
    pane.load("viewType");
    await context.sync();
    
    console.log("View type after second sync:", pane.viewType);
    
    // Untrack when done to free up resources
    pane.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Pane`

#### Examples

**Example**: Access a window pane, perform operations with it, then untrack it to free memory after use

```typescript
await Word.run(async (context) => {
    // Get the first pane of the active window
    const pane = context.document.application.activeWindow.panes.getFirst();
    
    // Track the pane object for use
    pane.track();
    
    // Load properties to work with the pane
    pane.load("index");
    await context.sync();
    
    // Use the pane (e.g., log its index)
    console.log(`Pane index: ${pane.index}`);
    
    // Untrack the pane to release memory
    pane.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.pane
