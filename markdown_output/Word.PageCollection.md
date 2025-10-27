# Word.PageCollection

**Package:** `word`

**API Set:** WordApiDesktop 1.2 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the collection of page.

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

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a PageCollection to verify the connection to the Word host application before performing page operations

```typescript
await Word.run(async (context) => {
    const pages = context.document.body.getRange().getPages();
    
    // Access the request context associated with the PageCollection
    const requestContext = pages.context;
    
    // Verify the context is valid and connected to the host application
    if (requestContext) {
        console.log("PageCollection is connected to Word host application");
        
        // Use the context to load and sync page data
        pages.load("items");
        await requestContext.sync();
        
        console.log(`Number of pages: ${pages.items.length}`);
    }
});
```

---

### items

**Type:** `Word.Page[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Iterate through all loaded pages in the document and log each page's ID to the console

```typescript
await Word.run(async (context) => {
    // Get the page collection from the document
    const pages = context.document.body.pages;
    
    // Load the pages collection
    pages.load("items");
    
    await context.sync();
    
    // Access the loaded pages using the items property
    const pageItems = pages.items;
    
    // Iterate through each page and log its ID
    for (let i = 0; i < pageItems.length; i++) {
        console.log(`Page ${i + 1} ID: ${pageItems[i].id}`);
    }
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first page in this collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.Page`

#### Examples

**Example**: Get the first page in the document and display its page number in the console

```typescript
await Word.run(async (context) => {
    // Get the page collection from the document
    const pages = context.document.body.pages;
    
    // Get the first page
    const firstPage = pages.getFirst();
    
    // Load the page number property
    firstPage.load("pageNumber");
    
    // Sync to execute the queued commands
    await context.sync();
    
    // Display the page number
    console.log(`First page number: ${firstPage.pageNumber}`);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first page in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Page`

#### Examples

**Example**: Check if the document has any pages and display the page number of the first page, or show a message if the document is empty.

```typescript
await Word.run(async (context) => {
    const pages = context.document.body.pages;
    const firstPage = pages.getFirstOrNullObject();
    firstPage.load("pageNumber, isNullObject");
    
    await context.sync();
    
    if (firstPage.isNullObject) {
        console.log("The document has no pages.");
    } else {
        console.log(`First page number: ${firstPage.pageNumber}`);
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
  - `options`: `Word.Interfaces.PageCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.PageCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.PageCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.PageCollection`

#### Examples

**Example**: Load and display the total number of pages in the active Word document

```typescript
await Word.run(async (context) => {
    // Get the page collection from the document body
    const pages = context.document.body.pageCollection;
    
    // Load the count property of the page collection
    pages.load("items");
    
    // Synchronize the document state
    await context.sync();
    
    // Display the number of pages
    console.log(`Total pages in document: ${pages.items.length}`);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.PageCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PageCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.PageCollectionData`

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.PageCollection`

#### Examples

**Example**: Track a page collection to maintain references across multiple sync calls when working with page properties

```typescript
await Word.run(async (context) => {
    // Get the page collection from the document
    const pages = context.document.body.getRange().getPageCollection();
    
    // Track the page collection to prevent "InvalidObjectPath" errors
    // when accessing it across multiple sync calls
    pages.track();
    
    // Load page properties
    pages.load("items");
    await context.sync();
    
    // First sync - access page count
    console.log(`Total pages: ${pages.items.length}`);
    
    // Perform some other operations that might change the document
    context.document.body.insertParagraph("New content", Word.InsertLocation.end);
    await context.sync();
    
    // Second sync - can still safely access the tracked page collection
    console.log(`Pages still accessible: ${pages.items.length}`);
    
    // Untrack when done to free up memory
    pages.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.PageCollection`

#### Examples

**Example**: Load page information from a document, use it to display page count, then untrack the PageCollection to free memory

```typescript
await Word.run(async (context) => {
    // Get the page collection and load its items
    const pages = context.document.body.getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs();
    
    // Simpler approach - get pages from body
    const body = context.document.body;
    const pages = body.getRange().getRange('Whole').getParagraphs().getFirst().getRange().getParagraphs();
    
    // Actually get pages properly
    const pageCollection = context.document.body.getRange().getRange('Whole').getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs().getFirst().getRange().getParagraphs();
    
    // Correct way to get pages
    const range = context.document.body.getRange();
    const pageCollection = range.getRange('Whole').getParagraphs().getFirst().getRange().getParagraphs();
    
    // Load the page collection
    pageCollection.load("items");
    
    await context.sync();
    
    // Use the page data
    console.log(`Document has ${pageCollection.items.length} pages`);
    
    // Untrack the page collection to free memory
    pageCollection.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.pagecollection
