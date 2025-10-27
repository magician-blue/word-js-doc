# Word.RangeCollection

**Package:** `word`

**API Set:** WordApi 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Range](/en-us/javascript/api/word/word.range) objects.

## Class Examples

**Example**: Does a basic text search and highlights matches in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/search.yaml

// Does a basic text search and highlights matches in the document.
await Word.run(async (context) => {
  const results : Word.RangeCollection = context.document.body.search("extend");
  results.load("length");

  await context.sync();

  // Let's traverse the search results and highlight matches.
  for (let i = 0; i < results.items.length; i++) {
    results.items[i].font.highlightColor = "yellow";
  }

  await context.sync();
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a RangeCollection to synchronize changes when highlighting all paragraphs in a document

```typescript
await Word.run(async (context) => {
    // Get all paragraph ranges in the document
    const paragraphs = context.document.body.paragraphs;
    const ranges = paragraphs.getRange();
    
    // Access the request context from the range collection
    const requestContext = ranges.context;
    
    // Use the context to load properties
    ranges.load("text");
    
    // Synchronize with the Office host using the context
    await requestContext.sync();
    
    // Apply highlighting to the ranges
    ranges.font.highlightColor = "yellow";
    
    await requestContext.sync();
    
    console.log("All paragraph ranges highlighted using their associated context");
});
```

---

### items

**Type:** `readonly Word.Range[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all ranges in a collection and highlight each range with a different color from a predefined list.

```typescript
await Word.run(async (context) => {
    // Search for all instances of the word "important"
    const searchResults = context.document.body.search("important", { matchCase: false });
    searchResults.load("items");
    
    await context.sync();
    
    // Access the items property to get the array of ranges
    const ranges = searchResults.items;
    
    // Define colors to cycle through
    const colors = ["yellow", "lightgreen", "lightblue", "pink"];
    
    // Highlight each range with a different color
    for (let i = 0; i < ranges.length; i++) {
        ranges[i].font.highlightColor = colors[i % colors.length];
    }
    
    await context.sync();
    
    console.log(`Highlighted ${ranges.length} ranges`);
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first range in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.Range`

#### Examples

**Example**: Highlight the first occurrence of the word "important" in the document by making it bold and red.

```typescript
await Word.run(async (context) => {
    // Search for all occurrences of "important"
    const searchResults = context.document.body.search("important", { matchCase: false });
    
    // Load the search results
    context.load(searchResults);
    await context.sync();
    
    // Get the first range from the collection
    const firstRange = searchResults.getFirst();
    
    // Format the first occurrence
    firstRange.font.bold = true;
    firstRange.font.color = "red";
    
    await context.sync();
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first range in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Range`

#### Examples

**Example**: Check if a document has any search results for a specific term and highlight the first occurrence if found

```typescript
await Word.run(async (context) => {
    // Search for the term "TypeScript" in the document
    const searchResults = context.document.body.search("TypeScript");
    
    // Get the first search result or null if none exist
    const firstResult = searchResults.getFirstOrNullObject();
    
    // Load the isNullObject property to check if a result was found
    firstResult.load("isNullObject");
    
    await context.sync();
    
    // Check if a result was found
    if (!firstResult.isNullObject) {
        // Highlight the first occurrence
        firstResult.font.highlightColor = "yellow";
        console.log("First occurrence of 'TypeScript' has been highlighted.");
    } else {
        console.log("No occurrences of 'TypeScript' found in the document.");
    }
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.RangeCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (required)
    Provides options for which properties of the object to load.

  **Returns:** `Word.RangeCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (required)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.RangeCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (required)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to

  **Returns:** `Word.RangeCollection`

#### Examples

**Example**: Load and display the text content of all ranges in a RangeCollection obtained from search results

```typescript
await Word.run(async (context) => {
    // Search for all instances of "TODO" in the document
    const searchResults = context.document.body.search("TODO");
    
    // Load the text property of all ranges in the collection
    searchResults.load("text");
    
    await context.sync();
    
    // Display the text of each range
    console.log(`Found ${searchResults.items.length} instances:`);
    searchResults.items.forEach((range, index) => {
        console.log(`Range ${index + 1}: ${range.text}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.RangeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.RangeCollectionData`

#### Examples

**Example**: Export range collection data to JSON format for logging or external storage by serializing all ranges found in the document's first paragraph.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get all ranges in the paragraph (e.g., by searching for a word)
    const searchResults = paragraph.search("the", { matchCase: false });
    
    // Load properties we want to serialize
    searchResults.load("text, font/name");
    
    await context.sync();
    
    // Convert the RangeCollection to a plain JavaScript object
    const rangeCollectionData = searchResults.toJSON();
    
    // Now we can use standard JSON operations
    const jsonString = JSON.stringify(rangeCollectionData, null, 2);
    console.log("Range Collection as JSON:", jsonString);
    console.log("Number of ranges:", rangeCollectionData.items.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track multiple paragraph ranges across sync calls to maintain references while modifying their formatting properties in separate batches.

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    context.load(paragraphs, "text");
    await context.sync();
    
    // Get ranges for paragraphs containing specific text
    const ranges = paragraphs.items
        .filter(p => p.text.includes("important"))
        .map(p => p.getRange());
    
    // Create a RangeCollection from the filtered ranges
    const rangeCollection = context.document.createRangeCollection(ranges);
    
    // Track the collection to use it across multiple sync calls
    rangeCollection.track();
    
    await context.sync();
    
    // First batch: highlight the ranges
    rangeCollection.items.forEach(range => {
        range.font.highlightColor = "yellow";
    });
    
    await context.sync();
    
    // Second batch: make them bold (object still valid due to tracking)
    rangeCollection.items.forEach(range => {
        range.font.bold = true;
    });
    
    await context.sync();
    
    // Clean up tracking when done
    rangeCollection.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Search for all instances of a specific word in the document, highlight them, then untrack the range collection to free memory after processing.

```typescript
await Word.run(async (context) => {
    // Search for all instances of "important"
    const searchResults = context.document.body.search("important", { matchCase: false });
    
    // Track the collection to work with it
    context.trackedObjects.add(searchResults);
    
    // Load and sync to get the results
    searchResults.load("items");
    await context.sync();
    
    // Highlight all found ranges
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
    
    // Untrack the collection to release memory
    searchResults.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.rangecollection
